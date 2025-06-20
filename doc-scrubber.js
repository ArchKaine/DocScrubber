#!/usr/bin/env node

// Bootstrapper: This part of the script only uses built-in Node.js modules
// to check for dependencies before the main application logic is required.
const fs = require('fs').promises;
const path = require('path');
const readline = require('readline');
const { spawn } = require('child_process');

// List of required external dependencies
const DEPENDENCIES = [
    'yargs', 'mammoth', 'docx', 'jszip', 
    'fast-xml-parser', 'sharp', 'html-to-docx'
];

/**
 * Checks if all required dependencies are installed by trying to resolve them.
 * @returns {string[]} A list of missing package names.
 */
function findMissingDependencies() {
    const missing = [];
    console.log("Checking for required packages...");
    for (const dep of DEPENDENCIES) {
        try {
            require.resolve(dep);
        } catch (err) {
            console.log(` -> Missing package: ${dep}`);
            missing.push(dep);
        }
    }
    return missing;
}

/**
 * Installs the given list of packages using npm.
 * @param {string[]} packages - The npm packages to install.
 * @returns {Promise<boolean>} True if installation was successful, false otherwise.
 */
function installDependencies(packages) {
    return new Promise((resolve) => {
        console.log(`\nInstalling missing packages with npm: ${packages.join(' ')}...`);
        
        // Use spawn to run npm install, inheriting stdio to show progress to the user
        const npm = spawn('npm', ['install', ...packages], { stdio: 'inherit', shell: true });

        npm.on('close', (code) => {
            if (code === 0) {
                console.log("\nDependencies installed successfully.");
                resolve(true);
            } else {
                console.error(`\nError: npm install failed with code ${code}. Please try installing the packages manually.`);
                resolve(false);
            }
        });

        npm.on('error', (err) => {
            console.error('\nFailed to start npm. Please make sure Node.js and npm are installed correctly.');
            console.error(err);
            resolve(false);
        });
    });
}

/**
 * This function contains the entire application logic. It will only be called
 * after the dependency check has passed.
 */
async function startApp() {
    // --- All external requires are safely placed inside this function ---
    const yargs = require('yargs/yargs');
    const { hideBin } = require('yargs/helpers');
    const mammoth = require('mammoth');
    const { Document, Packer, Paragraph } = require('docx');
    const JSZip = require('jszip');
    const { XMLParser, XMLBuilder } = require('fast-xml-parser');
    const sharp = require('sharp');
    const HTMLtoDOCX = require('html-to-docx');
    
    // =================================================================================
    // HELPER & CORE LOGIC FUNCTIONS
    // =================================================================================

    async function askForConfirmation(question) {
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
        const answer = await new Promise(resolve => rl.question(question, resolve));
        rl.close();
        return answer.toLowerCase() === 'y' || answer.toLowerCase() === 'yes';
    }
    
    async function cleanDirectory(directoryPath, force) {
        console.log(`\n--- Scanning directory for generated files: ${directoryPath} ---`);
        try {
            const stats = await fs.stat(directoryPath);
            if (!stats.isDirectory()) { throw new Error("The path for 'clean' must be a directory."); }
            const allFiles = await fs.readdir(directoryPath);
            const generatedFileRegex = /(_optimized|_rebuilt|_merged|_\d{2,})\.docx$/;
            const filesToDelete = allFiles.filter(file => generatedFileRegex.test(file));
            if (filesToDelete.length === 0) { console.log("No generated files found to clean."); return; }
            console.log(`Found ${filesToDelete.length} generated file(s) to delete:`);
            filesToDelete.forEach(file => console.log(`  - ${file}`));
            let proceed = force;
            if (!force) {
                proceed = await askForConfirmation('\nAre you sure you want to permanently delete these files? (y/n) ');
                if (!proceed) console.log("\nClean operation aborted by user.");
            }
            if (proceed) {
                console.log("\nDeleting files...");
                let deleteCount = 0;
                for (const file of filesToDelete) {
                    try {
                        await fs.unlink(path.join(directoryPath, file));
                        console.log(`  - Deleted ${file}`);
                        deleteCount++;
                    } catch (err) { console.error(`  - Failed to delete ${file}: ${err.message}`); }
                }
                console.log(`\nSuccessfully deleted ${deleteCount} file(s).`);
            }
        } catch (err) {
            console.error(`\n❌ Error during clean operation: ${err.message}`);
        }
    }

    async function rebuildDocument(filePath, options) {
        console.log(`\n--- Rebuilding (with Formatting): ${path.basename(filePath)} ---`);
        const originalSize = (await fs.stat(filePath)).size;
        try {
            console.log("  -> Stage 1: Extracting clean content to HTML...");
            const { value: html } = await mammoth.convertToHtml({ path: filePath });
            if (!html) throw new Error("Could not extract any content from the document.");
            console.log("  -> Stage 2: Building new, clean document from HTML...");
            const docxBuffer = await HTMLtoDOCX(html, null, { table: { row: { cantSplit: true } }, footer: false, header: false });
            const newSize = docxBuffer.length;
            if (newSize >= originalSize && originalSize > 0) {
                console.log(`\nRebuild did not result in a smaller file. No file was written.`);
                return;
            }
            const outputPath = options.overwrite ? filePath : path.join(path.dirname(filePath), `${path.parse(filePath).name}_rebuilt.docx`);
            await fs.writeFile(outputPath, docxBuffer);
            const savings = originalSize - newSize;
            console.log(`\n  ✅ Successfully wrote file: ${path.basename(outputPath)}`);
            console.log(`     Original Size: ${(originalSize / 1024).toFixed(0)} KB, New Size: ${(newSize / 1024).toFixed(0)} KB, Saved: ${(savings / 1024).toFixed(0)} KB (${((savings / originalSize) * 100).toFixed(1)}%)`);
        } catch (err) {
            console.error(`\n❌ Failed to rebuild ${path.basename(filePath)}. Error:`, err.message);
        }
    }

    function getUsedStyleIds(docXml) { const usedIds = new Set(); function traverse(node) { if (typeof node !== 'object' || node === null) return; const pStyle = node['w:pPr']?.['w:pStyle']?.['@_w:val']; const rStyle = node['w:rPr']?.['w:rStyle']?.['@_w:val']; if (pStyle) usedIds.add(pStyle); if (rStyle) usedIds.add(rStyle); for (const key in node) { if (Array.isArray(node[key])) { node[key].forEach(traverse); } else { traverse(node[key]); } } } traverse(docXml); return usedIds; }
    
    async function optimizeDocument(filePath, options) {
        console.log(`\n--- Optimizing: ${path.basename(filePath)} ---`);
        const originalSize = (await fs.stat(filePath)).size;
        try {
            const fileBuffer = await fs.readFile(filePath);
            const zip = await JSZip.loadAsync(fileBuffer);
            if (options.removeEmbeddedFonts) {
                console.log("  -> Stage 1: Removing embedded fonts...");
                const fontFolder = zip.folder('word/fonts');
                if (fontFolder) {
                    let removedCount = 0;
                    fontFolder.forEach((relativePath, file) => { zip.remove(`word/fonts/${file.name}`); removedCount++; });
                    if (removedCount > 0) {
                        console.log(`    - Removed ${removedCount} font file(s).`);
                        const relsFile = zip.file('word/_rels/document.xml.rels');
                        if (relsFile) {
                            const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
                            const relsXml = parser.parse(await relsFile.async('string'));
                            if (relsXml.Relationships && relsXml.Relationships.Relationship) {
                                const originalRels = Array.isArray(relsXml.Relationships.Relationship) ? relsXml.Relationships.Relationship : [relsXml.Relationships.Relationship];
                                const fontRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
                                const filteredRels = originalRels.filter(r => r['@_Type'] !== fontRelationshipType);
                                if (filteredRels.length < originalRels.length) {
                                    relsXml.Relationships.Relationship = filteredRels.length > 0 ? filteredRels : undefined;
                                    const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", format: false });
                                    zip.file('word/_rels/document.xml.rels', builder.build(relsXml));
                                    console.log(`    - Removed font table relationships.`);
                                }
                            }
                        }
                        const settingsFile = zip.file('word/settings.xml');
                        if (settingsFile) {
                            const parser = new XMLParser({ ignoreAttributes: false, isArray: (tagName) => tagName === 'w:embedFont' });
                            const settingsXml = parser.parse(await settingsFile.async('string'));
                            if (settingsXml['w:settings'] && settingsXml['w:settings']['w:embedTrueTypeFonts']) {
                                delete settingsXml['w:settings']['w:embedTrueTypeFonts'];
                                console.log('    - Removed font embedding flag from document settings.');
                                const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", format: false });
                                zip.file('word/settings.xml', builder.build(settingsXml));
                            }
                        }
                    } else { console.log("    - No embedded fonts found to remove."); }
                } else { console.log("    - No embedded fonts found."); }
            }
            if (options.removeUnusedMedia) {
                console.log("  -> Stage 2: Searching for unused media...");
                const relsFile = zip.file('word/_rels/document.xml.rels');
                if (relsFile) {
                    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
                    const relsXml = parser.parse(await relsFile.async('string'));
                    let relationships = [];
                    if (relsXml.Relationships && relsXml.Relationships.Relationship) { relationships = Array.isArray(relsXml.Relationships.Relationship) ? relsXml.Relationships.Relationship : [relsXml.Relationships.Relationship]; }
                    const usedMedia = new Set();
                    relationships.forEach(rel => { if (rel['@_Type'] && rel['@_Type'].includes('image')) { usedMedia.add('word/' + rel['@_Target'].replace('../', '')); } });
                    const mediaFolder = zip.folder('word/media');
                    let removedCount = 0;
                    if (mediaFolder) {
                        const removalPromises = [];
                        mediaFolder.forEach((relativePath, file) => {
                            const fullPath = `word/media/${file.name}`;
                            if (!usedMedia.has(fullPath)) {
                                removalPromises.push(zip.remove(fullPath));
                                console.log(`    - Removing orphaned media: ${file.name}`);
                                removedCount++;
                            }
                        });
                        await Promise.all(removalPromises);
                    }
                    if (removedCount > 0) console.log(`    Removed ${removedCount} orphaned media file(s).`);
                    else console.log(`    No orphaned media found.`);
                }
            }
            if (options.recompressImages) {
                console.log(`  -> Stage 3: Re-compressing images with quality level ${options.imageQuality}...`);
                const imageJobs = [];
                zip.folder('word/media').forEach((relativePath, file) => {
                    const fullPath = `word/media/${file.name}`;
                    const extension = path.extname(fullPath).toLowerCase();
                    if (['.jpeg', '.jpg', '.png'].includes(extension)) {
                        const job = async () => {
                            const imageBuffer = await file.async('nodebuffer');
                            const originalImageSize = imageBuffer.length;
                            let newBuffer;
                            try {
                                if (extension === '.png') {
                                    newBuffer = await sharp(imageBuffer).png({ quality: options.imageQuality, effort: 8 }).toBuffer();
                                } else {
                                    newBuffer = await sharp(imageBuffer).jpeg({ quality: options.imageQuality }).toBuffer();
                                }
                                if (newBuffer.length < originalImageSize) {
                                    console.log(`    - Compressing ${file.name} (${(originalImageSize / 1024).toFixed(0)}KB -> ${(newBuffer.length / 1024).toFixed(0)}KB)`);
                                    zip.file(fullPath, newBuffer);
                                }
                            } catch (sharpErr) { console.warn(`    - Could not process ${file.name}: ${sharpErr.message}`); }
                        };
                        imageJobs.push(job());
                    }
                });
                await Promise.all(imageJobs);
            }
            const stylesFile = zip.file('word/styles.xml');
            if (stylesFile) {
                const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: "@_" });
                const docFile = zip.file('word/document.xml');
                if (docFile) {
                    const docXml = parser.parse(await docFile.async('string'));
                    const stylesXml = parser.parse(await stylesFile.async('string'));
                    if (stylesXml['w:styles'] && stylesXml['w:styles']['w:style']) {
                        const allStyles = Array.isArray(stylesXml['w:styles']['w:style']) ? stylesXml['w:styles']['w:style'] : [stylesXml['w:styles']['w:style']];
                        const usedStyleIds = getUsedStyleIds(docXml);
                        const stylesMap = new Map(allStyles.map(s => [s['@_w:styleId'], s]));
                        const requiredStyleIds = new Set();
                        const queue = [...usedStyleIds];
                        allStyles.forEach(s => { if (s['@_w:default'] === '1') { queue.push(s['@_w:styleId']); } });
                        while (queue.length > 0) {
                            const currentId = queue.shift();
                            if (!currentId || requiredStyleIds.has(currentId)) continue;
                            requiredStyleIds.add(currentId);
                            const styleDef = stylesMap.get(currentId);
                            const basedOnId = styleDef?.['w:basedOn']?.['@_w:val'];
                            if (basedOnId) queue.push(basedOnId);
                            const linkId = styleDef?.['w:link']?.['@_w:val'];
                            if (linkId) queue.push(linkId);
                        }
                        const filteredStyles = allStyles.filter(s => requiredStyleIds.has(s['@_w:styleId']));
                        if (filteredStyles.length < allStyles.length) {
                            console.log(`  -> Stage 4: Stripping ${allStyles.length - filteredStyles.length} unused style definitions.`);
                            stylesXml['w:styles']['w:style'] = filteredStyles;
                            const builder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: "@_", format: false });
                            const newStylesXmlString = builder.build(stylesXml);
                            zip.file('word/styles.xml', newStylesXmlString);
                        } else { console.log(`  -> Stage 4: No unused styles found to remove.`); }
                    }
                }
            }
            const outputBuffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
            const newSize = outputBuffer.length;
            if (newSize >= originalSize) { console.log(`\nOptimization did not result in a smaller file. No file was written.`); return; }
            const outputPath = options.overwrite ? filePath : path.join(path.dirname(filePath), `${path.parse(filePath).name}_optimized.docx`);
            await fs.writeFile(outputPath, outputBuffer);
            const savings = originalSize - newSize;
            console.log(`\n  ✅ Successfully wrote file: ${path.basename(outputPath)}`);
            console.log(`     Original Size: ${(originalSize / 1024).toFixed(0)} KB, New Size:      ${(newSize / 1024).toFixed(0)} KB, Saved:         ${(savings / 1024).toFixed(0)} KB (${((savings / originalSize) * 100).toFixed(1)}%)`);
        } catch (err) {
            console.error(`\n❌ Failed to optimize ${path.basename(filePath)}. Error:`, err.message);
        }
    }
    
    function createOptimizedDocument(paragraphs) { return new Document({ creator: "DocTool", title: "Optimized Document", description: "Generated by a script", styles: { paragraphStyles: [ { id: "Normal", name: "Normal", quickFormat: true, run: { size: 22 }, paragraph: { spacing: { line: 276 } } } ] }, sections: [{ properties: {}, children: paragraphs }], }); }
    
    async function processDocumentForSplit(filePath, mode, value) { try { const stats = await fs.stat(filePath); let filesToProcess = []; if (stats.isFile()) { if (path.extname(filePath).toLowerCase() === '.docx') filesToProcess.push(filePath); } else if (stats.isDirectory()) { const allFiles = await fs.readdir(filePath); filesToProcess = allFiles.filter(file => path.extname(file).toLowerCase() === '.docx' && !/_rebuilt\.docx$|_optimized\.docx$|_\d{2,}\.docx$/.test(file)).map(file => path.join(filePath, file)); } if (!filesToProcess.length) return console.log("No applicable .docx files found to split."); console.log(`Found ${filesToProcess.length} file(s) to split.`); for (const fp of filesToProcess) { console.log(`\n--- Splitting: ${path.basename(fp)} ---`); const text = (await mammoth.extractRawText({ path: fp })).value; const allParagraphs = text.split('\n\n').filter(p => p.trim().length > 0); const outputPrefix = path.join(path.dirname(fp), path.parse(fp).name); if (mode === 'parts') { const total = allParagraphs.length; const perFile = Math.ceil(total / value); console.log(`Splitting into ${value} parts...`); for (let i = 0; i < value; i++) { const start = i * perFile; if (start >= total) break; const chunk = allParagraphs.slice(start, start + perFile); const doc = createOptimizedDocument(chunk.map(p=>new Paragraph({text: p}))); const outPath = `${outputPrefix}_${(i+1).toString().padStart(2,'0')}.docx`; await fs.writeFile(outPath, await Packer.toBuffer(doc)); console.log(`  -> Wrote ${path.basename(outPath)}`); } } else { console.log("Split by size not yet implemented."); } } } catch (err) { console.error(`Error processing path for split: ${err.message}`); } }
    
    async function mergeDocuments(directoryPath, deleteSource, useOriginalName) { try { const files = await fs.readdir(directoryPath); const groups = files.filter(f => /_\d{2,}\.docx$/.test(f)).reduce((acc, f) => { const base = f.replace(/_\d{2,}\.docx$/, ''); if (!acc[base]) acc[base] = []; acc[base].push(f); return acc; }, {}); if (Object.keys(groups).length === 0) return console.log("No sets of split files found to merge."); for (const baseName in groups) { console.log(`\n--- Merging set: ${baseName} ---`); const fileNames = groups[baseName].sort(); const paragraphs = []; for (const f of fileNames) { const text = (await mammoth.extractRawText({path: path.join(directoryPath, f)})).value; paragraphs.push(...text.split('\n\n').filter(p => p.trim().length > 0)); } const doc = createOptimizedDocument(paragraphs.map(p=>new Paragraph({text:p}))); const outName = useOriginalName ? `${baseName}.docx` : `${baseName}_merged.docx`; await fs.writeFile(path.join(directoryPath, outName), await Packer.toBuffer(doc)); console.log(`  -> Created ${outName}`); if (deleteSource) { for (const f of fileNames) await fs.unlink(path.join(directoryPath, f)); console.log("  -> Deleted source split files."); } } } catch (err) { console.error(`Error during merge: ${err.message}`); } }
    
    // =================================================================================
    // YARGS COMMAND-LINE INTERFACE
    // =================================================================================
    
    async function main() {
        await yargs(hideBin(process.argv))
            .command(
                'clean <path>',
                'Removes all generated files from a directory.',
                yargs => yargs.positional('path', { describe: 'The directory to clean', type: 'string' }).option('force', { alias: 'y', type: 'boolean', description: 'Skip confirmation prompt.', default: false }),
                async (argv) => { await cleanDirectory(argv.path, argv.force); }
            )
            .command(
                'rebuild <path>',
                'Rebuilds .docx file(s) from scratch to remove all bloat. Preserves basic formatting.',
                yargs => yargs.positional('path', { describe: 'Path to a source .docx file or a directory', type: 'string' })
                    .option('overwrite', { type: 'boolean', description: 'Overwrite the original file(s).', default: false })
                    .option('force', { alias: 'y', type: 'boolean', description: 'Skip confirmation prompts.', default: false }),
                async (argv) => {
                    try {
                        const stats = await fs.stat(argv.path);
                        let filesToProcess = [];
                        if (stats.isFile()) {
                            if (path.extname(argv.path).toLowerCase() === '.docx') filesToProcess.push(argv.path);
                        } else if (stats.isDirectory()) {
                            const allFiles = await fs.readdir(argv.path);
                            filesToProcess = allFiles.filter(file => path.extname(file).toLowerCase() === '.docx' && !/_rebuilt\.docx$|_optimized\.docx$|_\d{2,}\.docx$/.test(file)).map(file => path.join(argv.path, file));
                        }
                        if (!filesToProcess.length) return console.log("No applicable .docx files found to rebuild.");
                        
                        let proceed = true;
                        if (argv.overwrite && !argv.force) {
                            console.log(`The following ${filesToProcess.length} file(s) will be overwritten:`);
                            filesToProcess.forEach(f => console.log(`  - ${path.basename(f)}`));
                            proceed = await askForConfirmation('\nThis action cannot be undone. Are you sure? (y/n) ');
                            if (!proceed) console.log("Operation aborted by user.");
                        }
                        if (proceed) {
                            console.log(`\nProcessing ${filesToProcess.length} file(s)...`);
                            for (const filePath of filesToProcess) await rebuildDocument(filePath, argv);
                        }
                    } catch (err) { console.error(`Error: ${err.message}`); }
                }
            )
            .command(
                'optimize <path>',
                'Attempts to optimize .docx file(s) by cleaning media, fonts, and styles.',
                yargs => yargs.positional('path', { describe: 'Path to a source file or directory', type: 'string' })
                    .option('overwrite', { type: 'boolean', description: 'Overwrite the original file(s).', default: false })
                    .option('force', { alias: 'y', type: 'boolean', description: 'Skip confirmation prompts.', default: false })
                    .option('remove-embedded-fonts', { type: 'boolean', default: false })
                    .option('remove-unused-media', { type: 'boolean', default: false })
                    .option('recompress-images', { type: 'boolean', default: false })
                    .option('image-quality', { type: 'number', default: 80 }),
                async (argv) => {
                    try {
                        const stats = await fs.stat(argv.path);
                        let filesToProcess = [];
                        if (stats.isFile()) {
                            if (path.extname(argv.path).toLowerCase() === '.docx') filesToProcess.push(argv.path);
                        } else if (stats.isDirectory()) {
                            const allFiles = await fs.readdir(argv.path);
                            filesToProcess = allFiles.filter(file => path.extname(file).toLowerCase() === '.docx' && !/_optimized\.docx$|_\d{2,}\.docx$/.test(file)).map(file => path.join(argv.path, file));
                        }
                        if (!filesToProcess.length) return console.log("No applicable .docx files found to optimize.");
                        let proceed = true;
                        if (argv.overwrite && !argv.force) {
                            console.log(`The following ${filesToProcess.length} file(s) will be overwritten:`);
                            filesToProcess.forEach(f => console.log(`  - ${path.basename(f)}`));
                            proceed = await askForConfirmation('\nThis action cannot be undone. Are you sure? (y/n) ');
                            if (!proceed) console.log("Operation aborted by user.");
                        }
                        if (proceed) {
                            console.log(`\nProcessing ${filesToProcess.length} file(s)...`);
                            for (const filePath of filesToProcess) await optimizeDocument(filePath, argv);
                        }
                    } catch (err) { console.error(`Error: ${err.message}`); }
                }
            )
            .command(
                'split <path>',
                'Split a .docx file into multiple plain-text versions.',
                yargs => yargs.positional('path', { describe: 'Path to source file or directory', type: 'string' }).option('by', { choices: ['parts', 'size'], default: 'parts' }).option('value', { type: 'number', default: 10 }),
                async (argv) => { await processDocumentForSplit(argv.path, argv.by, argv.value); }
            )
            .command(
                'merge <path>',
                'Merge sets of plain-text split .docx files.',
                yargs => yargs.positional('path', { describe: 'Path to directory with split files', type: 'string' }).option('delete', { type: 'boolean', default: false }).option('use-original-name', { type: 'boolean', default: false }),
                async (argv) => { await mergeDocuments(argv.path, argv.delete, argv.useOriginalName); }
            )
            .demandCommand(1, 'You must provide a command: clean, rebuild, optimize, split, or merge.')
            .alias('help', 'h')
            .strict()
            .parse();
    }
    
    await main();
}


/**
 * BOOTSTRAPPER
 * This self-executing function is the entry point of the script.
 */
(async () => {
    const missing = findMissingDependencies();

    if (missing.length > 0) {
        console.log("\nThis script requires some additional packages to run.");
        console.log("Missing:", missing.join(', '));
        
        const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
        const answer = await new Promise(resolve => {
            rl.question('\nMay I install them for you using npm? (y/n) ', resolve);
        });
        rl.close();

        if (answer.toLowerCase() === 'y' || answer.toLowerCase() === 'yes') {
            const success = await installDependencies(missing);
            if (!success) {
                process.exit(1); // Exit if installation failed
            }
        } else {
            console.log("Installation aborted. Please install the missing packages manually and run the script again.");
            process.exit(1);
        }
    }
    
    console.log("\nAll dependencies satisfied. Starting DocScrubber...");
    await startApp();
    console.log("\n✅ All commands complete.");
})().catch(err => {
    console.error("\nA critical error occurred during bootstrap: ", err);
});
