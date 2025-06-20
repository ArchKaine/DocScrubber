# DocScrubber

**A powerful command-line toolkit for batch processing, optimizing, and repairing Microsoft Word `.docx` files.**

DocScrubber is a Node.js-based utility designed to handle common but difficult tasks for managing `.docx` documents. Whether you need to drastically reduce the size of bloated files, split large documents into smaller parts, or simply clean up a workspace, DocScrubber provides a suite of powerful, command-driven tools to get the job done.

---

### Core Features

* **Rebuild:** Creates a brand-new, clean version of a document to eliminate all possible bloat from tracked changes, comments, and other invisible data. Preserves basic formatting (headings, lists, tables, etc.).
* **Optimize:** Surgically removes specific sources of bloat without a full rebuild.
    * Remove embedded fonts.
    * Remove unused (orphaned) images and media.
    * Re-compress existing images with adjustable quality settings.
    * Strip unused style definitions.
* **Split:** Splits large documents into smaller, more manageable plain-text files, either by a fixed number of parts or by a maximum file size.
* **Merge:** Combines a set of split plain-text files back into a single document.
* **Clean:** Safely removes all generated files (`_rebuilt`, `_optimized`, etc.) from your project folder, with an interactive prompt to prevent accidents.
* **Batch Processing:** All commands can operate on either a single file or an entire directory of `.docx` files.

---

### Prerequisites

* [Node.js](https://nodejs.org/) (v16 or later recommended)
* NPM (comes bundled with Node.js)

---

### Installation

1.  Clone the repository:
    ```bash
    git clone [https://github.com/your-username/doc-scrubber.git](https://github.com/your-username/doc-scrubber.git)
    ```
2.  Navigate into the project directory:
    ```bash
    cd doc-scrubber
    ```
3.  Install the required dependencies:
    ```bash
    npm install
    ```

---

### Usage

DocScrubber is a command-line tool. All commands follow the pattern: `node . <command> <path> [options]`

#### `rebuild`
The most powerful command for reducing file size. Rebuilds a document from scratch, preserving basic formatting but removing all invisible bloat.

**Syntax:**
`node . rebuild <path-to-file-or-folder> [--overwrite] [--force]`

**Examples:**
```bash
# Rebuild a single file, creating MyDoc_rebuilt.docx
node . rebuild "./docs/MyDoc.docx"

# Rebuild all files in a folder, overwriting the originals after a confirmation prompt
node . rebuild "./docs" --overwrite

# Rebuild and overwrite all files without a confirmation prompt
node . rebuild "./docs" --overwrite --force
