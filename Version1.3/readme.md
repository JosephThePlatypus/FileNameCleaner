# FileCleaner

**FileCleaner** is a batch file renamer and metadata manager built in PowerShell with a modern WPF GUI.  
It’s designed for speed, safety, and flexibility — giving you granular control over how filenames are cleaned, prefixed, and numbered, with full preview and undo support.

---

## ✨ Features

- **Folder Selection**
  - Choose a target folder and optionally include subfolders.
  - Status bar shows progress as files are scanned.

- **Prefix Options**
  - Change an **Old Prefix → New Prefix** (removes the old completely, replaces with the new).
  - Add a prefix to all files.
  - Restrict processing to only files with the old prefix.
  - Dry‑run mode to simulate changes without renaming.

- **Season / Episode Numbering**
  - Add season tokens (`S01`, `S001`, etc.) with configurable digit padding.
  - Add episode numbers (`E01`, `E001`, etc.) with configurable digit padding.
  - Option to renumber all files alphabetically.
  - Choose whether season comes before episode (`S01E01`) or separated (`S01 E01`).
  - If *Add Episode* is not checked, existing episode numbers are preserved.

- **Filename Cleaning**
  - Automatically strip common tokens like `720p`, `1080p`, `4k`, `HD`.
  - Add custom tokens to remove (comma‑separated).
  - Normalizes separators (`-`, `_`, `.`) to avoid duplicates.

- **File Type Filters**
  - Preset groups: Video (default), Pictures, Documents, Audio, Archives.
  - Add custom extensions (comma‑separated).
  - Only matching files are scanned.

- **Metadata Options (placeholders)**
  - Future support for audio metadata (artist/album/title).
  - Future support for video metadata (show/title/season/episode).

- **Actions**
  - Select All / None.
  - Scan & Preview changes.
  - Apply renames.
  - Undo last batch.
  - Export preview to CSV.
  - Reset options.
  - Exit.

- **Preview Grid**
  - Shows original name, proposed name, status, and metadata summary.
  - Checkbox to selectively apply changes.
  - Fully scrollable and resizable.

- **Safety**
  - Dry‑run mode.
  - Undo support.
  - Logging and CSV export for transparency.

---

## 🚀 Usage

1. Launch the app (`FileCleaner.ps1` or the packaged `FileCleaner.exe`).
2. Select a target folder.
3. Choose your options:
   - Prefix changes
   - Season/Episode numbering
   - Cleaning tokens
   - File type filters
4. Click **Scan / Preview** to see proposed changes.
5. Review the preview grid.
6. Click **Apply Changes** to rename selected files.
7. Use **Undo Last** if needed.

---

## 🛠️ Building an EXE

You can package the script into a standalone EXE using [PS2EXE](https://www.powershellgallery.com/packages/PS2EXE).

### Install PS2EXE
```powershell
Install-Module PS2EXE -Scope CurrentUser


FileCleaner/
│
├─ filecleaner.ps1        # main script (Parts 1–7 combined)
├─ FileCleaner.ico        # custom icon
├─ Build-FileCleaner.ps1  # helper script to package into EXE
├─ Logs/                  # optional: undo logs / CSV exports
└─ README.md              # this file