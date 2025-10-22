\# FileCleaner



FileCleaner is a PowerShell + WPF GUI tool for batch-renaming and cleaning media files.  

It is designed to be \*\*modular, error-resistant, and user-friendly\*\*, with a focus on clarity, reversibility, and professional polish.



---



\## ✨ Features



\### 📂 Folder \& File Selection

\- Browse for a folder using a standard Windows folder picker.

\- Optionally include subfolders.

\- Filter by file type (Video, Audio, Pictures, Documents, Archives).

\- Add custom extensions (comma-separated).



\### 🔖 Prefix Management

\- \*\*Change Prefix\*\*: Replace an old prefix with a new one.

\- \*\*Add Prefix to All Files\*\*: Ensure all files start with a given prefix.

\- \*\*Detected Prefix\*\*: Auto-detected prefix can be applied if no old prefix is specified.

\- \*\*Only If Old Prefix\*\*: Restrict processing to files that already have the old prefix.

\- \*\*Dry-run Mode\*\*: Simulate renames without touching the filesystem.



\### 📺 Season \& Episode Tagging

\- Add or update \*\*Season\*\* and \*\*Episode\*\* tags.

\- Configure digit padding (e.g., `S01E01`, `S001E0001`).

\- Choose whether Season comes before Episode (`S01E01`) or not.

\- \*\*Renumber All Files Alphabetically\*\*:

&nbsp; - Sorts files alphabetically.

&nbsp; - Assigns fresh sequential episode numbers (`E01`, `E02`, …).

&nbsp; - Always places tags immediately after the prefix.

&nbsp; - Strips stray tags elsewhere in the filename.



\### 🧹 Cleaning Options

\- Remove common tokens (`720p`, `1080p`, `4K`, `HD`).

\- Remove custom tokens (comma-separated).

\- Ensures episode/season tags appear \*\*only after the prefix\*\*.



\### 🗂 Metadata Support

\- Optional integration with `ffprobe` or `exiftool`.

\- Extracts audio metadata (Artist, Album, Title).

\- Extracts video metadata (Show, Season, Episode, Title).

\- Displays metadata summary in the preview grid.



\### 🖥 Preview Grid

\- Shows \*\*Original\*\*, \*\*Proposed\*\*, \*\*Status\*\*, \*\*Directory\*\*, and \*\*Metadata\*\*.

\- Automatically marks unchanged rows as `"No change"`.

\- Allows selective application of changes.



\### ⚡ Apply / Undo / Export

\- \*\*Apply\*\*: Renames files according to preview.

\- \*\*Undo\*\*: Reverts the last batch of renames.

\- \*\*Export CSV\*\*: Saves preview results to a CSV file.

\- \*\*Reset\*\*: Clears all inputs and preview.



---



\## 🛠 Implementation Details



\### Prefix \& Episode Placement

\- Episode and season tags are \*\*always placed immediately after the prefix\*\*.

\- If no explicit prefix is provided:

&nbsp; - When \*\*Renumber All\*\* is checked, the tool infers a prefix:

&nbsp;   - Leading token before the first space/underscore/dash.

&nbsp;   - If no separator, the entire base name is treated as prefix.

&nbsp; - Tags are then inserted after this inferred prefix.



\### Renumbering Logic

\- \*\*Renumber All\*\* forces `-ForceEpisode`, ensuring every file gets a fresh sequential episode number.

\- Old tags are stripped from the remainder to prevent duplicates.

\- Counter increments consistently across files.



\### Apply Handler

\- Recomputes proposed names at apply time to guarantee consistency with preview.

\- Skips `"No change"` rows.

\- Honors \*\*Dry-run\*\* mode (marks rows as `"Dry-run"` without renaming).

\- Logs operations for undo.



\### Undo Handler

\- Reverts only files that were actually renamed in the last batch.

\- Reports how many files were successfully reverted.



---



\## 🚀 Usage



1\. Launch the script in PowerShell:

&nbsp;  ```powershell

&nbsp;  .\\FileCleaner.ps1

