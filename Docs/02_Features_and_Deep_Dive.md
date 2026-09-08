# 02. Features & Deep-Dive Analysis

This document provides an exhaustive, technical deep dive into every subsystem, data structure, and execution flow inside **ComComTag**.

---

## 1. Tab 1: "Tag & Rename MP3" Engine

The primary purpose of the Tagging Tab is to standardize inconsistent or raw audio files into a unified format:
```
Filename:   YYYY-MM-DD - Artist - [Show - ]Location.mp3
ID3 Title:  YYYY-MM-DD - [Show - ]Location
ID3 Artist: Artist (Saved to both Performers and AlbumArtists)
ID3 Album:  Show (Optional)
ID3 Year:   YYYY (Derived from Date)
ID3 Cover:  [Embedded JPEG/PNG Front Cover]
```

### 1.1 UI Features & Layout Design
- **Top Folder Bar**: 
  - **Browse Folder...** *(Folder Tab Icon)*: Selects a folder containing MP3 files.
  - **Refresh Folder** *(Sync Arrow Icon, `F5`)*: Reloads current folder.
  - **Open in Explorer** *(Folder Launch Icon)*: Opens the active folder in Windows Explorer.
- **Top Action Toolbar with ToolTips**:
  - **Clear Tags** *(Red Broom Icon)*: Resets input fields, with a confirmation prompt to physically strip and wipe all ID3 metadata tags (when 'Clear button erases all ID3 tags' is enabled in Settings), or reset only ComComTag-managed fields (when unticked).
  - **From Filename** *(Magic Wand Icon)*: Hover tooltip: *"Attempts to extract ID3 tags from filename given current Naming Conventions in Settings"*. Uses a two-stage token replacement regex engine that supports custom delimiters, parentheses syntax e.g. `{Date} - {Artist} - {Show} ({Location})`, and fallback heuristics for voice recorder recordings. Automatically greyed out when multiple files are selected.
  - **Copy Tags / Paste Tags** *(Clipboard Icons)*: Copies and pastes metadata between tracks. Copy Tags is automatically greyed out when multiple files are selected.
- **Maximized Left MP3 List & Compact Controls**:
  - The right-side metadata panel is narrowed to `310px`, shortening Artist, Show, and Venue text fields by 50% (`175px` / `205px`) to expand the left MP3 file list horizontally, ensuring long filenames are displayed clearly without clipping.
- **Multi-Artist Dynamic '+' Button**:
  - A '+' button beside the primary Artist field enables adding multiple artist fields for split sets or shared shows.
  - The '+' button only enables once the preceding artist field has text entered.
  - Generates filenames with grammatically formatted multiple artists:
    - 2 artists: `Comedian A and Comedian B`
    - 3+ artists: `Comedian A, Comedian B, and Comedian C`
  - In ID3 tagging, all entered artists are saved as Contributing Artists (`Performers` / `TPE1`), and Album Artist (`TPE2`) is assigned based on the "Use first artist as album artist" setting.
- **Field Inclusion Checkboxes**:
  - Individual checkboxes precede **Date**, **Artist**, **Show**, and **Venue** (all checked by default).
  - Unchecking any box omits that field from the generated *filename* on disk while still saving the full value into the ID3 *metadata* tags.
- **Streamlined Cover Art Direct Controls**:
  - Direct image preview box with zero manual text editing.
  - **Single Click**: Click anywhere on the cover box to browse an image file from your computer.
  - **Drag and Drop**: Drag any image (`.jpg`, `.jpeg`, `.png`, `.webp`, `.bmp`) directly onto the preview box.
  - **Instagram Icon Button**: Compact camera glyph button to download artwork directly from Instagram.
  - **Carousel `<` / `>` Buttons**: Fast slide cycling.
  - **Double-Click**: Re-opens the multi-image carousel picker modal.
- **Smart "Rename MP3" Checkbox & Save Button**:
  - **Rename MP3** checkbox (ticked by default).
  - When checked, the button displays **"Save Tags & Rename"** and updates both ID3 tags and physical filenames on disk.
  - When unchecked, the button dynamically updates to **"Save ID3 tags"** and writes all metadata while preserving physical filenames untouched.
- **System-Safe Filename Sanitization**:
  - Automatically converts illegal filesystem characters (`\ / : * ? " < > |`) and ampersands (`&` $\rightarrow$ `+`) into clean, readable filename equivalents (`:` $\rightarrow$ ` - `, `/` and `\` $\rightarrow$ `-`, `"` $\rightarrow$ `'`, `<` and `>` $\rightarrow$ `[` and `]`, `|` $\rightarrow$ `-`, strips `?` and `*`).

---

## 2. 100% Native C# Instagram Artwork Scraper & Carousel Picker

### 2.1 Multi-Tiered Architecture (`InstagramDownloader.cs`)
ComComTag uses a lightweight, pure C# HTTP/JSON scraper with a 3-tier resolution engine:
1. **Tier 1 (Polaris GraphQL API)**:
   - Negotiates CSRF session tokens with Instagram's web client.
   - Executes `PolarisPostRootQuery` (`doc_id: 27128499623469141`).
   - Retrieves all carousel items at full $1440 \times 1440$ / $1080\text{p}$ uncropped resolution, bypassing center-crop parameters (`stp=c...`).
2. **Tier 2 (Official Meta oEmbed API)**:
   - Queries `https://www.instagram.com/api/v1/oembed/` for signed uncropped square $640 \times 640$ assets.
3. **Tier 3 (Crawler Meta Scraping)**:
   - Uses crawler User-Agents (`facebookexternalhit/1.1`) as a secondary fallback.
4. **Flush Temp Files on Exit**:
   - When enabled (default: `true`), all downloaded temporary Instagram images and directories are automatically deleted upon closing the application.

### 2.2 Interactive Multi-Image Carousel Modal (`ImageCarouselDialog`)
When an Instagram post contains multiple images:
1. `ImageCarouselDialog` automatically opens.
2. Users see a large high-resolution preview with an image counter (`Image 1 of 6`).
3. Navigation can be performed using:
   - **`< Previous`** / **`Next >`** buttons
   - **Left / Right Arrow keys** on the keyboard
4. Clicking **"Select This Cover Art"** (or pressing Enter) applies the selected slide to the cover art box.

---

## 3. Tab 2: High-Speed M4B Audiobook Builder

### 3.1 60/40 Proportional Layout & Responsive Expansion
- **Full Width Utilization**: Eliminates dead space on the right by utilizing the full canvas width.
- **60/40 Split**: Section 1 (**Available MP3s**) takes up 60% of the available list width so long filenames are fully readable without horizontal scrolling; Section 2 (**Chapters**) takes up the remaining 40%.
- **Transfer & Reordering Icon Buttons**:
  - **Add to Chapters** *(Right Double Arrow Icon)*: `ToolTip: "Add selected MP3s as chapters (Enter / Double-Click)"`
  - **Remove from Chapters** *(Left Double Arrow Icon)*: `ToolTip: "Remove selected chapters (Del / Double-Click)"`
  - **Move Chapter Up** *(Upward Arrow Icon)*: `ToolTip: "Move chapter up (Ctrl+Up)"`
  - **Move Chapter Down** *(Downward Arrow Icon)*: `ToolTip: "Move chapter down (Ctrl+Down)"`
- **List Interaction & In-Line Renaming**:
  - **Transfer via Double-Click / Buttons**: Double-clicking files in **Available MP3s** or pressing Enter shifts the track directly into the chapter list.
  - **In-Line Chapter Renaming**: Double-clicking any chapter in **Chapters** or selecting it and pressing `F2` activates in-line in-place editing. Press `Enter` to commit or `Escape` to cancel.
  - **Source Filename Tooltips**: Hovering the cursor over any chapter item dynamically displays the original source filename from disk in a tooltip.
  - **Full Height Chapters Queue**: Removing the old separate rename box allows both list boxes to span the full canvas height seamlessly.
- **Shortened Metadata Inputs & 90x90 Cover Preview**:
  - Metadata fields (`Artist`, `Album`, `Venue`) are ~245px, closely matching the **Tag & Rename MP3** tab.
  - The 90×90 cover art preview box sits cleanly in the center column alongside the metadata inputs.
  - The **Build Audiobook** button (`Ctrl+B`) sits on the right with dedicated breathing room.
- **Keyboard Shortcuts**:
  - `Ctrl + S`: Save Tags & Rename on disk.
  - `Ctrl + B`: Build Audiobook.
- **Bottom Status Bar**: The **Filename Preview** and **Progress Bar** are docked near the bottom edge across the full tab width.

### 3.2 In-Memory TagLib# Duration Probing (400x Speedup)
Rather than spawning `ffmpeg.exe` per track in a loop, `M4bBuilder` uses `TagLib.File.Create(path).Properties.Duration.TotalMilliseconds` directly in-memory. Probing 50 tracks finishes in under 50 milliseconds instead of 20+ seconds.

### 3.3 FFmpeg Transcoding Pipeline
- Generates `files.txt` (concat demuxer list with escaped single quotes).
- Generates `metadata.txt` in standard `FFMETADATA1` format with millisecond offsets (`TIMEBASE=1/1000`).
- Executes FFmpeg to transcode to AAC inside an M4B wrapper with optional embedded MJPEG cover art.
- Parses `stderr` timestamps in real time to update the progress bar via `IProgress<int>`.

---

## 4. Settings Subsystem (`SettingsForm.cs`)

Accessible via **Edit -> Settings**, organised into three tabs:

### 4.1 "General" Tab
- **FFmpeg Path**: Path to `ffmpeg.exe` with file browser.
- **Instagram Temp Dir**: Custom download directory with folder browser.
- **Theme**: Dropdown (`System Default`, `Light`, `Dark`). Automatically detects Windows dark mode preferences and enables native Windows 10/11 immersive dark title bars.
- **Flush temporary Instagram files on exit**: Automatically purges temporary downloads on application close.
- **Add new artists to dropdown & Add new venues to dropdown**: Side-by-side checkboxes to toggle auto-learning new artists and venues into settings.
- **Use first artist as album artist**: Saves first artist as Album Artist (`TPE2`) and all named performers as Contributing Artists (`TPE1`).
- **Clear button erases all ID3 tags**: Toggles between stripping all ID3 tags vs. wiping only ComComTag-managed fields.
- **Group additional artists in filename**: Checkbox (enabled by default) with configurable threshold numeric spinner (default: `if count >= 3`) and custom text box (default `"and friends"`). When total entered artists meets or exceeds the threshold, generates filenames like `Comedian A and friends` while writing all individual artists to ID3 tags. If below threshold (e.g. 2), lists each artist by name.
- **Artists & Venues**: Expanded multi-line editors taking full dialog height and width.

### 4.2 "Audiobook" Tab
- **Default Bitrate**: Dropdown (`64k` to `320k`) setting default encoder bitrate for M4B compilation.
- **Default Artist**: Customisable, user-editable text field (default `"Various Artists"`) that populates the initial artist on the Audiobook tab.
- **Append duration to chapter name**: Checkbox (default: checked) controlling whether track duration `(mm:ss)` is automatically appended to chapter titles upon import.

### 4.3 "Conventions" Tab
- **Date Format Selector**:
  - Dropdown providing standard formatting patterns (`yyyy-MM-dd`, `dd-MM-yyyy`, `dd-MMM-yy`, `yyyy.MM.dd`, `dd.MM.yyyy`, `yyyy_MM_dd`, `dd_MM_yyyy`, `dd-MMM-yyyy`, `yyyyMMdd`).
  - Live dynamic example displaying today's date formatted according to the selection.
- **Custom Naming Pattern & Placeholders**:
  - Configurable template string (default: `{Date} - {Artist} - {Show} - {Location}`).
  - Quick-insert buttons for `{Date}`, `{Artist}`, `{Show}`, `{Location}`, `({Track})`, `" - "`, `" _ "`.
  - **Case Sensitivity**:
    - `{Artist}` / `{Show}` / `{Location}` = As entered.
    - `{ARTIST}` / `{SHOW}` / `{LOCATION}` = ALL CAPS.
    - `{artist}` / `{show}` / `{location}` = lowercase.
  - **Smart Cleanup**: Automatically cleans up consecutive separators and trims missing optional fields.
- **Interactive Experiment Playground**:
  - Real-time sandbox initialized with default dummy data (*1999-06-13*, *Wally Bazoom*, *Smile Time*, *The Aigburth Arms, Liverpool*).
  - Any custom values entered by the user are automatically persisted in `settings.ini`.
