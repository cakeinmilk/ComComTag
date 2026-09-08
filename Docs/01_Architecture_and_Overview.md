# 01. Architecture & System Overview

## 1. System Overview

**ComComTag** is a standalone Windows desktop tool crafted specifically for audio archivists, podcast managers, and live performance collectors. It addresses the workflow of converting raw recordings or disparate MP3 files into a cleanly tagged, uniformly named library or a single chapterized M4B audiobook.

### Core Goals:
1. **Batch Tagging & Uniform Renaming**: Standardize files to a consistent naming scheme:
   `YYYY-MM-DD - Artist - Show - Location.mp3`
2. **Metadata Enrichment**: Automatically synchronize ID3v2 metadata (Title, Artist/Performer, Album/Show, Year, and embedded Front Cover Art) to match the file's standardized naming.
3. **Autofill Artist & Location Management**: User-configurable lists with auto-complete and auto-learning.
4. **100% Native Instagram Scraper & Visual Carousel**: Fetch promotional artwork directly from Instagram post URLs/shortcodes in pure C# (zero Python or PyInstaller dependencies) with an interactive multi-image carousel modal.
5. **High-Speed M4B Audiobook Assembly**: Probes track durations in-memory using TagLib#, generates `FFMETADATA1` chapter markers, and stitches MP3s into `.m4b` containers via FFmpeg.
6. **Pure C# / Zero External Dependencies**: Runs out-of-the-box on any Windows machine with .NET Framework 4.0+.

---

## 2. Architectural Diagram

```
+------------------------------------------------------------------------------------------------+
|                                         ComComTag.exe                                          |
|                                                                                                |
|  +---------------------+   +---------------------+   +---------------------+   +------------+  |
|  |     Program.cs      |-->|     MainForm.cs     |-->|  ImageCarouselDialog|-->|SettingsForm|  |
|  | (STAThread / Entry) |   | (WinForms UI Shell) |   | (Visual Multi-Image)|   | (Config)   |  |
|  +---------------------+   +----------+----------+   +---------------------+   +-----+------+  |
|                                       |                                              |         |
|            +--------------------------+-----------------------+                      |         |
|            |                          |                       |                      |         |
|            v                          v                       v                      v         |
|  +-------------------+      +-------------------+   +--------------------+  +---------------+  |
|  | TagLibSharp.dll   |      |  M4bBuilder.cs    |   | InstagramDownloader|  |  Settings.cs  |  |
|  | (In-Memory ID3 &  |      | (FFmpeg Process   |   | (Pure C# HTTP/JSON |  | (INI Parser/  |  |
|  |  Duration Probe)  |      |  Orchestrator)    |   |  Scraper)          |  |  Store)       |  |
|  +-------------------+      +---------+---------+   +--------------------+  +---------------+  |
|                                       |                                                        |
+---------------------------------------|--------------------------------------------------------+
                                        | Spawns
                                        v
                               +-----------------+
                               |   ffmpeg.exe    |
                               | (External CLI)  |
                               +-----------------+
```

---

## 3. Component Breakdown

| Source File | Responsibility |
| :--- | :--- |
| **[build.bat](../build.bat)** | One-click Windows batch compilation script. Locates `csc.exe` and builds `ComComTag.exe`. |
| **[Program.cs](../Program.cs)** | Application entry point (`Main()`). Configures visual styles and initializes `MainForm`. |
| **[MainForm.cs](../MainForm.cs)** | The core UI controller. Constructs tab views, handles autocomplete comboboxes, folder refresh (`F5`), background async tasks, and launches dialogs. |
| **[InstagramDownloader.cs](../InstagramDownloader.cs)** | Pure C# HTTP scraper that extracts post images from Instagram public endpoints and HTML meta tags directly without Python. |
| **[ImageCarouselDialog.cs](../ImageCarouselDialog.cs)** | Visual multi-image Instagram cover art selector modal with image counter, preview, and keyboard navigation. |
| **[InputPromptDialog.cs](../InputPromptDialog.cs)** | Native C# modal input prompt replacing `Microsoft.VisualBasic.Interaction.InputBox`. |
| **[M4bBuilder.cs](../M4bBuilder.cs)** | Static processing engine that probes durations in-memory using TagLib#, prepares FFmpeg concat and `FFMETADATA1` files, streams transcoding stderr, and parses progress. |
| **[Settings.cs](../Settings.cs)** | Configuration manager for Artists, Locations, DefaultBitrate, DefaultDirectory, FFmpegPath, InstaTempDirectory, Theme, and Naming Conventions in `%APPDATA%\ComComTag\settings.ini`. |
| **[SettingsForm.cs](../SettingsForm.cs)** | User configuration dialog supporting multi-line editing of Artists and Venues, Theme selection, FFmpeg browsing, and custom Instagram temp folder selection. |
| **[ThemeHelper.cs](../ThemeHelper.cs)** | Theme management subsystem providing Windows registry dark mode detection, Windows 10/11 immersive dark title bar P/Invoke, and recursive dark/light palette styling. |
| **[HelpDialog.cs](../HelpDialog.cs)** | Built-in User Guide and Quick Start documentation dialog with UK English guidance. |
| **[ComComTag.csproj](../ComComTag.csproj)** | Project file declaring build properties, WinForms enablement, and reference to `TagLibSharp.dll`. |

---

## 4. Storage & Configuration Contracts

- **Settings Path**: `%APPDATA%\ComComTag\settings.ini`
- **Keys**:
  - `DefaultDirectory=<path>`
  - `FFmpegPath=<path>`
  - `DefaultBitrate=<64k|96k|128k|192k|256k|320k>`
  - `InstaTempDir=<custom temp folder>`
  - `FlushTempOnExit=<True|False>`
  - `AutoAddArtists=<True|False>`
  - `AutoAddLocations=<True|False>`
  - `UseFirstArtistAsAlbumArtist=<True|False>`
  - `GroupAdditionalArtists=<True|False>`
  - `GroupAdditionalArtistsThreshold=<int>`
  - `GroupAdditionalArtistsText=<string>`
  - `DefaultM4bArtist=<string>`
  - `ClearErasesAllTags=<True|False>`
  - `Theme=<System|Light|Dark>`
  - `DateFormat=<pattern>`
  - `FilenamePattern=<pattern>`
  - `Artist=<name>` (Multiple entries)
  - `Location=<venue>` (Multiple entries)
- **Instagram Temp Storage**: Defaults to `%TEMP%\ComComTag_Insta\<shortcode>\` or custom folder set in Settings.
