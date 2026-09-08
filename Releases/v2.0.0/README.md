# ComComTag

ComComTag is a lightweight, portable Windows application for standardising MP3 tags, pulling cover art directly from Instagram, and consolidating chapters seamlessly into beautifully labelled M4B Audiobooks.

![ComComTag Icon](icon.ico)

## Features

- **Batch MP3 Tagging & Renaming**: Select multiple MP3s and batch-update Artist, Show, and Venue/Location with clean filename formatting (`YYYY-MM-DD - Artist - Show - Location (track).mp3`).
- **Multi-Artist Support (`+` Button)**: Dynamically add extra Artist dropdowns for split sets or shared shows. Generates grammatically correct filenames (`Artist A and Artist B` or `Artist A, Artist B, and Artist C`) or groups additional artists in filenames (e.g. `Artist A and friends`) while writing all individual performers to ID3 tags.
- **Filesystem-Safe Filename Sanitisation**: Automatically converts illegal filesystem characters (`:`, `/`, `\`, `"`, `*`, `?`, `<`, `>`, `|`) and ampersands (`&` -> `+`) into clean, system-safe equivalents while keeping rich punctuation in ID3 metadata tags.
- **Clear Existing Tags**: One-click button to strip all ID3 metadata tags (or reset only ComComTag-managed fields based on Settings) from selected files on disk.
- **Autofill & Custom Lists**: Configure and autofill custom lists of common Artists and Venues with smart autocomplete, with configurable auto-learning toggles.
- **Folder Refresh**: Dedicated one-click refresh button (`F5`) to instantly reload the active folder without re-browsing.
- **100% Pure C# Instagram Scraper & Carousel Picker**: Direct native Instagram scraper with an interactive multi-image carousel modal for choosing uncropped cover art slides. **No Python or external executables needed!**
- **Automatic Temp Cleanup**: Optional "Flush temporary Instagram files on exit" setting (enabled by default) to keep your temporary directories clean.
- **Theme Support (System / Light / Dark)**: Full dark mode styling with Windows 10/11 immersive title bar support, automatically following system preferences or selectable via Settings.
- **Cover Art Drag & Drop & Click-to-Browse**: Drag images directly onto the cover art area or click to open the file browser.
- **Fast M4B Audiobook Builder**: High-performance in-memory track duration probing, native chaptering, double-click shifting, sorting, and stitching into M4B audiobooks using FFmpeg with real-time progress and cancellation support.
- **Built-in User Guide & Help (`F1`)**: Comprehensive Help dialogue and Quick Start guide written in UK English.
- **Configurable Settings**: Multi-tab settings dialog with dedicated tabs for General options, Audiobook encoder configuration, and Naming Conventions.

## Compilation

You can compile ComComTag with **one click** using `build.bat`, which produces `ComComTag_v2.0.0.exe` and packages a complete release archive into `Releases/`:

```powershell
.\build.bat
```

Or manually via the native Windows C# compiler (`csc.exe`):
```powershell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe `
  /t:winexe `
  /win32icon:icon.ico `
  /res:icon.ico,ComComTag.icon.ico `
  /res:TagLibSharp.dll,ComComTag.TagLibSharp.dll `
  /out:ComComTag_v2.0.0.exe `
  /r:System.Windows.Forms.dll `
  /r:System.Drawing.dll `
  /r:TagLibSharp.dll `
  AssemblyInfo.cs Program.cs MainForm.cs M4bBuilder.cs Settings.cs SettingsForm.cs InputPromptDialog.cs ImageCarouselDialog.cs InstagramDownloader.cs HelpDialog.cs ThemeHelper.cs
```

## Distribution

ComComTag is completely self-contained and ultra-lightweight (~150 KB):

1. **Standalone Executable**: `ComComTag_v2.0.0.exe` (and alias `ComComTag.exe`) embeds `TagLibSharp.dll` internally and resolves it in-memory. **No external DLLs are required.**
2. **FFmpeg**: To build M4B Audiobooks, the end-user **MUST have `ffmpeg.exe`**. Configure its location via **Edit -> Settings** or place it alongside `ComComTag_v2.0.0.exe`.

## Documentation

Full documentation is available in the **[Docs/](Docs/README.md)** directory:
- [01. Architecture & System Overview](Docs/01_Architecture_and_Overview.md)
- [02. Features & Deep Dive](Docs/02_Features_and_Deep_Dive.md)
- [03. Compilation, Versioning & Releases](Docs/03_Compilation_and_Execution_Guide.md)
- [04. Code Audit & Optimization Blueprint](Docs/04_Code_Audit_Redundancy_and_Optimization.md)
- [05. Instagram Carousel Scraping & Troubleshooting](Docs/05_Instagram_Carousel_Scraping_and_Troubleshooting.md)
- [06. Instagram Extraction & Chrome Extension Guide](Docs/06_Instagram_Extraction_and_Chrome_Extension_Guide.md)
