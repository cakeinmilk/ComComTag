# ComComTag Documentation

Welcome to the comprehensive technical documentation for **ComComTag**.

ComComTag is a lightweight, portable Windows desktop application written in C# (Windows Forms) designed to streamline audio archival, standardize MP3 ID3 tags, download high-resolution cover art from Instagram posts in pure C#, and consolidate chapters seamlessly into beautifully labeled M4B Audiobooks with embedded metadata and artwork.

---

## Documentation Index

| Document | Description |
| :--- | :--- |
| **[01_Architecture_and_Overview.md](01_Architecture_and_Overview.md)** | High-level system architecture, UI layout, dependencies, and component design. |
| **[02_Features_and_Deep_Dive.md](02_Features_and_Deep_Dive.md)** | Deep dive into core features including batch ID3 tagging, M4B audiobook generation, and Instagram image downloading. |
| **[03_Compilation_and_Execution_Guide.md](03_Compilation_and_Execution_Guide.md)** | Comprehensive guide on how to build, run, and package the application using the standalone build script (`build.bat`) or Visual Studio. |
| **[04_Code_Audit_Redundancy_and_Optimization.md](04_Code_Audit_Redundancy_and_Optimization.md)** | Full pre-release code audit, security & bug mitigation strategies, and future improvements roadmap. |
| **[05_Instagram_Carousel_Scraping_and_Troubleshooting.md](05_Instagram_Carousel_Scraping_and_Troubleshooting.md)** | Technical deep dive into Instagram GraphQL and oEmbed scrapers, CDN crop parameters, and carousel extraction. |
| **[06_Instagram_Extraction_and_Chrome_Extension_Guide.md](06_Instagram_Extraction_and_Chrome_Extension_Guide.md)** | Full extraction methodology, CDN HMAC signature mechanics, and complete Chrome Extension (Manifest V3) implementation guide. |

---

## Quick Reference

- **Language & Platform**: 100% Pure C# (.NET Framework 4.0+ / .NET 8.0 Windows Forms)
- **Primary Dependencies**:
  - `TagLibSharp.dll` (ID3/MP4 metadata reader and writer)
  - `ffmpeg.exe` (External binary for AAC audio transcoding and M4B container muxing)
- **Build Pipeline**: One-click [`build.bat`](../build.bat) generates portable bundles in [`Releases/`](../Releases/).
- **Settings Store**: `%APPDATA%\ComComTag\settings.ini`
- **Temp Directory**: `%TEMP%\ComComTag_Insta` (or custom folder set in Settings).
