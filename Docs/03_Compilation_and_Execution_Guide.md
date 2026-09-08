# 03. Compilation and Execution Guide

## Overview

**ComComTag** is built as a **100% standalone single-binary executable**. `TagLibSharp.dll` is embedded directly into `ComComTag.exe` as an internal resource and resolved seamlessly in-memory at runtime.

---

## 1. Prerequisites

- **Operating System**: Windows 7, 8, 10, or 11 (x86 / x64).
- **.NET Framework**: .NET Framework 4.0 or higher (pre-installed by default on Windows 8, 10, and 11).
- **C# Compiler**: Native `csc.exe` located automatically at `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe`.
- *(Optional for M4B Audiobook creation)*: `ffmpeg.exe` placed in the program folder or configured in **Edit -> Settings**.

---

## 2. Compiling and Packaging (`build.bat`)

To build the project, simply double-click **`build.bat`** (or run `build.bat` from Command Prompt/PowerShell).

### Build Pipeline:
1. **Locates Native .NET C# Compiler** (`csc.exe`).
2. **Terminates Any Running Instances** to prevent file-locking conflicts.
3. **Compiles Single Standalone Executable**:
   ```cmd
   csc.exe /t:winexe /win32icon:icon.ico /res:icon.ico,ComComTag.icon.ico /res:TagLibSharp.dll,ComComTag.TagLibSharp.dll /out:ComComTag.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:TagLibSharp.dll AssemblyInfo.cs Program.cs MainForm.cs M4bBuilder.cs Settings.cs SettingsForm.cs InputPromptDialog.cs ImageCarouselDialog.cs InstagramDownloader.cs HelpDialog.cs ThemeHelper.cs
   ```
4. **Packages Release**:
   - Output folder: `Releases\v2.0.0\`
   - Output package: `Releases\ComComTag_v2.0.0.zip`

---

## 3. Distribution

To distribute **ComComTag** to other users:
- Simply share **`ComComTag.exe`** (or the release ZIP `Releases\ComComTag_v2.0.0.zip`).
- **No external DLLs are needed**: `TagLibSharp.dll` is embedded inside the binary.
- **First-Run Behavior**: On first launch, `ComComTag` automatically initializes `%APPDATA%\ComComTag\settings.ini` with all default conventions and settings.
