# ComComTag

ComComTag is a lightweight, portable Windows application for standardizing MP3 tags, pulling cover art directly from Instagram, and consolidating chapters seamlessly into beautifully labeled M4B Audiobooks.

![ComComTag Icon](icon.ico)

## Features

- **Batch MP3 Tagging**: Select one or multiple MP3s and instantly rename/tag them using a clean `YYYY-MM-DD - Artist - Show - Location` layout.
- **Instagram Art Downloader**: Built-in Instagram shortcode scraper (powered natively by an embedded PyInstaller package). No Python installation required by the end-user!
- **M4B Audiobook Builder**: Append track lengths natively to chapter titles, sort, and securely stitch together standalone MP3s into an M4B file using FFmpeg.
- **Zero-Clutter Settings**: Configuration is tracked invisibly in your Windows `AppData\Roaming\ComComTag` folder natively.

## Distribution

ComComTag is designed to be highly portable. When distributing or running the compiled `ComComTag.exe`, **you MUST include `TagLibSharp.dll` in the exact same directory alongside the executable.** The program requires this link library at runtime to natively read and write ID3 tags to the media wrappers.

No external Python software or `ffmpeg` path modification is required natively out of the box (the app will prompt users to locate their local `ffmpeg.exe` if they decide to build an M4B).

## Compilation Instructions

ComComTag bypasses heavy IDE build suites like Visual Studio. You can compile it directly from the source code natively using the built-in Microsoft .NET Framework compiler (`csc.exe`) available standard on all modern Windows distributions.

If you edit the Python downloader script, you must first compile it using PyInstaller:

```powershell
pip install pyinstaller
pyinstaller --onefile download_instagram.py
```

Then, from the root repository directory, run the C# compiler to embed the Icon and python `.exe` payload directly into the final desktop application:

```powershell
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /t:winexe /win32icon:icon.ico /res:icon.ico,ComComTag.icon.ico /res:dist\download_instagram.exe,download_instagram.exe /out:ComComTag.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.VisualBasic.dll /r:TagLibSharp.dll Program.cs MainForm.cs M4bBuilder.cs Settings.cs SettingsForm.cs
```

## Contributing

Feel free to open issues or PRs. Enjoy managing your audio archive natively!
