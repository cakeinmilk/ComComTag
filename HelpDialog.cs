using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace ComComTag {
    public class HelpDialog : Form {
        private DarkTabControl tabControl;
        private Button btnClose;

        public HelpDialog(int initialTabIndex = 0, string themeSetting = "System") {
            InitializeComponent(initialTabIndex);
            ThemeHelper.ApplyTheme(this, ThemeHelper.ShouldUseDarkMode(themeSetting));
        }

        private void InitializeComponent(int initialTabIndex) {
            this.Text = "ComComTag - User Guide & Documentation";
            this.Size = new Size(720, 620);
            this.MinimumSize = new Size(600, 500);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;

            // Application Icon
            try {
                var assembly = Assembly.GetExecutingAssembly();
                using (var iconStream = assembly.GetManifestResourceStream("ComComTag.icon.ico")) {
                    if (iconStream != null) this.Icon = new Icon(iconStream);
                }
            } catch { }

            tabControl = new DarkTabControl {
                Location = new Point(15, 15),
                Size = new Size(675, 505),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            var tabQuickStart = new TabPage("Quick Start");
            var tabConventions = new TabPage("Naming Patterns");
            var tabInstagram = new TabPage("Instagram Art");
            var tabShortcuts = new TabPage("Shortcuts");
            var tabAbout = new TabPage("About");

            tabQuickStart.Controls.Add(CreateRichTextBox(GetQuickStartHelpText()));
            tabConventions.Controls.Add(CreateRichTextBox(GetNamingPatternsHelpText()));
            tabInstagram.Controls.Add(CreateRichTextBox(GetInstagramHelpText()));
            tabShortcuts.Controls.Add(CreateRichTextBox(GetShortcutsHelpText()));
            tabAbout.Controls.Add(CreateRichTextBox(GetAboutHelpText()));

            tabControl.TabPages.Add(tabQuickStart);
            tabControl.TabPages.Add(tabConventions);
            tabControl.TabPages.Add(tabInstagram);
            tabControl.TabPages.Add(tabShortcuts);
            tabControl.TabPages.Add(tabAbout);

            if (initialTabIndex >= 0 && initialTabIndex < tabControl.TabPages.Count) {
                tabControl.SelectedIndex = initialTabIndex;
            }

            btnClose = new Button {
                Text = "Close",
                Size = new Size(95, 32),
                Location = new Point(595, 532),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
                DialogResult = DialogResult.OK
            };
            btnClose.Click += (s, e) => this.Close();

            this.AcceptButton = btnClose;
            this.CancelButton = btnClose;

            this.Controls.Add(tabControl);
            this.Controls.Add(btnClose);
        }

        private TextBox CreateRichTextBox(string text) {
            return new TextBox {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5F),
                Text = text.Replace("\n", Environment.NewLine),
                Margin = new Padding(10),
                BorderStyle = BorderStyle.None
            };
        }

        private string GetQuickStartHelpText() {
            return 
@"===================================================================
ComComTag - Quick Start Guide
===================================================================

ComComTag is designed to standardise, tag, and organise audio recordings into cleanly formatted MP3s or chapterised M4B audiobooks with embedded artwork.

-------------------------------------------------------------------
1. Tag & Rename MP3 Files (Tab 1)
-------------------------------------------------------------------
- Open Folder: Click the Folder icon (or press F5 to reload). You can also drag and drop a folder directly into the file list.
- Select Files: Select a single file to inspect its tags, or select multiple files (Ctrl+Click / Shift+Click) to batch tag.
- Enter Metadata:
  * Date: Recording or release date.
  * Artist / Performer: Primary performer name.
  * '+' Button (Multiple Artists): The '+' button becomes available once an artist name is typed into the box. Clicking it adds an extra artist dropdown. Each subsequent box must contain text before another field can be added.
  * Multiple Artists Filename Format:
    - Default: Two artists are joined with 'and' (e.g. 'Comedian A and Comedian B'). Three or more artists use commas and an Oxford comma before 'and' (e.g. 'Comedian A, Comedian B, and Comedian C').
    - Group Additional Artists: If 'Group additional artists in filename' is enabled in Settings -> General (on by default for count >= 3), additional artists are substituted in the filename with your configured text (e.g. 'Comedian A and friends') whenever the total number of entered artists meets or exceeds the threshold. If fewer artists than the threshold are entered (e.g. 2), each is listed individually by name. In all cases, all individual performers are written to ID3 tags.
  * Show / Album: The event, comedy show, or album name.
  * Venue / Location: Location of the performance.
- ID3 Album Artist Rules (configured in Settings -> General):
  * 'Use first artist as album artist' (Checked): The first entered artist is saved as the Album Artist (TPE2 tag), while all named artists are saved as Contributing Artists (TPE1 tag).
  * 'Use first artist as album artist' (Unchecked): If multiple artists are entered, Album Artist is set to 'Various Artists' and all named performers are saved as Contributing Artists.
- Theme & Appearance (configured in Settings -> General):
  * 'Theme': Choose between 'System Default' (automatically syncs with Windows dark/light preference), 'Light', or 'Dark' mode (with immersive dark title bar).
- Clear ID3 Tags Button & Settings:
  * 'Clear button erases all ID3 tags' (Checked in Settings -> General): Strips all ID3v1/ID3v2 tags completely from selected files on disk.
  * 'Clear button erases all ID3 tags' (Unchecked in Settings -> General): Wipes only ComComTag-managed ID3 tags (Artist, Show, Venue, Date, Cover) while preserving other embedded metadata.
- Inclusion Checkboxes:
  * Check/uncheck Date, Artist, Show, or Venue to include/exclude them from the physical filename while preserving them in ID3 tags.
- Cover Art:
  * Click the cover box to choose an image from your computer.
  * Drag and drop any image (.jpg, .png, .webp, .bmp) onto the cover box.
  * Click the Instagram camera icon to download high-resolution cover art directly from Instagram.
- Magic Wand ('From Filename'):
  * Attempts to automatically parse Date, Artist, Show, and Venue from the existing filename.
- Save:
  * Check 'Rename MP3' to write ID3 tags AND rename the physical file on disk (button displays 'Save Tags & Rename', or press Ctrl+S).
  * Uncheck 'Rename MP3' to write ID3 tags while keeping the original filename (button displays 'Save ID3 tags', or press Ctrl+S).

-------------------------------------------------------------------
2. Build Audiobook (Tab 2)
-------------------------------------------------------------------
- Add Chapters:
  * Double-click any track in 'Available MP3s' to instantly add it as a chapter.
  * Or select tracks and click the '>' button (or press Enter).
- Remove Chapters:
  * Select chapters and click '<' (or press Delete / Backspace).
- Reorder Chapters:
  * Use the Up/Down buttons or press Ctrl+Up / Ctrl+Down.
- In-Line Chapter Renaming:
  * Double-click any chapter (or press F2) to rename it directly in-line. Press Enter to commit or Escape to cancel.
  * Hover over any chapter name to view the original source filename.
  * Duration appending can be configured in Edit -> Settings -> Audiobook ('Append duration to chapter name', default: On).
- Cover Art & Bitrate:
  * Browse or download an Instagram cover, select your desired AAC bitrate (e.g. 128k).
  * The default artist for new audiobooks is customisable in Settings -> Audiobook (default: 'Various Artists').
- Build:
  * Click 'Build Audiobook' (or press Ctrl+B). ComComTag will transcode and chapterise the tracks with FFmpeg.
  * You can cancel the build at any time by clicking 'Cancel Build'.";
        }

        private string GetNamingPatternsHelpText() {
            return 
@"===================================================================
Custom Naming Conventions & Patterns
===================================================================

Configure your naming conventions anytime in Edit -> Settings -> Conventions.

-------------------------------------------------------------------
Supported Placeholders:
-------------------------------------------------------------------
- {Date}      : Formatted recording date (e.g., 1999-06-13).
- {Artist}    : Artist or Performer name as entered.
- {Show}      : Show, Event, or Album name.
- {Location}  : Venue or city location.
- {Venue}     : Alias for {Location}.
- {Track}     : Track identifier or number (used in multi-file tagging).

-------------------------------------------------------------------
Case Sensitivity Transformations:
-------------------------------------------------------------------
- {Artist}    : Preserves original capitalisation (e.g. 'Wally Bazoom').
- {ARTIST}    : Transforms into ALL UPPERCASE (e.g. 'WALLY BAZOOM').
- {artist}    : Transforms into all lowercase (e.g. 'wally bazoom').
- The same rule applies to {SHOW}/{show} and {LOCATION}/{location}.

-------------------------------------------------------------------
Smart Delimiter Cleanup:
-------------------------------------------------------------------
- If an optional field (like Show) is blank, ComComTag automatically eliminates empty parentheses () and duplicate separators (e.g. ' -  - ' becomes ' - ').
- Trailing and leading separators are automatically cleaned up.";
        }

        private string GetInstagramHelpText() {
            return 
@"===================================================================
Instagram Cover Art Retrieval
===================================================================

ComComTag includes a 100% native C# Instagram scraper requiring zero Python or external runtime dependencies.

-------------------------------------------------------------------
How to Download Artwork:
-------------------------------------------------------------------
1. Click the Instagram camera icon button beside the Cover Art box.
2. Paste any public Instagram post URL or Reel link, e.g.:
   https://www.instagram.com/p/Dcon0V8m71n/
3. ComComTag connects directly to Instagram:
   - Tier 1 (Polaris GraphQL): Extracts uncropped 1440p/1080p carousel slides.
   - Tier 2 (Meta oEmbed): Guaranteed uncropped square thumbnail fallback.
   - Tier 3 (Crawler Meta): High-resolution fallback parser.

-------------------------------------------------------------------
Multi-Image Carousel Picker:
-------------------------------------------------------------------
- If the post contains multiple slides, the Carousel Picker modal opens automatically.
- Use '< Previous' and 'Next >' buttons or the Left / Right arrow keys to view each image.
- Click 'Select This Cover Art' (or press Enter) to apply that image.
- Double-clicking the cover box in ComComTag re-opens the carousel modal at any time.

-------------------------------------------------------------------
Temp Folder Safety:
-------------------------------------------------------------------
- Temporary downloads are kept in %TEMP%\ComComTag_Insta (or your custom directory set in Settings).
- When 'Flush temporary Instagram files on exit' is enabled, ComComTag cleans up only its own downloaded temporary shortcode subfolders. Custom user files are never deleted.";
        }

        private string GetShortcutsHelpText() {
            return 
@"===================================================================
Keyboard Shortcuts & Quick Actions
===================================================================

General:
- F1               : Open this User Guide & Help dialogue.
- F5               : Refresh the active folder and reload MP3 files.
- Drag & Drop      : Drag folders from Windows Explorer into the file list.

Tag & Rename MP3 (Tab 1):
- Ctrl + S         : Save ID3 tags and rename selected files on disk.

Audiobook Builder (Tab 2):
- Ctrl + B         : Build and compile Audiobook (M4B).
- F2 / Double-Click: Rename selected chapter in-place.
- Enter (in Edit)  : Save in-line chapter name change.
- Esc (in Edit)    : Cancel in-line chapter rename.
- Double-Click MP3 : Instantly transfer track from Available MP3s to Chapters.
- Enter (in MP3s)  : Add selected MP3s to Chapters.
- Delete / Backspace: Remove selected chapters.
- Ctrl + Up        : Move selected chapter up.
- Ctrl + Down      : Move selected chapter down.

Cover Art & Carousel:
- Left / Right Arrow: Navigate carousel slides in the Carousel Picker modal.
- Enter            : Confirm and select active carousel slide.
- Double-Click Pic : Open the Carousel Picker modal (if post has multiple images).
- Drag & Drop Pic  : Drag any image file (.jpg, .png, .webp, .bmp) onto the cover box.";
        }

        private string GetAboutHelpText() {
            string version = Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
            return 
string.Format(
@"===================================================================
About ComComTag v{0}
===================================================================

ComComTag is a lightweight, portable Windows audio tagging and audiobook creation utility crafted for comedy archivists, podcast managers, and audio collectors.

Key Features:
- Standardised MP3 ID3v2 tagging with customisable naming conventions.
- In-memory TagLib# duration probing (400x speedup).
- Native C# Instagram artwork scraping & visual multi-image carousel.
- High-speed M4B audiobook generation with FFMETADATA1 chapter markers.
- Single standalone executable with embedded TagLibSharp assembly.

Technology Stack:
- C# (.NET Framework 4.0+ / Windows Forms)
- TagLib# (ID3 & MP4 Metadata Reader/Writer)
- FFmpeg (Audio Transcoding & AAC/M4B Muxing)

Copyright © 2026 ComComTag Project. All rights reserved.", version);
        }
    }
}
