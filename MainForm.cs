using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ComComTag {
    public class MainForm : Form {
        private Settings _settings;
        private string _selectedFolder = "";

        // Menu Bar
        private MenuStrip menuStrip;

        // Tabs
        private TabControl tabControl;
        private TabPage tabTagging;
        private TabPage tabAudiobook;

        // Shared / Main
        private Button btnBrowse;
        private Label lblCurrentFolder;

        // --- Tagging Tab ---
        private ListBox listTagFiles;
        private Label lblDate;
        private DateTimePicker dtpDate;
        private Label lblYear;
        private TextBox txtYear;
        private Label lblArtist;
        private TextBox txtArtist;
        private Label lblShow;
        private TextBox txtShow;
        private Label lblLocation;
        private ComboBox cmbLocation;
        private Button btnCopyTags;
        private Button btnPasteTags;
        private Button btnSaveRename;

        // Tag clipboard
        private bool _hasClipboardTags = false;
        private DateTime _clipDate;
        private string _clipArtist;
        private string _clipShow;
        private string _clipLocation;
        
        // --- Tagging Tab Images ---
        private TextBox txtTagCover;
        private Button btnTagBrowseCover;
        private Button btnTagInstaDownload;
        private PictureBox picTagCoverPreview;
        private Button btnTagInstaPrev;
        private Button btnTagInstaNext;
        
        private Label lblTitlePreview;
        private List<string> _tagInstaImages = new List<string>();
        private int _tagInstaImageIndex = 0;

        // --- Audiobook Tab ---
        private ListBox listAvailableMp3s;
        private Button btnAddChapter;
        private Button btnRemoveChapter;
        private Button btnMoveUp;
        private Button btnMoveDown;
        
        // Chapters list contains custom objects to hold file path + custom name
        private ListBox listChapters;
        private TextBox txtChapterName;
        private Button btnUpdateChapterName;
        private CheckBox chkAppendDuration;

        private DateTimePicker dtpM4bDate;
        private TextBox txtM4bYear;
        private TextBox txtM4bArtist;
        private TextBox txtM4bAlbum;
        private ComboBox cmbM4bLocation;
        private TextBox txtCover;
        private Button btnBrowseCover;
        private Button btnInstaDownload;
        private PictureBox picCoverPreview;
        
        private List<string> _instaImages = new List<string>();
        private int _instaImageIndex = 0;
        private Button btnInstaPrev;
        private Button btnInstaNext;
        
        private ComboBox cmbBitrate;
        private ProgressBar pbM4bProgress;
        
        private Button btnBuildM4b;

        public MainForm() {
            _settings = new Settings();
            
            // Set Taskbar Icon from embedded resource
            try {
                using (Stream iconStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ComComTag.icon.ico")) {
                    if (iconStream != null) {
                        this.Icon = new Icon(iconStream);
                    }
                }
            } catch { }

            InitializeComponent();
            CheckFFmpegOnStartup();
        }

        private void CheckFFmpegOnStartup() {
            string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string fullFfmpeg = _settings.FFmpegPath;
            if (!Path.IsPathRooted(fullFfmpeg)) {
                fullFfmpeg = Path.Combine(exePath, fullFfmpeg);
            }
            if (!File.Exists(fullFfmpeg)) {
                MessageBox.Show("ffmpeg.exe was not found. The 'Build Audiobook' feature will not work correctly until you configure it in Edit -> Settings.", "FFmpeg Missing", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void InitializeComponent() {
            this.Text = "ComComTag";
            this.Size = new Size(850, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // --- Menu Bar ---
            menuStrip = new MenuStrip();
            ToolStripMenuItem menuFile = new ToolStripMenuItem("File");
            ToolStripMenuItem menuBrowse = new ToolStripMenuItem("Open Folder...");
            menuBrowse.Click += BtnBrowse_Click;
            ToolStripMenuItem menuExit = new ToolStripMenuItem("Exit");
            menuExit.Click += (s, e) => this.Close();
            menuFile.DropDownItems.Add(menuBrowse);
            menuFile.DropDownItems.Add(new ToolStripSeparator());
            menuFile.DropDownItems.Add(menuExit);

            ToolStripMenuItem menuEdit = new ToolStripMenuItem("Edit");
            ToolStripMenuItem menuSettings = new ToolStripMenuItem("Settings");
            menuSettings.Click += MenuSettings_Click;
            menuEdit.DropDownItems.Add(menuSettings);

            menuStrip.Items.Add(menuFile);
            menuStrip.Items.Add(menuEdit);
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);

            // --- Top Folder Selection ---
            btnBrowse = new Button { Location = new Point(15, 35), Size = new Size(120, 30), Text = "Browse Folder..." };
            btnBrowse.Click += BtnBrowse_Click;
            lblCurrentFolder = new Label { Location = new Point(145, 42), Size = new Size(650, 20), Text = "No folder selected." };
            this.Controls.Add(btnBrowse);
            this.Controls.Add(lblCurrentFolder);

            // --- Tabs ---
            tabControl = new TabControl { Location = new Point(15, 75), Size = new Size(800, 520) };
            tabTagging = new TabPage("Tag & Rename MP3");
            tabAudiobook = new TabPage("Build Audiobook (M4B)");
            tabControl.TabPages.Add(tabTagging);
            tabControl.TabPages.Add(tabAudiobook);
            this.Controls.Add(tabControl);

            InitializeTaggingTab();
            InitializeAudiobookTab();

            if (!string.IsNullOrEmpty(_settings.DefaultDirectory) && Directory.Exists(_settings.DefaultDirectory)) {
                LoadDirectory(_settings.DefaultDirectory);
            }
        }

        private void MenuSettings_Click(object sender, EventArgs e) {
            using (var sf = new SettingsForm(_settings)) {
                if (sf.ShowDialog() == DialogResult.OK) {
                    cmbLocation.Items.Clear();
                    cmbLocation.Items.AddRange(_settings.Locations.ToArray());
                    cmbM4bLocation.Items.Clear();
                    cmbM4bLocation.Items.AddRange(_settings.Locations.ToArray());
                }
            }
        }

        private void InitializeTaggingTab() {
            listTagFiles = new ListBox { Location = new Point(15, 15), Size = new Size(300, 420), SelectionMode = SelectionMode.MultiExtended };
            listTagFiles.SelectedIndexChanged += ListTagFiles_SelectedIndexChanged;

            int xOffset = 340;
            int yOffset = 15;
            int labelWidth = 60;
            int inputWidth = 250;
            int inputX = xOffset + labelWidth;

            Label lblHelp = new Label { Location = new Point(xOffset, yOffset), Size = new Size(400, 20), Text = "Select one or more MP3s to tag and rename.", Font = new Font(this.Font, FontStyle.Italic) };

            yOffset += 35;
            lblDate = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Date:" };
            dtpDate = new DateTimePicker { Location = new Point(inputX, yOffset), Size = new Size(inputWidth, 20), Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtpDate.ValueChanged += DtpDate_ValueChanged;
            dtpDate.ValueChanged += UpdateTitlePreview;

            yOffset += 35;
            lblYear = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Year:" };
            txtYear = new TextBox { Location = new Point(inputX, yOffset), Size = new Size(inputWidth, 20), ReadOnly = true, BackColor = SystemColors.Control, TabStop = false };
            txtYear.Text = dtpDate.Value.Year.ToString();

            yOffset += 35;
            lblArtist = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Artist:" };
            txtArtist = new TextBox { Location = new Point(inputX, yOffset), Size = new Size(inputWidth, 20) };
            txtArtist.TextChanged += UpdateTitlePreview;

            yOffset += 35;
            lblShow = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Show:" };
            txtShow = new TextBox { Location = new Point(inputX, yOffset), Size = new Size(inputWidth, 20) };
            txtShow.TextChanged += UpdateTitlePreview;

            yOffset += 35;
            lblLocation = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Location:" };
            cmbLocation = new ComboBox { Location = new Point(inputX, yOffset), Size = new Size(inputWidth, 20), DropDownStyle = ComboBoxStyle.DropDown };
            cmbLocation.Items.AddRange(_settings.Locations.ToArray());
            cmbLocation.TextChanged += UpdateTitlePreview;

            yOffset += 40;
            Label lblTagCov = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Cover:" };
            txtTagCover = new TextBox { Location = new Point(inputX, yOffset - 2), Size = new Size(195, 20) };
            txtTagCover.TextChanged += TxtTagCover_TextChanged;
            
            btnTagBrowseCover = new Button { Location = new Point(inputX + 205, yOffset - 4), Size = new Size(30, 24), Text = "..." };
            btnTagBrowseCover.Click += BtnTagBrowseCover_Click;
            
            btnTagInstaDownload = new Button { Location = new Point(inputX + 240, yOffset - 4), Size = new Size(110, 24), Text = "Paste Insta URL" };
            btnTagInstaDownload.Click += BtnTagInstaDownload_Click;
            
            yOffset += 30;
            picTagCoverPreview = new PictureBox { Location = new Point(inputX, yOffset), Size = new Size(100, 100), SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle };

            btnTagInstaPrev = new Button { Location = new Point(inputX + 110, yOffset + 35), Size = new Size(25, 25), Text = "<", Visible = false };
            btnTagInstaPrev.Click += BtnTagInstaPrev_Click;
            
            btnTagInstaNext = new Button { Location = new Point(inputX + 145, yOffset + 35), Size = new Size(25, 25), Text = ">", Visible = false };
            btnTagInstaNext.Click += BtnTagInstaNext_Click;

            yOffset += 110;
            Label lblPreviewPrompt = new Label { Location = new Point(xOffset, yOffset), Size = new Size(labelWidth, 20), Text = "Preview:" };
            lblTitlePreview = new Label { Location = new Point(inputX, yOffset), Size = new Size(350, 20), ForeColor = SystemColors.GrayText, Text = "" };

            yOffset += 30;
            btnCopyTags = new Button { Location = new Point(inputX, yOffset), Size = new Size(100, 30), Text = "Copy Tags", Enabled = false };
            btnCopyTags.Click += BtnCopyTags_Click;
            btnPasteTags = new Button { Location = new Point(inputX + 110, yOffset), Size = new Size(100, 30), Text = "Paste Tags", Enabled = false };
            btnPasteTags.Click += BtnPasteTags_Click;

            yOffset += 45;
            btnSaveRename = new Button { Location = new Point(inputX + 60, yOffset), Size = new Size(150, 40), Text = "Save && Rename", Enabled = false };
            btnSaveRename.Click += BtnSaveRename_Click;

            tabTagging.Controls.Add(listTagFiles);
            tabTagging.Controls.Add(lblHelp);
            tabTagging.Controls.Add(lblDate);
            tabTagging.Controls.Add(dtpDate);
            tabTagging.Controls.Add(lblYear);
            tabTagging.Controls.Add(txtYear);
            tabTagging.Controls.Add(lblArtist);
            tabTagging.Controls.Add(txtArtist);
            tabTagging.Controls.Add(lblShow);
            tabTagging.Controls.Add(txtShow);
            tabTagging.Controls.Add(lblLocation);
            tabTagging.Controls.Add(cmbLocation);
            tabTagging.Controls.Add(lblTagCov);
            tabTagging.Controls.Add(txtTagCover);
            tabTagging.Controls.Add(btnTagBrowseCover);
            tabTagging.Controls.Add(btnTagInstaDownload);
            tabTagging.Controls.Add(picTagCoverPreview);
            tabTagging.Controls.Add(btnTagInstaPrev);
            tabTagging.Controls.Add(btnTagInstaNext);
            
            tabTagging.Controls.Add(lblPreviewPrompt);
            tabTagging.Controls.Add(lblTitlePreview);
            tabTagging.Controls.Add(btnCopyTags);
            tabTagging.Controls.Add(btnPasteTags);
            tabTagging.Controls.Add(btnSaveRename);
        }

        private void DtpDate_ValueChanged(object sender, EventArgs e) {
            txtYear.Text = dtpDate.Value.Year.ToString();
        }

        private void InitializeAudiobookTab() {
            Label lblAvail = new Label { Location = new Point(15, 15), Size = new Size(200, 20), Text = "1. Available MP3s:" };
            listAvailableMp3s = new ListBox { Location = new Point(15, 35), Size = new Size(250, 200), SelectionMode = SelectionMode.MultiExtended };
            
            btnAddChapter = new Button { Location = new Point(275, 100), Size = new Size(40, 30), Text = ">>" };
            btnAddChapter.Click += BtnAddChapter_Click;
            btnRemoveChapter = new Button { Location = new Point(275, 140), Size = new Size(40, 30), Text = "<<" };
            btnRemoveChapter.Click += BtnRemoveChapter_Click;

            Label lblChaps = new Label { Location = new Point(325, 15), Size = new Size(200, 20), Text = "2. Audiobook Chapters:" };
            listChapters = new ListBox { Location = new Point(325, 35), Size = new Size(250, 200) };
            listChapters.SelectedIndexChanged += ListChapters_SelectedIndexChanged;
            
            btnMoveUp = new Button { Location = new Point(585, 100), Size = new Size(60, 30), Text = "Up" };
            btnMoveUp.Click += BtnMoveUp_Click;
            btnMoveDown = new Button { Location = new Point(585, 140), Size = new Size(60, 30), Text = "Down" };
            btnMoveDown.Click += BtnMoveDown_Click;

            // Chapter Naming
            Label lblChapName = new Label { Location = new Point(325, 245), Size = new Size(100, 20), Text = "Rename Chapter:" };
            txtChapterName = new TextBox { Location = new Point(425, 242), Size = new Size(150, 20), Enabled = false };
            btnUpdateChapterName = new Button { Location = new Point(585, 240), Size = new Size(60, 25), Text = "Update", Enabled = false };
            btnUpdateChapterName.Click += BtnUpdateChapterName_Click;
            
            chkAppendDuration = new CheckBox { Location = new Point(325, 265), Size = new Size(250, 20), Text = "Append Duration (mm:ss) to Title" };
            
            // Global settings
            int yOffset = 295;
            Label lblMeta = new Label { Location = new Point(15, yOffset), Size = new Size(200, 20), Text = "3. Audiobook Metadata:", Font = new Font(this.Font, FontStyle.Bold) };
            
            yOffset += 28;
            Label lblM4bDateLbl = new Label { Location = new Point(15, yOffset), Size = new Size(50, 20), Text = "Date:" };
            dtpM4bDate = new DateTimePicker { Location = new Point(70, yOffset - 2), Size = new Size(195, 20), Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtpM4bDate.ValueChanged += DtpM4bDate_ValueChanged;

            Label lblM4bYearLbl = new Label { Location = new Point(325, yOffset), Size = new Size(50, 20), Text = "Year:" };
            txtM4bYear = new TextBox { Location = new Point(380, yOffset - 2), Size = new Size(195, 20) };
            txtM4bYear.Text = dtpM4bDate.Value.Year.ToString();

            yOffset += 30;
            Label lblArt = new Label { Location = new Point(15, yOffset), Size = new Size(50, 20), Text = "Artist:" };
            txtM4bArtist = new TextBox { Location = new Point(70, yOffset - 2), Size = new Size(195, 20), Text = "Various" };

            Label lblAlb = new Label { Location = new Point(325, yOffset), Size = new Size(50, 20), Text = "Album:" };
            txtM4bAlbum = new TextBox { Location = new Point(380, yOffset - 2), Size = new Size(195, 20) };

            yOffset += 30;
            Label lblM4bLocLbl = new Label { Location = new Point(15, yOffset), Size = new Size(50, 20), Text = "Location:" };
            cmbM4bLocation = new ComboBox { Location = new Point(70, yOffset - 2), Size = new Size(195, 20), DropDownStyle = ComboBoxStyle.DropDown };
            cmbM4bLocation.Items.AddRange(_settings.Locations.ToArray());

            Label lblBitrate = new Label { Location = new Point(325, yOffset), Size = new Size(50, 20), Text = "Bitrate:" };
            cmbBitrate = new ComboBox { Location = new Point(380, yOffset - 2), Size = new Size(195, 20), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbBitrate.Items.AddRange(new object[] { "64k", "96k", "128k", "192k", "256k", "320k" });
            if (!string.IsNullOrEmpty(_settings.DefaultBitrate) && cmbBitrate.Items.Contains(_settings.DefaultBitrate)) {
                cmbBitrate.SelectedIndex = cmbBitrate.Items.IndexOf(_settings.DefaultBitrate);
            } else {
                cmbBitrate.SelectedIndex = 2; // Default 128k
            }

            yOffset += 30;
            Label lblCov = new Label { Location = new Point(15, yOffset), Size = new Size(50, 20), Text = "Cover:" };
            txtCover = new TextBox { Location = new Point(70, yOffset - 2), Size = new Size(195, 20) };
            txtCover.TextChanged += TxtCover_TextChanged;
            btnBrowseCover = new Button { Location = new Point(275, yOffset - 4), Size = new Size(30, 24), Text = "..." };
            btnBrowseCover.Click += BtnBrowseCover_Click;
            
            btnInstaDownload = new Button { Location = new Point(310, yOffset - 4), Size = new Size(110, 24), Text = "Paste Insta URL" };
            btnInstaDownload.Click += BtnInstaDownload_Click;
            
            // Move Build Button here
            btnBuildM4b = new Button { Location = new Point(585, 308), Size = new Size(195, 45), Text = "Build Audiobook (M4B)", Font = new Font(this.Font, FontStyle.Bold) };
            btnBuildM4b.Click += BtnBuildM4b_Click;
            
            picCoverPreview = new PictureBox { Location = new Point(585, 360), Size = new Size(80, 80), SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle };

            btnInstaPrev = new Button { Location = new Point(555, 390), Size = new Size(25, 25), Text = "<", Visible = false };
            btnInstaPrev.Click += BtnInstaPrev_Click;
            btnInstaNext = new Button { Location = new Point(670, 390), Size = new Size(25, 25), Text = ">", Visible = false };
            btnInstaNext.Click += BtnInstaNext_Click;

            yOffset = 445;
            pbM4bProgress = new ProgressBar { Location = new Point(15, yOffset), Size = new Size(770, 30), Minimum = 0, Maximum = 100 };
            

            tabAudiobook.Controls.Add(lblAvail);
            tabAudiobook.Controls.Add(listAvailableMp3s);
            tabAudiobook.Controls.Add(btnAddChapter);
            tabAudiobook.Controls.Add(btnRemoveChapter);
            tabAudiobook.Controls.Add(lblChaps);
            tabAudiobook.Controls.Add(listChapters);
            tabAudiobook.Controls.Add(btnMoveUp);
            tabAudiobook.Controls.Add(btnMoveDown);
            tabAudiobook.Controls.Add(lblChapName);
            tabAudiobook.Controls.Add(txtChapterName);
            tabAudiobook.Controls.Add(btnUpdateChapterName);
            tabAudiobook.Controls.Add(chkAppendDuration);

            tabAudiobook.Controls.Add(lblMeta);
            tabAudiobook.Controls.Add(lblM4bDateLbl);
            tabAudiobook.Controls.Add(dtpM4bDate);
            tabAudiobook.Controls.Add(lblM4bYearLbl);
            tabAudiobook.Controls.Add(txtM4bYear);
            tabAudiobook.Controls.Add(lblArt);
            tabAudiobook.Controls.Add(txtM4bArtist);
            tabAudiobook.Controls.Add(lblAlb);
            tabAudiobook.Controls.Add(txtM4bAlbum);
            tabAudiobook.Controls.Add(lblM4bLocLbl);
            tabAudiobook.Controls.Add(cmbM4bLocation);
            tabAudiobook.Controls.Add(lblCov);
            tabAudiobook.Controls.Add(txtCover);
            tabAudiobook.Controls.Add(btnBrowseCover);
            tabAudiobook.Controls.Add(btnInstaDownload);
            tabAudiobook.Controls.Add(picCoverPreview);
            tabAudiobook.Controls.Add(btnInstaPrev);
            tabAudiobook.Controls.Add(btnInstaNext);
            tabAudiobook.Controls.Add(lblBitrate);
            tabAudiobook.Controls.Add(cmbBitrate);
            tabAudiobook.Controls.Add(btnBuildM4b);
            tabAudiobook.Controls.Add(pbM4bProgress);
        }

        private void DtpM4bDate_ValueChanged(object sender, EventArgs e) {
            txtM4bYear.Text = dtpM4bDate.Value.Year.ToString();
        }

        private void BtnBrowse_Click(object sender, EventArgs e) {
            using (var dialog = new FolderBrowserDialog()) {
                dialog.Description = "Select the folder containing MP3 files";
                if (!string.IsNullOrEmpty(_selectedFolder)) dialog.SelectedPath = _selectedFolder;
                
                if (dialog.ShowDialog() == DialogResult.OK) {
                    LoadDirectory(dialog.SelectedPath);
                    _settings.DefaultDirectory = dialog.SelectedPath;
                    _settings.Save();
                }
            }
        }

        private void LoadDirectory(string path) {
            _selectedFolder = path;
            lblCurrentFolder.Text = path;
            
            listTagFiles.Items.Clear();
            listAvailableMp3s.Items.Clear();
            listChapters.Items.Clear();

            var mp3s = Directory.GetFiles(_selectedFolder, "*.mp3").OrderBy(f => f).ToArray();
            foreach (var file in mp3s) {
                string name = Path.GetFileName(file);
                listTagFiles.Items.Add(name);
                listAvailableMp3s.Items.Add(name);
            }
            btnSaveRename.Enabled = false;
            btnCopyTags.Enabled = false;
        }

        // --- Tagging Logic ---
        // Note: In multi-select, this fires sequentially for each selected item.
        // The fields are intentionally populated from the last-selected file,
        // which means "Copy Tags" will copy the tags from that final item.
        private void ListTagFiles_SelectedIndexChanged(object sender, EventArgs e) {
            if (listTagFiles.SelectedIndex >= 0) {
                btnSaveRename.Enabled = true;
                string fileName = listTagFiles.SelectedItem.ToString();
                string fullPath = Path.Combine(_selectedFolder, fileName);
                
                // Clear fields first so stale data doesn't persist
                txtArtist.Text = "";
                txtShow.Text = "";
                cmbLocation.Text = "";
                txtTagCover.Text = "";

                bool tagsLoaded = false;
                bool hasTags = false;
                
                try {
                    using (TagLib.File tagFile = TagLib.File.Create(fullPath)) {
                        string tArtist = tagFile.Tag.FirstPerformer;
                        string tAlbum = tagFile.Tag.Album;
                        string tTitle = tagFile.Tag.Title;
                        uint tYear = tagFile.Tag.Year;

                        if (!string.IsNullOrEmpty(tArtist)) {
                            txtArtist.Text = tArtist;
                            hasTags = true;
                        }
                        if (!string.IsNullOrEmpty(tAlbum)) {
                            txtShow.Text = tAlbum;
                            hasTags = true;
                        }
                        
                        // Try to parse title for date and location if it follows "YYYY-MM-DD - Location"
                        if (!string.IsNullOrEmpty(tTitle)) {
                            var parts = tTitle.Split(new[] { " - " }, 2, StringSplitOptions.None);
                            if (parts.Length == 2) {
                                DateTime d;
                                if (DateTime.TryParse(parts[0], out d)) dtpDate.Value = d;
                                cmbLocation.Text = parts[1];
                            }
                            hasTags = true;
                        }

                        // Populate year from the Date tag's year if available
                        if (tYear > 0) {
                            try {
                                dtpDate.Value = new DateTime((int)tYear, dtpDate.Value.Month, dtpDate.Value.Day);
                            } catch { }
                        }
                        tagsLoaded = true;
                    }
                } catch { }

                // If no tags were loaded, try regex on filename
                if (!tagsLoaded) {
                    var match = System.Text.RegularExpressions.Regex.Match(fileName, @"Voice\s+(?<YY>\d{2})(?<MM>\d{2})(?<DD>\d{2})_");
                    if (match.Success) {
                        int year = int.Parse("20" + match.Groups["YY"].Value);
                        int month = int.Parse(match.Groups["MM"].Value);
                        int day = int.Parse(match.Groups["DD"].Value);
                        try {
                            dtpDate.Value = new DateTime(year, month, day);
                        } catch { }
                    }
                }

                // Also try to parse the filename if it follows our naming convention
                // yyyy-mm-dd - artist - show - location.mp3
                if (!hasTags) {
                    string baseName = Path.GetFileNameWithoutExtension(fileName);
                    var fparts = baseName.Split(new[] { " - " }, StringSplitOptions.None);
                    if (fparts.Length >= 4) {
                        DateTime d;
                        if (DateTime.TryParse(fparts[0], out d)) {
                            dtpDate.Value = d;
                        }
                        txtArtist.Text = fparts[1];
                        txtShow.Text = fparts[2];
                        cmbLocation.Text = fparts[3];
                        hasTags = true;
                    }
                }

                btnCopyTags.Enabled = hasTags;
                UpdateTitlePreview(null, null);
            } else {
                btnSaveRename.Enabled = false;
                btnCopyTags.Enabled = false;
                UpdateTitlePreview(null, null);
            }
        }
        
        private void UpdateTitlePreview(object sender, EventArgs e) {
            if (listTagFiles.SelectedItems.Count == 0) {
                lblTitlePreview.Text = "";
                return;
            }
            
            string dateStr = dtpDate.Value.ToString("yyyy-MM-dd");
            string artist = txtArtist.Text.Trim();
            string show = txtShow.Text.Trim();
            string location = cmbLocation.Text.Trim();
            
            if (string.IsNullOrWhiteSpace(artist)) artist = "[Artist]";
            if (string.IsNullOrWhiteSpace(show)) show = "[Show]";
            if (string.IsNullOrWhiteSpace(location)) location = "[Location]";

            if (listTagFiles.SelectedItems.Count > 1) {
                lblTitlePreview.Text = string.Format("{0} - {1} - {2} - {3} (original_name).mp3", dateStr, artist, show, location);
            } else {
                lblTitlePreview.Text = string.Format("{0} - {1} - {2} - {3}.mp3", dateStr, artist, show, location);
            }
        }

        private void BtnCopyTags_Click(object sender, EventArgs e) {
            _clipDate = dtpDate.Value;
            _clipArtist = txtArtist.Text;
            _clipShow = txtShow.Text;
            _clipLocation = cmbLocation.Text;
            _hasClipboardTags = true;
            btnPasteTags.Enabled = true;
        }

        private void BtnPasteTags_Click(object sender, EventArgs e) {
            if (!_hasClipboardTags) return;
            dtpDate.Value = _clipDate;
            txtArtist.Text = _clipArtist;
            txtShow.Text = _clipShow;
            cmbLocation.Text = _clipLocation;
        }

        private void BtnSaveRename_Click(object sender, EventArgs e) {
            if (listTagFiles.SelectedItems.Count == 0) return;

            string dateStr = dtpDate.Value.ToString("yyyy-MM-dd");
            string artist = txtArtist.Text.Trim();
            string show = txtShow.Text.Trim();
            string location = cmbLocation.Text.Trim();
            uint yearValue = (uint)dtpDate.Value.Year;

            if (string.IsNullOrWhiteSpace(artist) || string.IsNullOrWhiteSpace(show) || string.IsNullOrWhiteSpace(location)) {
                MessageBox.Show("Please fill out Artist, Show, and Location.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!_settings.Locations.Contains(location)) {
                _settings.Locations.Add(location);
                _settings.Save();
                cmbLocation.Items.Add(location);
            }

            List<string> selectedNames = new List<string>();
            foreach (var item in listTagFiles.SelectedItems) selectedNames.Add(item.ToString());

            foreach (var fileName in selectedNames) {
                string sourcePath = Path.Combine(_selectedFolder, fileName);
                
                string newName = selectedNames.Count > 1 
                                 ? string.Format("{0} - {1} - {2} - {3} ({4}).mp3", dateStr, artist, show, location, Path.GetFileNameWithoutExtension(fileName))
                                 : string.Format("{0} - {1} - {2} - {3}.mp3", dateStr, artist, show, location);

                string destPath = Path.Combine(_selectedFolder, newName);

                try {
                    // Overwrite protection
                    if (!string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase)) {
                         if (File.Exists(destPath)) {
                             DialogResult dr = MessageBox.Show(
                                 string.Format("The file '{0}' already exists. Overwrite it?", newName), 
                                 "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                             if (dr == DialogResult.No) continue;
                             File.Delete(destPath);
                         }
                         File.Move(sourcePath, destPath);
                    }
                    
                    using (TagLib.File tagFile = TagLib.File.Create(destPath)) {
                        tagFile.Tag.Title = string.Format("{0} - {1}", dateStr, location);
                        tagFile.Tag.Performers = new[] { artist };
                        tagFile.Tag.Album = show;
                        tagFile.Tag.Year = yearValue;
                        
                        string cvr = txtTagCover.Text.Trim();
                        if (!string.IsNullOrEmpty(cvr) && File.Exists(cvr)) {
                            TagLib.Picture pic = new TagLib.Picture(cvr);
                            tagFile.Tag.Pictures = new TagLib.IPicture[] { pic };
                        }
                        
                        tagFile.Save();
                    }
                } catch (Exception ex) {
                    MessageBox.Show(string.Format("Error processing {0}: {1}", fileName, ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            LoadDirectory(_selectedFolder);
            MessageBox.Show("Files saved and renamed!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // --- Audiobook Logic ---
        public class ChapterItem {
            public string FilePath { get; set; }
            public string ChapterName { get; set; }
            public override string ToString() {
                return ChapterName;
            }
        }

        private void BtnAddChapter_Click(object sender, EventArgs e) {
            foreach (var item in listAvailableMp3s.SelectedItems) {
                string fn = item.ToString();
                string fullPath = Path.Combine(_selectedFolder, fn);
                string chapterName = Path.GetFileNameWithoutExtension(fn);
                
                if (chkAppendDuration.Checked) {
                    try {
                        using (TagLib.File tagFile = TagLib.File.Create(fullPath)) {
                            TimeSpan d = tagFile.Properties.Duration;
                            chapterName += string.Format(" ({0:D2}:{1:D2})", (int)d.TotalMinutes, d.Seconds);
                        }
                    } catch { }
                }
                
                listChapters.Items.Add(new ChapterItem {
                    FilePath = fullPath,
                    ChapterName = chapterName
                });
            }
            if (string.IsNullOrWhiteSpace(txtM4bAlbum.Text)) txtM4bAlbum.Text = txtShow.Text;
        }

        private void BtnRemoveChapter_Click(object sender, EventArgs e) {
            while (listChapters.SelectedItems.Count > 0) {
                listChapters.Items.Remove(listChapters.SelectedItems[0]);
            }
            txtChapterName.Enabled = false;
            btnUpdateChapterName.Enabled = false;
        }

        private void BtnMoveUp_Click(object sender, EventArgs e) {
            if (listChapters.SelectedItem == null || listChapters.SelectedIndex <= 0) return;
            int newIndex = listChapters.SelectedIndex - 1;
            object selected = listChapters.SelectedItem;
            listChapters.Items.Remove(selected);
            listChapters.Items.Insert(newIndex, selected);
            listChapters.SetSelected(newIndex, true);
        }

        private void BtnMoveDown_Click(object sender, EventArgs e) {
            if (listChapters.SelectedItem == null || listChapters.SelectedIndex >= listChapters.Items.Count - 1) return;
            int newIndex = listChapters.SelectedIndex + 1;
            object selected = listChapters.SelectedItem;
            listChapters.Items.Remove(selected);
            listChapters.Items.Insert(newIndex, selected);
            listChapters.SetSelected(newIndex, true);
        }

        private void ListChapters_SelectedIndexChanged(object sender, EventArgs e) {
            ChapterItem ci = listChapters.SelectedItem as ChapterItem;
            if (ci != null) {
                txtChapterName.Enabled = true;
                btnUpdateChapterName.Enabled = true;
                txtChapterName.Text = ci.ChapterName;
            } else {
                txtChapterName.Enabled = false;
                btnUpdateChapterName.Enabled = false;
            }
        }

        private void BtnUpdateChapterName_Click(object sender, EventArgs e) {
            ChapterItem ci = listChapters.SelectedItem as ChapterItem;
            if (ci != null) {
                ci.ChapterName = txtChapterName.Text.Trim();
                int idx = listChapters.SelectedIndex;
                listChapters.Items[idx] = ci;
                listChapters.SetSelected(idx, true);
            }
        }

        private void BtnInstaPrev_Click(object sender, EventArgs e) {
            if (_instaImages.Count > 0) {
                _instaImageIndex = (_instaImageIndex - 1 + _instaImages.Count) % _instaImages.Count;
                txtCover.Text = _instaImages[_instaImageIndex];
            }
        }

        private void BtnInstaNext_Click(object sender, EventArgs e) {
            if (_instaImages.Count > 0) {
                _instaImageIndex = (_instaImageIndex + 1) % _instaImages.Count;
                txtCover.Text = _instaImages[_instaImageIndex];
            }
        }

        private void BtnBrowseCover_Click(object sender, EventArgs e) {
            using (var dialog = new OpenFileDialog { Filter = "Image Files|*.jpg;*.jpeg;*.png" }) {
                if (dialog.ShowDialog() == DialogResult.OK) {
                    txtCover.Text = dialog.FileName;
                }
            }
        }
        
        private void BtnTagBrowseCover_Click(object sender, EventArgs e) {
            using (var dialog = new OpenFileDialog { Filter = "Image Files|*.jpg;*.jpeg;*.png" }) {
                if (dialog.ShowDialog() == DialogResult.OK) {
                    txtTagCover.Text = dialog.FileName;
                }
            }
        }
        
        private void TxtTagCover_TextChanged(object sender, EventArgs e) {
            string path = txtTagCover.Text.Trim();
            if (File.Exists(path)) {
                try {
                    picTagCoverPreview.ImageLocation = path;
                } catch {
                    picTagCoverPreview.Image = null;
                }
            } else {
                picTagCoverPreview.Image = null;
            }
        }

        private void TxtCover_TextChanged(object sender, EventArgs e) {
            string path = txtCover.Text.Trim();
            if (File.Exists(path)) {
                try {
                    picCoverPreview.ImageLocation = path;
                } catch {
                    picCoverPreview.Image = null;
                }
            } else {
                picCoverPreview.Image = null;
            }
        }
        
        private string ExtractInstaExe() {
            string tempDir = Path.Combine(Path.GetTempPath(), "ComComTag_Insta");
            if (!Directory.Exists(tempDir)) Directory.CreateDirectory(tempDir);
            string exePath = Path.Combine(tempDir, "download_instagram.exe");
            
            if (!File.Exists(exePath)) {
                try {
                    using (Stream resStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("download_instagram.exe")) {
                        if (resStream != null) {
                            using (FileStream fs = new FileStream(exePath, FileMode.Create)) {
                                resStream.CopyTo(fs);
                            }
                        }
                    }
                } catch { }
            }
            return exePath;
        }

        private void RunInstaDownload(string url, List<string> imageList, Action<int> setIndex, TextBox targetTextBox, Button prevBtn, Button nextBtn, Button downloadBtn) {
            string scriptPath = ExtractInstaExe();
            if (string.IsNullOrEmpty(scriptPath) || !File.Exists(scriptPath)) {
                MessageBox.Show("Could not extract internal download_instagram.exe payload.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            downloadBtn.Enabled = false;

            System.Threading.ThreadPool.QueueUserWorkItem((state) => {
                try {
                    var startInfo = new System.Diagnostics.ProcessStartInfo {
                        FileName = scriptPath,
                        Arguments = string.Format("\"{0}\"", url),
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };

                    var process = System.Diagnostics.Process.Start(startInfo);
                    string output = process.StandardOutput.ReadToEnd().Trim();
                    string error = process.StandardError.ReadToEnd().Trim();
                    process.WaitForExit();

                    if (!this.IsDisposed && this.IsHandleCreated) {
                        this.Invoke((MethodInvoker)(() => {
                            this.Cursor = Cursors.Default;
                            downloadBtn.Enabled = true;

                            var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            
                            imageList.Clear();
                            foreach (var line in lines) {
                                string l = line.Trim();
                                if (!string.IsNullOrEmpty(l) && File.Exists(l) && l.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase)) {
                                    imageList.Add(l);
                                }
                            }

                            if (process.ExitCode == 0 && imageList.Count > 0) {
                                setIndex(0);
                                targetTextBox.Text = imageList[0];
                                
                                bool multi = imageList.Count > 1;
                                prevBtn.Visible = multi;
                                nextBtn.Visible = multi;
                            } else {
                                MessageBox.Show(string.Format("Failed to download image.\nOutput: {0}\nError: {1}", output, error), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }));
                    }
                } catch (Exception ex) {
                    if (!this.IsDisposed && this.IsHandleCreated) {
                        this.Invoke((MethodInvoker)(() => {
                            this.Cursor = Cursors.Default;
                            downloadBtn.Enabled = true;
                            MessageBox.Show("Error running Python: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                }
            });
        }

        private void BtnInstaDownload_Click(object sender, EventArgs e) {
            string url = Microsoft.VisualBasic.Interaction.InputBox("Paste an Instagram Post URL or Shortcode:", "Download Cover from Instagram", "");
            if (string.IsNullOrWhiteSpace(url)) return;
            RunInstaDownload(url, _instaImages, idx => _instaImageIndex = idx, txtCover, btnInstaPrev, btnInstaNext, btnInstaDownload);
        }

        // ---------- TAGGING INSTA Logic ----------
        private void BtnTagInstaPrev_Click(object sender, EventArgs e) {
            if (_tagInstaImages.Count > 0) {
                _tagInstaImageIndex = (_tagInstaImageIndex - 1 + _tagInstaImages.Count) % _tagInstaImages.Count;
                txtTagCover.Text = _tagInstaImages[_tagInstaImageIndex];
            }
        }

        private void BtnTagInstaNext_Click(object sender, EventArgs e) {
            if (_tagInstaImages.Count > 0) {
                _tagInstaImageIndex = (_tagInstaImageIndex + 1) % _tagInstaImages.Count;
                txtTagCover.Text = _tagInstaImages[_tagInstaImageIndex];
            }
        }

        private void BtnTagInstaDownload_Click(object sender, EventArgs e) {
            string url = Microsoft.VisualBasic.Interaction.InputBox("Paste an Instagram Post URL or Shortcode:", "Download Cover from Instagram", "");
            if (string.IsNullOrWhiteSpace(url)) return;
            RunInstaDownload(url, _tagInstaImages, idx => _tagInstaImageIndex = idx, txtTagCover, btnTagInstaPrev, btnTagInstaNext, btnTagInstaDownload);
        }
        // ------------------------------------------

        private void BtnBuildM4b_Click(object sender, EventArgs e) {
            if (listChapters.Items.Count == 0) {
                MessageBox.Show("Please add MP3s to the Audiobook Chapters list.", "No files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string album = txtM4bAlbum.Text.Trim();
            string artist = txtM4bArtist.Text.Trim();
            string m4bDateStr = dtpM4bDate.Value.ToString("yyyy-MM-dd");
            string m4bLocation = cmbM4bLocation.Text.Trim();

            if (string.IsNullOrWhiteSpace(album)) {
                MessageBox.Show("Please provide an Album name for the M4B.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string defaultFileName = m4bDateStr;
            if (!string.IsNullOrWhiteSpace(artist)) defaultFileName += " - " + artist;
            defaultFileName += " - " + album;
            if (!string.IsNullOrWhiteSpace(m4bLocation)) defaultFileName += " - " + m4bLocation;
            defaultFileName += ".m4b";

            using (var dialog = new SaveFileDialog { Filter = "Audiobook (*.m4b)|*.m4b", DefaultExt = "m4b", FileName = defaultFileName, InitialDirectory = _selectedFolder }) {
                if (dialog.ShowDialog() == DialogResult.OK) {
                    
                    List<string> mp3s = new List<string>();
                    List<string> chapterTitles = new List<string>();
                    foreach (ChapterItem item in listChapters.Items) {
                        mp3s.Add(item.FilePath);
                        chapterTitles.Add(item.ChapterName);
                    }

                    string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    string fullFfmpeg = _settings.FFmpegPath;
                    if (!Path.IsPathRooted(fullFfmpeg)) {
                        fullFfmpeg = Path.Combine(exePath, fullFfmpeg);
                    }

                    if (!File.Exists(fullFfmpeg)) {
                        DialogResult dr = MessageBox.Show(
                            "FFmpeg is required to build M4B audiobooks. It was not found at:\n\n" + fullFfmpeg + 
                            "\n\nWould you like to locate ffmpeg.exe now? (If you don't have it, you can download it from ffmpeg.org)",
                            "FFmpeg Required", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (dr == DialogResult.Yes) {
                            using (var ofd = new OpenFileDialog { Filter = "Executable Files (*.exe)|*.exe", Title = "Locate ffmpeg.exe" }) {
                                if (ofd.ShowDialog() == DialogResult.OK) {
                                    _settings.FFmpegPath = ofd.FileName;
                                    _settings.Save();
                                    fullFfmpeg = _settings.FFmpegPath;
                                } else {
                                    return;
                                }
                            }
                        } else {
                            return;
                        }
                    }

                    string bitrate = cmbBitrate.SelectedItem.ToString();

                    this.Cursor = Cursors.WaitCursor;
                    btnBuildM4b.Enabled = false;
                    pbM4bProgress.Value = 0;

                    System.Threading.ThreadPool.QueueUserWorkItem((state) => {
                        string result = M4bBuilder.Build(mp3s, chapterTitles, dialog.FileName, txtCover.Text.Trim(), album, artist, bitrate, fullFfmpeg, 
                            (pct) => {
                                if (!this.IsDisposed && this.IsHandleCreated) {
                                    this.Invoke((MethodInvoker)(() => {
                                        pbM4bProgress.Value = pct;
                                    }));
                                }
                            },
                            (msg) => { Debug.WriteLine(msg); }
                        );

                        if (!this.IsDisposed && this.IsHandleCreated) {
                            this.Invoke((MethodInvoker)(() => {
                                this.Cursor = Cursors.Default;
                                btnBuildM4b.Enabled = true;
                                pbM4bProgress.Value = 0;

                                if (result == "Success") {
                                    MessageBox.Show("M4B Audiobook created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                } else {
                                    MessageBox.Show(result, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }));
                        }
                    });
                }
            }
        }
    }
}
