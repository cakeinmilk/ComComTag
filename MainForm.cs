using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ComComTag {
    public class MainForm : Form {
        private Settings _settings;
        private string _selectedFolder = "";
        private bool _isBatchUpdating = false;

        // UI Controls - Main
        private MenuStrip menuStrip;
        private Button btnBrowse;
        private Button btnRefresh;
        private Button btnOpenExplorer;
        private Label lblCurrentFolder;
        private DarkTabControl tabControl;

        // UI Controls - Tagging Tab
        private TabPage tabTagging;
        private ListBox listTagFiles;
        private Label lblHelp;
        
        // Field Inclusion Checkboxes & Labels
        private CheckBox chkIncDate;
        private Label lblDate;
        private DateTimePicker dtpDate;
        
        private CheckBox chkIncArtist;
        private Label lblArtist;
        private ComboBox cmbArtist;
        private Button btnAddArtistField;
        private List<ComboBox> _artistCombos = new List<ComboBox>();
        private List<Button> _removeArtistButtons = new List<Button>();
        private List<Label> _artistExtraLabels = new List<Label>();
        private Panel pnlTagControls;
        
        private CheckBox chkIncShow;
        private Label lblShow;
        private TextBox txtShow;
        
        private CheckBox chkIncLocation;
        private Label lblLocation;
        private ComboBox cmbLocation;
        
        // Cover Art Controls (Text field removed)
        private string _tagCoverPath = "";
        private Label lblCoverPrompt;
        private PictureBox picTagCoverPreview;
        private Button btnTagInstaDownload;
        private Button btnTagInstaPrev;
        private Button btnTagInstaNext;
        private Button btnTagBrowseCover;

        private Label lblPreviewPrompt;
        private Label lblTitlePreview;
        private CheckBox chkRenameMp3;
        private Button btnExecuteSave;
        private Button btnCopyTags;
        private Button btnPasteTags;
        private Button btnClearTags;
        private Button btnGuessFilename;

        // UI Controls - Audiobook Tab
        private TabPage tabAudiobook;
        private Label lblAvail;
        private ListBox listAvailableMp3s;
        private Button btnAddChapter;
        private Button btnRemoveChapter;
        private Label lblChaps;
        private Label lblChapsHint;
        private ListBox listChapters;
        private Button btnMoveUp;
        private Button btnMoveDown;
        private TextBox _inlineChapterEditBox;
        private int _editingChapterIndex = -1;

        private Label lblMeta;
        private Label lblM4bDateLbl;
        private DateTimePicker dtpM4bDate;
        private Label lblArt;
        private ComboBox cmbM4bArtist;
        private Label lblAlb;
        private TextBox txtM4bAlbum;
        private Label lblM4bLocLbl;
        private ComboBox cmbM4bLocation;
        private Label lblBitrate;
        private ComboBox cmbBitrate;
        
        // Audiobook Cover Controls (Text field removed)
        private string _m4bCoverPath = "";
        private Label lblM4bCover;
        private PictureBox picCoverPreview;
        private Button btnBrowseCover;
        private Button btnInstaDownload;
        private Button btnInstaPrev;
        private Button btnInstaNext;
        private Label lblM4bPreviewPrompt;
        private Label lblM4bTitlePreview;
        private Button btnBuildM4b;

        // State for Multi-Image Carousel
        private List<string> _tagInstaImages = new List<string>();
        private int _tagInstaImageIndex = 0;
        private List<string> _instaImages = new List<string>();
        private int _instaImageIndex = 0;

        // State for Tag Clipboard
        private DateTime _clipDate;
        private List<string> _clipArtists = new List<string>();
        private string _clipShow = "";
        private string _clipLocation = "";
        private bool _hasClipboardTags = false;

        private ToolTip toolTip;

        public MainForm() {
            _settings = new Settings();
            toolTip = new ToolTip();
            
            // Set Application Icon if embedded
            try {
                var assembly = Assembly.GetExecutingAssembly();
                using (var iconStream = assembly.GetManifestResourceStream("ComComTag.icon.ico")) {
                    if (iconStream != null) {
                        this.Icon = new Icon(iconStream);
                    }
                }
            } catch { }

            InitializeComponent();
            ThemeHelper.ApplyTheme(this, ThemeHelper.ShouldUseDarkMode(_settings.Theme));
            CheckFFmpegOnStartup();
        }

        private void CheckFFmpegOnStartup() {
            if (!_settings.IsFFmpegAvailable()) {
                MessageBox.Show(
                    "ffmpeg.exe was not found. The 'Build Audiobook' feature will require FFmpeg.\nYou can configure its path anytime in Edit -> Settings.",
                    "FFmpeg Missing", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void InitializeComponent() {
            string appVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString(3);
            this.Text = "ComComTag v" + appVersion;
            this.Size = new Size(900, 700);
            this.MinimumSize = new Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;

            // --- Menu Bar ---
            menuStrip = new MenuStrip();
            ToolStripMenuItem menuFile = new ToolStripMenuItem("File");
            ToolStripMenuItem menuBrowse = new ToolStripMenuItem("Open Folder...");
            menuBrowse.Click += BtnBrowse_Click;
            ToolStripMenuItem menuRefresh = new ToolStripMenuItem("Refresh Folder");
            menuRefresh.ShortcutKeys = Keys.F5;
            menuRefresh.Click += BtnRefresh_Click;
            ToolStripMenuItem menuExplorer = new ToolStripMenuItem("Open in Explorer");
            menuExplorer.Click += BtnOpenExplorer_Click;
            ToolStripMenuItem menuExit = new ToolStripMenuItem("Exit");
            menuExit.Click += (s, e) => this.Close();
            menuFile.DropDownItems.Add(menuBrowse);
            menuFile.DropDownItems.Add(menuRefresh);
            menuFile.DropDownItems.Add(menuExplorer);
            menuFile.DropDownItems.Add(new ToolStripSeparator());
            menuFile.DropDownItems.Add(menuExit);

            ToolStripMenuItem menuEdit = new ToolStripMenuItem("Edit");
            ToolStripMenuItem menuSettings = new ToolStripMenuItem("Settings");
            menuSettings.Click += MenuSettings_Click;
            menuEdit.DropDownItems.Add(menuSettings);

            ToolStripMenuItem menuHelp = new ToolStripMenuItem("Help");
            ToolStripMenuItem menuGuide = new ToolStripMenuItem("User Guide");
            menuGuide.ShortcutKeys = Keys.F1;
            menuGuide.Click += (s, e) => new HelpDialog(0, _settings.Theme).ShowDialog(this);
            ToolStripMenuItem menuPatterns = new ToolStripMenuItem("Naming Patterns");
            menuPatterns.Click += (s, e) => new HelpDialog(1, _settings.Theme).ShowDialog(this);
            ToolStripMenuItem menuInstaHelp = new ToolStripMenuItem("Instagram Art Downloader");
            menuInstaHelp.Click += (s, e) => new HelpDialog(2, _settings.Theme).ShowDialog(this);
            ToolStripMenuItem menuShortcuts = new ToolStripMenuItem("Keyboard Shortcuts");
            menuShortcuts.Click += (s, e) => new HelpDialog(3, _settings.Theme).ShowDialog(this);
            ToolStripMenuItem menuAbout = new ToolStripMenuItem("About ComComTag");
            menuAbout.Click += (s, e) => new HelpDialog(4, _settings.Theme).ShowDialog(this);

            menuHelp.DropDownItems.Add(menuGuide);
            menuHelp.DropDownItems.Add(menuPatterns);
            menuHelp.DropDownItems.Add(menuInstaHelp);
            menuHelp.DropDownItems.Add(menuShortcuts);
            menuHelp.DropDownItems.Add(new ToolStripSeparator());
            menuHelp.DropDownItems.Add(menuAbout);

            menuStrip.Items.Add(menuFile);
            menuStrip.Items.Add(menuEdit);
            menuStrip.Items.Add(menuHelp);
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);

            // --- Top Folder Selection & Refresh (Icon Buttons) ---
            btnBrowse = CreateIconButton(new Point(15, 33), new Size(34, 28), DrawFolderIcon);
            toolTip.SetToolTip(btnBrowse, "Open Folder...");
            btnBrowse.Click += BtnBrowse_Click;

            btnRefresh = CreateIconButton(new Point(53, 33), new Size(34, 28), DrawSyncIcon);
            toolTip.SetToolTip(btnRefresh, "Refresh Folder (F5)");
            btnRefresh.Click += BtnRefresh_Click;

            btnOpenExplorer = CreateIconButton(new Point(91, 33), new Size(34, 28), DrawExplorerIcon);
            toolTip.SetToolTip(btnOpenExplorer, "Open Folder in Windows Explorer");
            btnOpenExplorer.Click += BtnOpenExplorer_Click;

            lblCurrentFolder = new Label { 
                Location = new Point(135, 38), 
                Size = new Size(735, 20), 
                Text = "No folder selected.",
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right 
            };
            this.Controls.Add(btnBrowse);
            this.Controls.Add(btnRefresh);
            this.Controls.Add(btnOpenExplorer);
            this.Controls.Add(lblCurrentFolder);

            // --- Tabs ---
            tabControl = new DarkTabControl { 
                Location = new Point(15, 68), 
                Size = new Size(855, 575),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
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

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) {
            if (keyData == (Keys.Control | Keys.S)) {
                if (btnExecuteSave != null && btnExecuteSave.Enabled) {
                    btnExecuteSave.PerformClick();
                    return true;
                }
            } else if (keyData == (Keys.Control | Keys.B)) {
                if (btnBuildM4b != null && btnBuildM4b.Enabled && btnBuildM4b.Visible) {
                    btnBuildM4b.PerformClick();
                    return true;
                }
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void MenuSettings_Click(object sender, EventArgs e) {
            using (var sf = new SettingsForm(_settings)) {
                if (sf.ShowDialog(this) == DialogResult.OK) {
                    ThemeHelper.ApplyTheme(this, ThemeHelper.ShouldUseDarkMode(_settings.Theme));
                    RefreshDropdownLists();
                    UpdateTitlePreview();
                    UpdateM4bPreview();
                }
            }
        }

        private void RefreshDropdownLists() {
            // Refresh Locations
            cmbLocation.Items.Clear();
            cmbLocation.Items.AddRange(_settings.Locations.ToArray());
            cmbM4bLocation.Items.Clear();
            cmbM4bLocation.Items.AddRange(_settings.Locations.ToArray());

            // Refresh Artists across all active artist combo boxes
            foreach (var cmb in _artistCombos) {
                string cur = cmb.Text;
                cmb.Items.Clear();
                cmb.Items.AddRange(_settings.Artists.ToArray());
                cmb.Text = cur;
            }
            cmbM4bArtist.Items.Clear();
            cmbM4bArtist.Items.AddRange(_settings.Artists.ToArray());
        }

        private void InitializeTaggingTab() {
            // Right panel narrowed to 310px to expand the MP3 file list on the left by ~180px
            pnlTagControls = new Panel {
                Dock = DockStyle.Right,
                Width = 310,
                AutoScroll = true
            };

            listTagFiles = new ListBox {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.MultiExtended,
                IntegralHeight = false
            };
            listTagFiles.SelectedIndexChanged += ListTagFiles_SelectedIndexChanged;
            SetupListDragAndDrop(listTagFiles);

            // --- Top Action Icon Buttons ---
            btnClearTags = CreateIconButton(new Point(15, 8), new Size(32, 28), DrawClearIcon);
            toolTip.SetToolTip(btnClearTags, "Clear all ID3 tags from selected files and reset fields");
            btnClearTags.Click += BtnClearTags_Click;

            btnGuessFilename = CreateIconButton(new Point(52, 8), new Size(32, 28), DrawWandIcon);
            toolTip.SetToolTip(btnGuessFilename, "Attempts to extract ID3 tags from filename given current Naming Conventions in Settings");
            btnGuessFilename.Click += BtnGuessFilename_Click;

            btnCopyTags = CreateIconButton(new Point(89, 8), new Size(32, 28), DrawCopyIcon);
            toolTip.SetToolTip(btnCopyTags, "Copy tags to clipboard");
            btnCopyTags.Click += BtnCopyTags_Click;
            btnCopyTags.Enabled = false;

            btnPasteTags = CreateIconButton(new Point(126, 8), new Size(32, 28), DrawPasteIcon);
            toolTip.SetToolTip(btnPasteTags, "Paste tags from clipboard");
            btnPasteTags.Click += BtnPasteTags_Click;
            btnPasteTags.Enabled = false;

            lblHelp = new Label { Location = new Point(15, 42), Size = new Size(280, 28), Text = "Select MP3s to tag and rename.", Font = new Font(this.Font, FontStyle.Italic) };

            // --- Date Row ---
            chkIncDate = new CheckBox { Location = new Point(15, 74), Size = new Size(18, 20), Checked = true };
            chkIncDate.CheckedChanged += (s, e) => UpdateTitlePreview();
            toolTip.SetToolTip(chkIncDate, "Include Date in generated filename");

            lblDate = new Label { Location = new Point(35, 76), Size = new Size(45, 20), Text = "Date:" };
            dtpDate = new DateTimePicker { Location = new Point(85, 74), Size = new Size(115, 22), Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtpDate.ValueChanged += (s, e) => UpdateTitlePreview();

            // --- Primary Artist Row ---
            chkIncArtist = new CheckBox { Location = new Point(15, 104), Size = new Size(18, 20), Checked = true };
            chkIncArtist.CheckedChanged += (s, e) => UpdateTitlePreview();
            toolTip.SetToolTip(chkIncArtist, "Include Artist in generated filename");

            lblArtist = new Label { Location = new Point(35, 106), Size = new Size(45, 20), Text = "Artist:" };
            cmbArtist = new ComboBox { 
                Location = new Point(85, 104), 
                Size = new Size(175, 22), 
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            cmbArtist.Items.AddRange(_settings.Artists.ToArray());
            cmbArtist.TextChanged += (s, e) => {
                UpdateAddArtistButtonState();
                UpdateTitlePreview();
            };

            btnAddArtistField = new Button {
                Location = new Point(265, 104),
                Size = new Size(25, 22),
                Text = "+",
                Font = new Font(this.Font, FontStyle.Bold),
                Enabled = false
            };
            toolTip.SetToolTip(btnAddArtistField, "Add another Artist field (for shared shows)");
            btnAddArtistField.Click += (s, e) => {
                AddExtraArtistField("");
                RelayoutTagControls();
                UpdateTitlePreview();
            };

            _artistCombos.Clear();
            _artistCombos.Add(cmbArtist);
            UpdateAddArtistButtonState();

            // --- Show Row ---
            chkIncShow = new CheckBox { Location = new Point(15, 134), Size = new Size(18, 20), Checked = true };
            chkIncShow.CheckedChanged += (s, e) => UpdateTitlePreview();
            toolTip.SetToolTip(chkIncShow, "Include Show in generated filename");

            lblShow = new Label { Location = new Point(35, 136), Size = new Size(45, 20), Text = "Show:" };
            txtShow = new TextBox { Location = new Point(85, 134), Size = new Size(205, 22) };
            txtShow.TextChanged += (s, e) => UpdateTitlePreview();

            // --- Venue Row ---
            chkIncLocation = new CheckBox { Location = new Point(15, 164), Size = new Size(18, 20), Checked = true };
            chkIncLocation.CheckedChanged += (s, e) => UpdateTitlePreview();
            toolTip.SetToolTip(chkIncLocation, "Include Venue/Location in generated filename");

            lblLocation = new Label { Location = new Point(35, 166), Size = new Size(50, 20), Text = "Venue:" };
            cmbLocation = new ComboBox { 
                Location = new Point(85, 164), 
                Size = new Size(205, 22), 
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            cmbLocation.Items.AddRange(_settings.Locations.ToArray());
            cmbLocation.TextChanged += (s, e) => UpdateTitlePreview();

            // --- Cover Art Preview ---
            lblCoverPrompt = new Label { Location = new Point(15, 196), Size = new Size(50, 20), Text = "Cover:" };
            
            picTagCoverPreview = new PictureBox { 
                Location = new Point(85, 196), 
                Size = new Size(90, 90), 
                SizeMode = PictureBoxSizeMode.Zoom, 
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand
            };
            picTagCoverPreview.Click += (s, e) => PromptTagCoverFile();
            picTagCoverPreview.DoubleClick += (s, e) => {
                if (_tagInstaImages != null && _tagInstaImages.Count > 1) {
                    OpenCarouselModal(_tagInstaImages, _tagInstaImageIndex, p => SetTagCover(p), idx => _tagInstaImageIndex = idx);
                } else {
                    PromptTagCoverFile();
                }
            };
            SetupCoverDragAndDrop(picTagCoverPreview, p => SetTagCover(p));
            toolTip.SetToolTip(picTagCoverPreview, "Click to choose cover image from computer, or drag-and-drop an image here");

            btnTagBrowseCover = new Button { Location = new Point(185, 196), Size = new Size(65, 24), Text = "Browse..." };
            btnTagBrowseCover.Click += (s, e) => PromptTagCoverFile();

            btnTagInstaDownload = CreateInstagramButton(new Point(185, 224), new Size(30, 28));
            btnTagInstaDownload.Click += (s, e) => HandleInstaDownload(_tagInstaImages, idx => _tagInstaImageIndex = idx, p => SetTagCover(p), btnTagInstaPrev, btnTagInstaNext, btnTagInstaDownload);
            toolTip.SetToolTip(btnTagInstaDownload, "Download Cover from Instagram");

            btnTagInstaPrev = new Button { Location = new Point(185, 258), Size = new Size(25, 24), Text = "<", Visible = false };
            btnTagInstaPrev.Click += (s, e) => CycleCarousel(_tagInstaImages, ref _tagInstaImageIndex, -1, p => SetTagCover(p));
            
            btnTagInstaNext = new Button { Location = new Point(215, 258), Size = new Size(25, 24), Text = ">", Visible = false };
            btnTagInstaNext.Click += (s, e) => CycleCarousel(_tagInstaImages, ref _tagInstaImageIndex, 1, p => SetTagCover(p));

            // --- Filename Preview ---
            lblPreviewPrompt = new Label { Location = new Point(15, 296), Size = new Size(115, 18), Text = "Filename Preview:", Font = new Font(this.Font, FontStyle.Bold) };
            lblTitlePreview = new Label { Location = new Point(15, 316), Size = new Size(275, 45), ForeColor = SystemColors.GrayText, Text = "" };

            // --- Save & Rename Controls ---
            chkRenameMp3 = new CheckBox { 
                Location = new Point(15, 366), 
                Size = new Size(120, 20), 
                Text = "Rename MP3", 
                Checked = true, 
                Font = new Font(this.Font, FontStyle.Regular) 
            };
            chkRenameMp3.CheckedChanged += (s, e) => {
                btnExecuteSave.Text = chkRenameMp3.Checked ? "Save Tags & Rename" : "Save ID3 tags";
                UpdateTitlePreview();
            };

            btnExecuteSave = new Button { 
                Location = new Point(15, 390), 
                Size = new Size(275, 36), 
                Text = "Save Tags & Rename", 
                UseMnemonic = false,
                Enabled = false, 
                Font = new Font(this.Font, FontStyle.Bold) 
            };
            toolTip.SetToolTip(btnExecuteSave, "Save tags and rename on disk (Ctrl+S)");
            btnExecuteSave.Click += BtnExecuteSave_Click;

            pnlTagControls.Controls.Add(btnClearTags);
            pnlTagControls.Controls.Add(btnGuessFilename);
            pnlTagControls.Controls.Add(btnCopyTags);
            pnlTagControls.Controls.Add(btnPasteTags);
            pnlTagControls.Controls.Add(lblHelp);
            
            pnlTagControls.Controls.Add(chkIncDate);
            pnlTagControls.Controls.Add(lblDate);
            pnlTagControls.Controls.Add(dtpDate);
            
            pnlTagControls.Controls.Add(chkIncArtist);
            pnlTagControls.Controls.Add(lblArtist);
            pnlTagControls.Controls.Add(cmbArtist);
            pnlTagControls.Controls.Add(btnAddArtistField);
            
            pnlTagControls.Controls.Add(chkIncShow);
            pnlTagControls.Controls.Add(lblShow);
            pnlTagControls.Controls.Add(txtShow);
            
            pnlTagControls.Controls.Add(chkIncLocation);
            pnlTagControls.Controls.Add(lblLocation);
            pnlTagControls.Controls.Add(cmbLocation);
            
            pnlTagControls.Controls.Add(lblCoverPrompt);
            pnlTagControls.Controls.Add(picTagCoverPreview);
            pnlTagControls.Controls.Add(btnTagBrowseCover);
            pnlTagControls.Controls.Add(btnTagInstaDownload);
            pnlTagControls.Controls.Add(btnTagInstaPrev);
            pnlTagControls.Controls.Add(btnTagInstaNext);
            
            pnlTagControls.Controls.Add(lblPreviewPrompt);
            pnlTagControls.Controls.Add(lblTitlePreview);
            pnlTagControls.Controls.Add(chkRenameMp3);
            pnlTagControls.Controls.Add(btnExecuteSave);

            tabTagging.Controls.Add(listTagFiles);
            tabTagging.Controls.Add(pnlTagControls);

            RelayoutTagControls();
        }

        private void AddExtraArtistField(string initialText) {
            var cmbExtra = new ComboBox {
                Size = new Size(175, 22),
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                Text = initialText ?? ""
            };
            cmbExtra.Items.AddRange(_settings.Artists.ToArray());
            cmbExtra.TextChanged += (s, e) => {
                UpdateAddArtistButtonState();
                UpdateTitlePreview();
            };

            var btnRemove = new Button {
                Size = new Size(25, 22),
                Text = "-",
                Font = new Font(this.Font, FontStyle.Bold)
            };
            toolTip.SetToolTip(btnRemove, "Remove this artist field");

            var lblExtra = new Label {
                Size = new Size(65, 20),
                Text = string.Format("Artist {0}:", _artistCombos.Count + 1),
                TextAlign = ContentAlignment.MiddleRight
            };

            btnRemove.Click += (s, e) => {
                RemoveExtraArtistField(cmbExtra, btnRemove, lblExtra);
            };

            _artistCombos.Add(cmbExtra);
            _removeArtistButtons.Add(btnRemove);
            _artistExtraLabels.Add(lblExtra);

            pnlTagControls.Controls.Add(lblExtra);
            pnlTagControls.Controls.Add(cmbExtra);
            pnlTagControls.Controls.Add(btnRemove);

            UpdateAddArtistButtonState();
        }

        private void RemoveExtraArtistField(ComboBox cmb, Button btn, Label lbl) {
            pnlTagControls.Controls.Remove(cmb);
            pnlTagControls.Controls.Remove(btn);
            pnlTagControls.Controls.Remove(lbl);

            _artistCombos.Remove(cmb);
            _removeArtistButtons.Remove(btn);
            _artistExtraLabels.Remove(lbl);

            cmb.Dispose();
            btn.Dispose();
            lbl.Dispose();

            for (int i = 0; i < _artistExtraLabels.Count; i++) {
                _artistExtraLabels[i].Text = string.Format("Artist {0}:", i + 2);
            }

            RelayoutTagControls();
            UpdateAddArtistButtonState();
            UpdateTitlePreview();
        }

        private void SetArtistFields(IList<string> artists) {
            while (_artistCombos.Count > 1) {
                int lastIdx = _artistCombos.Count - 1;
                var cmb = _artistCombos[lastIdx];
                var btn = _removeArtistButtons[lastIdx - 1];
                var lbl = _artistExtraLabels[lastIdx - 1];

                pnlTagControls.Controls.Remove(cmb);
                pnlTagControls.Controls.Remove(btn);
                pnlTagControls.Controls.Remove(lbl);

                _artistCombos.RemoveAt(lastIdx);
                _removeArtistButtons.RemoveAt(lastIdx - 1);
                _artistExtraLabels.RemoveAt(lastIdx - 1);

                cmb.Dispose();
                btn.Dispose();
                lbl.Dispose();
            }

            if (artists != null && artists.Count > 0) {
                _artistCombos[0].Text = artists[0] ?? "";
                for (int i = 1; i < artists.Count; i++) {
                    if (!string.IsNullOrWhiteSpace(artists[i])) {
                        AddExtraArtistField(artists[i]);
                    }
                }
            } else {
                _artistCombos[0].Text = "";
            }

            RelayoutTagControls();
            UpdateAddArtistButtonState();
        }

        private void UpdateAddArtistButtonState() {
            if (btnAddArtistField == null) return;
            bool canAdd = true;
            if (_artistCombos == null || _artistCombos.Count == 0) {
                canAdd = false;
            } else {
                foreach (var cmb in _artistCombos) {
                    if (string.IsNullOrWhiteSpace(cmb.Text)) {
                        canAdd = false;
                        break;
                    }
                }
            }
            btnAddArtistField.Enabled = canAdd;
        }

        public static string FormatArtistsForFilename(IList<string> artists) {
            return FormatArtistsForFilename(artists, null);
        }

        public static string FormatArtistsForFilename(IList<string> artists, Settings settings) {
            if (artists == null || artists.Count == 0) return "";
            if (artists.Count == 1) return artists[0];

            int threshold = (settings != null && settings.GroupAdditionalArtistsThreshold >= 2) 
                ? settings.GroupAdditionalArtistsThreshold 
                : 3;

            if (settings != null && settings.GroupAdditionalArtists && artists.Count >= threshold) {
                string groupText = string.IsNullOrWhiteSpace(settings.GroupAdditionalArtistsText) ? "and friends" : settings.GroupAdditionalArtistsText.Trim();
                return string.Format("{0} {1}", artists[0], groupText);
            }

            if (artists.Count == 2) return string.Format("{0} and {1}", artists[0], artists[1]);

            var sb = new StringBuilder();
            for (int i = 0; i < artists.Count; i++) {
                if (i > 0) {
                    if (i == artists.Count - 1) {
                        sb.Append(", and ");
                    } else {
                        sb.Append(", ");
                    }
                }
                sb.Append(artists[i]);
            }
            return sb.ToString();
        }

        private List<string> GetAllEnteredArtists() {
            var list = new List<string>();
            foreach (var cmb in _artistCombos) {
                string t = cmb.Text.Trim();
                if (!string.IsNullOrEmpty(t) && !list.Contains(t)) {
                    list.Add(t);
                }
            }
            return list;
        }

        private void RelayoutTagControls() {
            if (pnlTagControls == null || cmbArtist == null) return;

            int yOffset = 8;
            
            btnClearTags.Location = new Point(15, yOffset);
            btnGuessFilename.Location = new Point(52, yOffset);
            btnCopyTags.Location = new Point(89, yOffset);
            btnPasteTags.Location = new Point(126, yOffset);
            
            yOffset += 34;
            lblHelp.Location = new Point(15, yOffset);

            // Date
            yOffset += 32;
            chkIncDate.Location = new Point(15, yOffset);
            lblDate.Location = new Point(35, yOffset + 2);
            dtpDate.Location = new Point(85, yOffset);

            // Artist 1
            yOffset += 30;
            chkIncArtist.Location = new Point(15, yOffset);
            lblArtist.Location = new Point(35, yOffset + 2);
            cmbArtist.Location = new Point(85, yOffset);
            btnAddArtistField.Location = new Point(265, yOffset);

            // Extra Artists
            for (int i = 1; i < _artistCombos.Count; i++) {
                yOffset += 28;
                var lbl = _artistExtraLabels[i - 1];
                var cmb = _artistCombos[i];
                var btn = _removeArtistButtons[i - 1];

                lbl.Location = new Point(15, yOffset + 2);
                lbl.Size = new Size(65, 20);
                cmb.Location = new Point(85, yOffset);
                cmb.Size = new Size(175, 22);
                btn.Location = new Point(265, yOffset);
                btn.Size = new Size(25, 22);
            }

            // Show
            yOffset += 30;
            chkIncShow.Location = new Point(15, yOffset);
            lblShow.Location = new Point(35, yOffset + 2);
            txtShow.Location = new Point(85, yOffset);

            // Location / Venue
            yOffset += 30;
            chkIncLocation.Location = new Point(15, yOffset);
            lblLocation.Location = new Point(35, yOffset + 2);
            cmbLocation.Location = new Point(85, yOffset);

            // Cover
            yOffset += 32;
            lblCoverPrompt.Location = new Point(15, yOffset);
            picTagCoverPreview.Location = new Point(85, yOffset);
            btnTagBrowseCover.Location = new Point(185, yOffset);
            btnTagInstaDownload.Location = new Point(185, yOffset + 28);
            btnTagInstaPrev.Location = new Point(185, yOffset + 62);
            btnTagInstaNext.Location = new Point(215, yOffset + 62);

            // Filename Preview
            yOffset += 100;
            lblPreviewPrompt.Location = new Point(15, yOffset);
            lblTitlePreview.Location = new Point(15, yOffset + 20);

            // Save Controls
            yOffset += 70;
            chkRenameMp3.Location = new Point(15, yOffset);
            btnExecuteSave.Location = new Point(15, yOffset + 24);
        }

        private void InitializeAudiobookTab() {
            tabAudiobook.Controls.Clear();

            // Top Area: Available MP3s and Chapters
            lblAvail = new Label { Location = new Point(15, 10), Size = new Size(200, 18), Text = "1. Available MP3s:" };
            listAvailableMp3s = new ListBox { 
                Location = new Point(15, 30), 
                Size = new Size(400, 326), 
                SelectionMode = SelectionMode.MultiExtended
            };
            listAvailableMp3s.DoubleClick += (s, e) => {
                if (listAvailableMp3s.SelectedItems.Count > 0) BtnAddChapter_Click(s, e);
            };
            listAvailableMp3s.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter && listAvailableMp3s.SelectedItems.Count > 0) {
                    BtnAddChapter_Click(s, e);
                    e.Handled = true;
                }
            };
            SetupListDragAndDrop(listAvailableMp3s);
            
            btnAddChapter = CreateIconButton(new Point(425, 145), new Size(36, 30), DrawForwardIcon);
            toolTip.SetToolTip(btnAddChapter, "Add selected MP3s as chapters (Enter / Double-Click)");
            btnAddChapter.Click += BtnAddChapter_Click;

            btnRemoveChapter = CreateIconButton(new Point(425, 185), new Size(36, 30), DrawBackIcon);
            toolTip.SetToolTip(btnRemoveChapter, "Remove selected chapters (Del / Double-Click)");
            btnRemoveChapter.Click += BtnRemoveChapter_Click;

            lblChaps = new Label { Location = new Point(475, 10), AutoSize = true, Text = "2. Chapters:" };
            lblChapsHint = new Label { 
                Location = new Point(565, 10), 
                AutoSize = true, 
                Font = new Font(this.Font, FontStyle.Italic), 
                ForeColor = SystemColors.GrayText, 
                Text = "(Double-click to edit name)" 
            };
            listChapters = new ListBox { 
                Location = new Point(475, 30), 
                Size = new Size(300, 326),
                SelectionMode = SelectionMode.MultiExtended
            };
            SetupChapterInlineEditing();
            
            // Move Up and Down Icon Buttons
            btnMoveUp = CreateIconButton(new Point(785, 145), new Size(34, 30), DrawUpIcon);
            toolTip.SetToolTip(btnMoveUp, "Move chapter up (Ctrl+Up)");
            btnMoveUp.Click += BtnMoveUp_Click;

            btnMoveDown = CreateIconButton(new Point(785, 185), new Size(34, 30), DrawDownIcon);
            toolTip.SetToolTip(btnMoveDown, "Move chapter down (Ctrl+Down)");
            btnMoveDown.Click += BtnMoveDown_Click;

            // Bottom Area: Metadata & Build
            int yOffset = 365;
            lblMeta = new Label { Location = new Point(15, yOffset), Size = new Size(200, 20), Text = "3. Audiobook Metadata:", Font = new Font(this.Font, FontStyle.Bold) };
            
            // Row 1: Date & Bitrate
            yOffset += 24;
            lblM4bDateLbl = new Label { Location = new Point(15, yOffset + 2), Size = new Size(45, 20), Text = "Date:" };
            dtpM4bDate = new DateTimePicker { Location = new Point(60, yOffset), Size = new Size(120, 22), Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtpM4bDate.ValueChanged += (s, e) => UpdateM4bPreview();

            lblBitrate = new Label { Location = new Point(190, yOffset + 2), Size = new Size(45, 20), Text = "Bitrate:" };
            cmbBitrate = new ComboBox { Location = new Point(240, yOffset), Size = new Size(65, 22), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbBitrate.Items.AddRange(new object[] { "64k", "96k", "128k", "192k", "256k", "320k" });
            if (!string.IsNullOrEmpty(_settings.DefaultBitrate) && cmbBitrate.Items.Contains(_settings.DefaultBitrate)) {
                cmbBitrate.SelectedIndex = cmbBitrate.Items.IndexOf(_settings.DefaultBitrate);
            } else {
                cmbBitrate.SelectedIndex = 2; // Default 128k
            }

            // Row 2: Artist
            yOffset += 28;
            lblArt = new Label { Location = new Point(15, yOffset + 2), Size = new Size(45, 20), Text = "Artist:" };
            cmbM4bArtist = new ComboBox { 
                Location = new Point(60, yOffset), 
                Size = new Size(245, 22), 
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                Text = _settings.DefaultM4bArtist ?? "Various Artists"
            };
            cmbM4bArtist.Items.AddRange(_settings.Artists.ToArray());
            cmbM4bArtist.TextChanged += (s, e) => UpdateM4bPreview();

            // Row 3: Album
            yOffset += 28;
            lblAlb = new Label { Location = new Point(15, yOffset + 2), Size = new Size(45, 20), Text = "Album:" };
            txtM4bAlbum = new TextBox { Location = new Point(60, yOffset), Size = new Size(245, 22) };
            txtM4bAlbum.TextChanged += (s, e) => UpdateM4bPreview();

            // Row 4: Venue
            yOffset += 28;
            lblM4bLocLbl = new Label { Location = new Point(15, yOffset + 2), Size = new Size(45, 20), Text = "Venue:" };
            cmbM4bLocation = new ComboBox { 
                Location = new Point(60, yOffset), 
                Size = new Size(245, 22), 
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            cmbM4bLocation.Items.AddRange(_settings.Locations.ToArray());
            cmbM4bLocation.TextChanged += (s, e) => UpdateM4bPreview();

            // Middle: Cover Art Preview Box (Size 90x90) and controls alongside fields
            lblM4bCover = new Label { Location = new Point(325, 391), Size = new Size(45, 20), Text = "Cover:" };
            picCoverPreview = new PictureBox { 
                Location = new Point(375, 389), 
                Size = new Size(90, 90), 
                SizeMode = PictureBoxSizeMode.Zoom, 
                BorderStyle = BorderStyle.FixedSingle,
                Cursor = Cursors.Hand
            };
            picCoverPreview.Click += (s, e) => PromptM4bCoverFile();
            picCoverPreview.DoubleClick += (s, e) => {
                if (_instaImages != null && _instaImages.Count > 1) {
                    OpenCarouselModal(_instaImages, _instaImageIndex, p => SetM4bCover(p), idx => _instaImageIndex = idx);
                } else {
                    PromptM4bCoverFile();
                }
            };
            SetupCoverDragAndDrop(picCoverPreview, p => SetM4bCover(p));
            toolTip.SetToolTip(picCoverPreview, "Click to choose cover image from computer, or drag-and-drop an image here");

            btnBrowseCover = new Button { Location = new Point(475, 389), Size = new Size(68, 25), Text = "Browse..." };
            btnBrowseCover.Click += (s, e) => PromptM4bCoverFile();

            btnInstaDownload = CreateInstagramButton(new Point(475, 419), new Size(30, 26));
            btnInstaDownload.Click += (s, e) => HandleInstaDownload(_instaImages, idx => _instaImageIndex = idx, p => SetM4bCover(p), btnInstaPrev, btnInstaNext, btnInstaDownload);
            toolTip.SetToolTip(btnInstaDownload, "Download Cover from Instagram");

            btnInstaPrev = new Button { Location = new Point(475, 453), Size = new Size(25, 24), Text = "<", Visible = false };
            btnInstaPrev.Click += (s, e) => CycleCarousel(_instaImages, ref _instaImageIndex, -1, p => SetM4bCover(p));
            btnInstaNext = new Button { Location = new Point(505, 453), Size = new Size(25, 24), Text = ">", Visible = false };
            btnInstaNext.Click += (s, e) => CycleCarousel(_instaImages, ref _instaImageIndex, 1, p => SetM4bCover(p));

            // Right: Build Button
            btnBuildM4b = new Button { 
                Location = new Point(650, 415), 
                Size = new Size(180, 44), 
                Text = "Build Audiobook", 
                Font = new Font(this.Font, FontStyle.Bold) 
            };
            toolTip.SetToolTip(btnBuildM4b, "Compile M4B Audiobook (Ctrl+B)");
            btnBuildM4b.Click += BtnBuildM4b_Click;

            // Filename Preview near bottom
            lblM4bPreviewPrompt = new Label { Location = new Point(15, 515), Size = new Size(115, 20), Text = "Filename Preview:", Font = new Font(this.Font, FontStyle.Bold) };
            lblM4bTitlePreview = new Label { Location = new Point(135, 515), Size = new Size(450, 20), ForeColor = SystemColors.GrayText, Text = "" };

            tabAudiobook.Controls.Add(lblAvail);
            tabAudiobook.Controls.Add(listAvailableMp3s);
            tabAudiobook.Controls.Add(btnAddChapter);
            tabAudiobook.Controls.Add(btnRemoveChapter);
            tabAudiobook.Controls.Add(lblChaps);
            tabAudiobook.Controls.Add(lblChapsHint);
            tabAudiobook.Controls.Add(listChapters);
            tabAudiobook.Controls.Add(btnMoveUp);
            tabAudiobook.Controls.Add(btnMoveDown);

            tabAudiobook.Controls.Add(lblMeta);
            tabAudiobook.Controls.Add(lblM4bDateLbl);
            tabAudiobook.Controls.Add(dtpM4bDate);
            tabAudiobook.Controls.Add(lblArt);
            tabAudiobook.Controls.Add(cmbM4bArtist);
            tabAudiobook.Controls.Add(lblAlb);
            tabAudiobook.Controls.Add(txtM4bAlbum);
            tabAudiobook.Controls.Add(lblM4bLocLbl);
            tabAudiobook.Controls.Add(cmbM4bLocation);
            tabAudiobook.Controls.Add(lblBitrate);
            tabAudiobook.Controls.Add(cmbBitrate);
            tabAudiobook.Controls.Add(lblM4bCover);
            tabAudiobook.Controls.Add(btnBrowseCover);
            tabAudiobook.Controls.Add(btnInstaDownload);
            tabAudiobook.Controls.Add(picCoverPreview);
            tabAudiobook.Controls.Add(btnInstaPrev);
            tabAudiobook.Controls.Add(btnInstaNext);
            tabAudiobook.Controls.Add(btnBuildM4b);
            tabAudiobook.Controls.Add(lblM4bPreviewPrompt);
            tabAudiobook.Controls.Add(lblM4bTitlePreview);

            // Responsive layout adjustments on resize
            tabAudiobook.Resize += (s, e) => LayoutAudiobookTab();

            LayoutAudiobookTab();
            UpdateM4bPreview();
        }

        private void LayoutAudiobookTab() {
            if (tabAudiobook == null || listAvailableMp3s == null || listChapters == null) return;
            
            int totalWidth = tabAudiobook.ClientSize.Width;
            if (totalWidth < 300) return;

            // Calculate 60% for Section 1 and 40% for Section 2
            int listGap = 10;
            int btnWidth = 36;
            int reservedWidth = 15 + btnWidth + listGap * 2 + btnWidth + listGap + 15; // Space for margins and button columns
            int availableForLists = Math.Max(200, totalWidth - reservedWidth);

            int w1 = (int)(availableForLists * 0.60);
            int w2 = availableForLists - w1;

            // Section 1 (Available MP3s)
            lblAvail.Left = 15;
            listAvailableMp3s.Left = 15;
            listAvailableMp3s.Width = w1;

            // Transfer Buttons
            int btnTransferX = 15 + w1 + listGap;
            btnAddChapter.Left = btnTransferX;
            btnRemoveChapter.Left = btnTransferX;

            // Section 2 (Audiobook Chapters)
            int chapsX = btnTransferX + btnWidth + listGap;
            lblChaps.Left = chapsX;
            if (lblChapsHint != null) lblChapsHint.Left = chapsX + lblChaps.PreferredWidth + 6;
            listChapters.Left = chapsX;
            listChapters.Width = w2;

            // Up / Down Buttons
            int btnUpDownX = chapsX + w2 + listGap;
            btnMoveUp.Left = btnUpDownX;
            btnMoveDown.Left = btnUpDownX;

            // Section 3 (Metadata & Build):
            int rightEdge = totalWidth - 15;
            
            // Build Button on Right
            btnBuildM4b.Left = rightEdge - 180;
            btnBuildM4b.Width = 180;

            // Cover Preview & Buttons in middle alongside fields
            lblM4bCover.Left = 325;
            picCoverPreview.Left = 375;
            btnBrowseCover.Left = 475;
            btnInstaDownload.Left = 475;
            btnInstaPrev.Left = 475;
            btnInstaNext.Left = 505;

            // Fields in left column matching Tag tab length (~245px)
            lblM4bDateLbl.Left = 15;
            dtpM4bDate.Left = 60;
            lblBitrate.Left = 190;
            cmbBitrate.Left = 240;

            lblArt.Left = 15;
            cmbM4bArtist.Left = 60;
            cmbM4bArtist.Width = 245;

            lblAlb.Left = 15;
            txtM4bAlbum.Left = 60;
            txtM4bAlbum.Width = 245;

            lblM4bLocLbl.Left = 15;
            cmbM4bLocation.Left = 60;
            cmbM4bLocation.Width = 245;

            // Bottom Filename Preview
            lblM4bPreviewPrompt.Left = 15;
            lblM4bTitlePreview.Left = 135;
            lblM4bTitlePreview.Width = Math.Max(200, totalWidth - 150);
        }

        // --- Custom Icon Button Factory & Renderers ---
        private Button CreateIconButton(Point location, Size size, Action<Graphics, Rectangle, bool> paintAction) {
            Button btn = new Button {
                Location = location,
                Size = size,
                Text = "",
                BackColor = SystemColors.Control,
                FlatStyle = FlatStyle.Standard,
                Cursor = Cursors.Hand
            };

            btn.Paint += (s, e) => {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle rect = new Rectangle(4, 3, btn.ClientSize.Width - 8, btn.ClientSize.Height - 6);
                paintAction(g, rect, btn.Enabled);
            };

            return btn;
        }

        private void DrawFolderIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(220, 160, 30) : Color.Gray;
            using (SolidBrush brush = new SolidBrush(c)) {
                // Tab
                g.FillRectangle(brush, r.X + 2, r.Y + 2, 8, 4);
                // Folder body
                g.FillRectangle(brush, r.X + 2, r.Y + 5, r.Width - 4, r.Height - 7);
            }
            using (Pen pen = new Pen(enabled ? Color.FromArgb(160, 110, 15) : Color.Gray, 1.2F)) {
                g.DrawRectangle(pen, r.X + 2, r.Y + 5, r.Width - 4, r.Height - 7);
            }
        }

        private void DrawSyncIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(40, 120, 200) : Color.Gray;
            using (Pen pen = new Pen(c, 2F)) {
                int cx = r.X + r.Width / 2;
                int cy = r.Y + r.Height / 2;
                int radius = Math.Min(r.Width, r.Height) / 2 - 2;
                g.DrawArc(pen, cx - radius, cy - radius, radius * 2, radius * 2, 45, 270);
                // Arrowhead
                g.DrawLine(pen, cx + radius - 3, cy - 2, cx + radius + 1, cy - 4);
                g.DrawLine(pen, cx + radius - 3, cy - 5, cx + radius + 1, cy - 4);
            }
        }

        private void DrawExplorerIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(60, 140, 60) : Color.Gray;
            using (Pen pen = new Pen(c, 1.8F)) {
                // Window frame
                g.DrawRectangle(pen, r.X + 2, r.Y + 2, r.Width - 5, r.Height - 5);
                // Window header
                g.DrawLine(pen, r.X + 2, r.Y + 7, r.Right - 3, r.Y + 7);
                // Small arrow pointing out
                g.DrawLine(pen, r.Right - 6, r.Bottom - 6, r.Right, r.Bottom);
                g.DrawLine(pen, r.Right - 4, r.Bottom, r.Right, r.Bottom);
                g.DrawLine(pen, r.Right, r.Bottom - 4, r.Right, r.Bottom);
            }
        }

        private void DrawUpIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(40, 120, 200) : Color.Gray;
            int cx = r.X + r.Width / 2;
            int cy = r.Y + r.Height / 2;
            using (SolidBrush brush = new SolidBrush(c)) {
                Point[] pts = new Point[] {
                    new Point(cx, cy - 5),
                    new Point(cx - 5, cy + 3),
                    new Point(cx + 5, cy + 3)
                };
                g.FillPolygon(brush, pts);
            }
        }

        private void DrawDownIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(40, 120, 200) : Color.Gray;
            int cx = r.X + r.Width / 2;
            int cy = r.Y + r.Height / 2;
            using (SolidBrush brush = new SolidBrush(c)) {
                Point[] pts = new Point[] {
                    new Point(cx, cy + 5),
                    new Point(cx - 5, cy - 3),
                    new Point(cx + 5, cy - 3)
                };
                g.FillPolygon(brush, pts);
            }
        }

        private void DrawForwardIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(40, 120, 200) : Color.Gray;
            int cx = r.X + r.Width / 2;
            int cy = r.Y + r.Height / 2;
            using (SolidBrush brush = new SolidBrush(c)) {
                // Left triangle
                Point[] pts1 = new Point[] {
                    new Point(cx - 5, cy - 5),
                    new Point(cx, cy),
                    new Point(cx - 5, cy + 5)
                };
                g.FillPolygon(brush, pts1);
                // Right triangle
                Point[] pts2 = new Point[] {
                    new Point(cx + 1, cy - 5),
                    new Point(cx + 6, cy),
                    new Point(cx + 1, cy + 5)
                };
                g.FillPolygon(brush, pts2);
            }
        }

        private void DrawBackIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(40, 120, 200) : Color.Gray;
            int cx = r.X + r.Width / 2;
            int cy = r.Y + r.Height / 2;
            using (SolidBrush brush = new SolidBrush(c)) {
                // Left triangle
                Point[] pts1 = new Point[] {
                    new Point(cx - 1, cy - 5),
                    new Point(cx - 6, cy),
                    new Point(cx - 1, cy + 5)
                };
                g.FillPolygon(brush, pts1);
                // Right triangle
                Point[] pts2 = new Point[] {
                    new Point(cx + 5, cy - 5),
                    new Point(cx, cy),
                    new Point(cx + 5, cy + 5)
                };
                g.FillPolygon(brush, pts2);
            }
        }

        private void DrawClearIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(190, 40, 40) : Color.Gray;
            using (Pen pen = new Pen(c, 2F)) {
                // Draw broom/trashcan outline
                g.DrawLine(pen, r.X + 2, r.Y + 4, r.Right - 2, r.Y + 4);
                g.DrawLine(pen, r.X + 5, r.Y + 4, r.X + 7, r.Bottom - 2);
                g.DrawLine(pen, r.Right - 5, r.Y + 4, r.Right - 7, r.Bottom - 2);
                g.DrawLine(pen, r.X + 7, r.Bottom - 2, r.Right - 7, r.Bottom - 2);
                // Lid handle
                g.DrawLine(pen, r.X + r.Width / 2 - 3, r.Y + 1, r.X + r.Width / 2 + 3, r.Y + 1);
            }
        }

        private void DrawWandIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(30, 110, 200) : Color.Gray;
            using (Pen pen = new Pen(c, 2.2F)) {
                // Wand diagonal stick
                g.DrawLine(pen, r.X + 3, r.Bottom - 3, r.Right - 5, r.Y + 5);
            }
            // Sparkles at tip
            using (SolidBrush brush = new SolidBrush(enabled ? Color.FromArgb(240, 170, 0) : Color.Gray)) {
                int tx = r.Right - 4;
                int ty = r.Y + 4;
                g.FillPolygon(brush, new Point[] {
                    new Point(tx, ty - 4),
                    new Point(tx + 2, ty - 1),
                    new Point(tx + 5, ty),
                    new Point(tx + 2, ty + 1),
                    new Point(tx, ty + 4),
                    new Point(tx - 2, ty + 1),
                    new Point(tx - 5, ty),
                    new Point(tx - 2, ty - 1)
                });
            }
        }

        private void DrawCopyIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(50, 130, 60) : Color.Gray;
            using (Pen pen = new Pen(c, 1.8F)) {
                // Back page
                g.DrawRectangle(pen, r.X + 5, r.Y + 2, r.Width - 8, r.Height - 7);
                // Front page
                using (SolidBrush bg = new SolidBrush(SystemColors.Control)) {
                    g.FillRectangle(bg, r.X + 2, r.Y + 6, r.Width - 8, r.Height - 7);
                }
                g.DrawRectangle(pen, r.X + 2, r.Y + 6, r.Width - 8, r.Height - 7);
            }
        }

        private void DrawPasteIcon(Graphics g, Rectangle r, bool enabled) {
            Color c = enabled ? Color.FromArgb(130, 80, 160) : Color.Gray;
            using (Pen pen = new Pen(c, 1.8F)) {
                // Clipboard board
                g.DrawRectangle(pen, r.X + 3, r.Y + 4, r.Width - 6, r.Height - 6);
                // Clip top
                g.DrawLine(pen, r.X + r.Width / 2 - 4, r.Y + 2, r.X + r.Width / 2 + 4, r.Y + 2);
                g.DrawLine(pen, r.X + r.Width / 2 - 4, r.Y + 2, r.X + r.Width / 2 - 4, r.Y + 5);
                g.DrawLine(pen, r.X + r.Width / 2 + 4, r.Y + 2, r.X + r.Width / 2 + 4, r.Y + 5);
            }
            // Arrow pointing down into clipboard
            using (Pen pen = new Pen(enabled ? Color.FromArgb(200, 100, 30) : Color.Gray, 2F)) {
                int cx = r.X + r.Width / 2;
                int cy = r.Y + r.Height / 2 + 1;
                g.DrawLine(pen, cx, cy - 4, cx, cy + 3);
                g.DrawLine(pen, cx - 3, cy, cx, cy + 3);
                g.DrawLine(pen, cx + 3, cy, cx, cy + 3);
            }
        }

        private Button CreateInstagramButton(Point location, Size size) {
            Button btn = new Button {
                Location = location,
                Size = size,
                Text = "",
                BackColor = Color.White,
                FlatStyle = FlatStyle.Standard,
                Cursor = Cursors.Hand
            };

            btn.Paint += (s, e) => {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                
                int w = btn.ClientSize.Width;
                int h = btn.ClientSize.Height;
                int pad = 4;
                Rectangle rect = new Rectangle(pad, pad, w - pad * 2, h - pad * 2);
                
                // Draw rounded gradient camera outline
                using (LinearGradientBrush brush = new LinearGradientBrush(rect, Color.FromArgb(225, 48, 108), Color.FromArgb(253, 29, 29), 45F)) {
                    using (Pen pen = new Pen(brush, 2.0F)) {
                        int r = 4;
                        using (GraphicsPath path = new GraphicsPath()) {
                            path.AddArc(rect.X, rect.Y, r * 2, r * 2, 180, 90);
                            path.AddArc(rect.Right - r * 2, rect.Y, r * 2, r * 2, 270, 90);
                            path.AddArc(rect.Right - r * 2, rect.Bottom - r * 2, r * 2, r * 2, 0, 90);
                            path.AddArc(rect.X, rect.Bottom - r * 2, r * 2, r * 2, 90, 90);
                            path.CloseFigure();
                            g.DrawPath(pen, path);
                        }

                        // Center circle
                        int cx = rect.X + rect.Width / 2;
                        int cy = rect.Y + rect.Height / 2;
                        int cr = (int)(rect.Width * 0.28);
                        g.DrawEllipse(pen, cx - cr, cy - cr, cr * 2, cr * 2);

                        // Flash dot
                        int dotR = (int)(rect.Width * 0.08);
                        int dotX = rect.Right - dotR * 3 - 1;
                        int dotY = rect.Y + dotR * 2;
                        using (SolidBrush dotBrush = new SolidBrush(Color.FromArgb(225, 48, 108))) {
                            g.FillEllipse(dotBrush, dotX, dotY, dotR * 2, dotR * 2);
                        }
                    }
                }
            };

            return btn;
        }

        private void BtnBrowse_Click(object sender, EventArgs e) {
            using (var dialog = new FolderBrowserDialog()) {
                dialog.Description = "Select the folder containing MP3 files";
                if (!string.IsNullOrEmpty(_selectedFolder) && Directory.Exists(_selectedFolder)) {
                    dialog.SelectedPath = _selectedFolder;
                }
                
                if (dialog.ShowDialog(this) == DialogResult.OK) {
                    LoadDirectory(dialog.SelectedPath);
                    _settings.DefaultDirectory = dialog.SelectedPath;
                    _settings.Save();
                }
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(_selectedFolder) && Directory.Exists(_selectedFolder)) {
                LoadDirectory(_selectedFolder);
            } else {
                BtnBrowse_Click(sender, e);
            }
        }

        private void BtnOpenExplorer_Click(object sender, EventArgs e) {
            if (!string.IsNullOrEmpty(_selectedFolder) && Directory.Exists(_selectedFolder)) {
                try {
                    Process.Start("explorer.exe", _selectedFolder);
                } catch (Exception ex) {
                    MessageBox.Show("Could not open folder in Explorer: " + ex.Message, "Explorer Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            } else {
                MessageBox.Show("Please select a folder first.", "Open in Explorer", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void LoadDirectory(string path) {
            _selectedFolder = path;
            lblCurrentFolder.Text = path;
            
            listTagFiles.Items.Clear();
            listAvailableMp3s.Items.Clear();

            // Retain existing chapters whose source files still exist on disk
            for (int i = listChapters.Items.Count - 1; i >= 0; i--) {
                ChapterItem ci = listChapters.Items[i] as ChapterItem;
                if (ci == null || !File.Exists(ci.FilePath)) {
                    listChapters.Items.RemoveAt(i);
                }
            }

            try {
                var mp3s = Directory.GetFiles(_selectedFolder, "*.mp3").OrderBy(f => f).ToArray();
                foreach (var file in mp3s) {
                    string name = Path.GetFileName(file);
                    listTagFiles.Items.Add(name);
                    listAvailableMp3s.Items.Add(name);
                }
            } catch (Exception ex) {
                MessageBox.Show("Error loading files: " + ex.Message, "Folder Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            btnExecuteSave.Enabled = false;
            btnCopyTags.Enabled = false;
            btnGuessFilename.Enabled = false;
            btnClearTags.Enabled = false;
            lblHelp.Text = "Select one or more MP3s to tag and rename.";
            UpdateTitlePreview();
        }

        // --- Cover Helpers ---
        private void SetTagCover(string path) {
            _tagCoverPath = path ?? "";
            UpdateCoverPreviewBox(_tagCoverPath, picTagCoverPreview);
        }

        private void SetM4bCover(string path) {
            _m4bCoverPath = path ?? "";
            UpdateCoverPreviewBox(_m4bCoverPath, picCoverPreview);
        }

        private void PromptTagCoverFile() {
            using (var dialog = new OpenFileDialog { Filter = "Image Files (*.jpg;*.jpeg;*.png;*.webp;*.bmp)|*.jpg;*.jpeg;*.png;*.webp;*.bmp", Title = "Select Cover Image" }) {
                if (dialog.ShowDialog(this) == DialogResult.OK) {
                    SetTagCover(dialog.FileName);
                }
            }
        }

        private void PromptM4bCoverFile() {
            using (var dialog = new OpenFileDialog { Filter = "Image Files (*.jpg;*.jpeg;*.png;*.webp;*.bmp)|*.jpg;*.jpeg;*.png;*.webp;*.bmp", Title = "Select Cover Image" }) {
                if (dialog.ShowDialog(this) == DialogResult.OK) {
                    SetM4bCover(dialog.FileName);
                }
            }
        }

        private void SetupListDragAndDrop(ListBox listBox) {
            listBox.AllowDrop = true;
            listBox.DragEnter += (s, e) => {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                    e.Effect = DragDropEffects.Copy;
                } else {
                    e.Effect = DragDropEffects.None;
                }
            };
            listBox.DragDrop += (s, e) => {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files != null && files.Length > 0) {
                        string target = files[0];
                        if (Directory.Exists(target)) {
                            LoadDirectory(target);
                            _settings.DefaultDirectory = target;
                            _settings.Save();
                        } else if (File.Exists(target)) {
                            string dir = Path.GetDirectoryName(target);
                            if (Directory.Exists(dir)) {
                                LoadDirectory(dir);
                                _settings.DefaultDirectory = dir;
                                _settings.Save();
                            }
                        }
                    }
                }
            };
        }

        public static TagLib.File CreateTagFileWithRetry(string path, int maxAttempts = 4, int delayMs = 60) {
            for (int i = 1; i <= maxAttempts; i++) {
                try {
                    return TagLib.File.Create(path);
                } catch (IOException) {
                    if (i < maxAttempts) {
                        System.Threading.Thread.Sleep(delayMs * i);
                    } else {
                        throw;
                    }
                }
            }
            return TagLib.File.Create(path);
        }

        public static void SaveTagFileWithRetry(TagLib.File tagFile, int maxAttempts = 4, int delayMs = 60) {
            for (int i = 1; i <= maxAttempts; i++) {
                try {
                    tagFile.Save();
                    return;
                } catch (IOException) {
                    if (i < maxAttempts) {
                        System.Threading.Thread.Sleep(delayMs * i);
                    } else {
                        throw;
                    }
                }
            }
            tagFile.Save();
        }

        private void SetupCoverDragAndDrop(Control control, Action<string> onFileChosen) {
            control.AllowDrop = true;
            control.DragEnter += (s, e) => {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                    e.Effect = DragDropEffects.Copy;
                } else {
                    e.Effect = DragDropEffects.None;
                }
            };
            control.DragDrop += (s, e) => {
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files != null && files.Length > 0) {
                        string f = files[0];
                        string ext = Path.GetExtension(f).ToLowerInvariant();
                        if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".webp" || ext == ".bmp") {
                            onFileChosen(f);
                        }
                    }
                }
            };
        }

        private void UpdateCoverPreviewBox(string path, PictureBox picBox) {
            if (File.Exists(path)) {
                try {
                    using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                        using (var original = Image.FromStream(fs)) {
                            var bmp = new Bitmap(original.Width, original.Height);
                            using (var g = Graphics.FromImage(bmp)) {
                                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                g.SmoothingMode = SmoothingMode.HighQuality;
                                g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height));
                            }
                            if (picBox.Image != null) picBox.Image.Dispose();
                            picBox.Image = bmp;
                        }
                    }
                } catch {
                    if (picBox.Image != null) picBox.Image.Dispose();
                    picBox.Image = null;
                }
            } else {
                if (picBox.Image != null) picBox.Image.Dispose();
                picBox.Image = null;
            }
        }

        private void CycleCarousel(List<string> imageList, ref int indexRef, int delta, Action<string> onSelect) {
            if (imageList == null || imageList.Count == 0) return;
            indexRef = (indexRef + delta + imageList.Count) % imageList.Count;
            onSelect(imageList[indexRef]);
        }

        private void OpenCarouselModal(List<string> imageList, int currentIndex, Action<string> onSelect, Action<int> setIndex) {
            if (imageList == null || imageList.Count == 0) return;
            using (var dlg = new ImageCarouselDialog(imageList, currentIndex)) {
                if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrEmpty(dlg.SelectedImagePath)) {
                    setIndex(dlg.SelectedIndex);
                    onSelect(dlg.SelectedImagePath);
                }
            }
        }

        // --- Tagging Logic ---
        private void ListTagFiles_SelectedIndexChanged(object sender, EventArgs e) {
            if (_isBatchUpdating) return;

            int count = listTagFiles.SelectedItems.Count;
            if (count > 0) {
                btnExecuteSave.Enabled = true;
                btnClearTags.Enabled = true;
                btnPasteTags.Enabled = _hasClipboardTags;

                if (count > 1) {
                    lblHelp.Text = string.Format("{0} files selected for batch tagging.", count);
                    // Disable From Filename & Copy Tags on multi-select
                    btnGuessFilename.Enabled = false;
                    btnCopyTags.Enabled = false;
                    UpdateTitlePreview();
                    return;
                }

                // Single selection
                btnGuessFilename.Enabled = true;

                string fileName = listTagFiles.SelectedItem.ToString();
                string fullPath = Path.Combine(_selectedFolder, fileName);
                
                // Clear fields first for single item inspection
                SetArtistFields(new string[0]);
                txtShow.Text = "";
                cmbLocation.Text = "";
                SetTagCover("");

                bool hasTags = false;
                
                try {
                    using (TagLib.File tagFile = CreateTagFileWithRetry(fullPath)) {
                        string[] performers = tagFile.Tag.Performers;
                        string tArtist = tagFile.Tag.FirstPerformer;
                        if (string.IsNullOrEmpty(tArtist)) tArtist = tagFile.Tag.FirstAlbumArtist;
                        if (string.IsNullOrEmpty(tArtist)) tArtist = tagFile.Tag.JoinedPerformers;

                        string tAlbum = tagFile.Tag.Album;
                        string tTitle = tagFile.Tag.Title;
                        uint tYear = tagFile.Tag.Year;

                        if (performers != null && performers.Length > 0) {
                            SetArtistFields(performers);
                            hasTags = true;
                        } else if (!string.IsNullOrEmpty(tArtist)) {
                            SetArtistFields(new[] { tArtist });
                            hasTags = true;
                        }
                        if (!string.IsNullOrEmpty(tAlbum)) {
                            txtShow.Text = tAlbum;
                            hasTags = true;
                        }
                        
                        if (!string.IsNullOrEmpty(tTitle)) {
                            var parts = tTitle.Split(new[] { " - " }, StringSplitOptions.None);
                            if (parts.Length >= 2) {
                                DateTime d;
                                if (DateTime.TryParse(parts[0], out d)) dtpDate.Value = d;
                                cmbLocation.Text = parts[parts.Length - 1];
                            }
                            hasTags = true;
                        }

                        if (tYear > 0 && tYear <= 9999) {
                            try {
                                int safeMonth = dtpDate.Value.Month;
                                int maxDays = DateTime.DaysInMonth((int)tYear, safeMonth);
                                int safeDay = Math.Min(dtpDate.Value.Day, maxDays);
                                dtpDate.Value = new DateTime((int)tYear, safeMonth, safeDay);
                            } catch {
                                try { dtpDate.Value = new DateTime((int)tYear, 1, 1); } catch { }
                            }
                        }
                    }
                } catch { }

                if (hasTags) {
                    lblHelp.Text = "1 file selected (ID3 tags loaded).";
                } else {
                    lblHelp.Text = "1 file selected (No ID3 tags found).";
                }

                btnCopyTags.Enabled = hasTags;
                UpdateTitlePreview();
            } else {
                lblHelp.Text = "Select one or more MP3s to tag and rename.";
                btnExecuteSave.Enabled = false;
                btnCopyTags.Enabled = false;
                btnGuessFilename.Enabled = false;
                btnClearTags.Enabled = false;
                btnPasteTags.Enabled = false;
                UpdateTitlePreview();
            }
        }

        private void BtnGuessFilename_Click(object sender, EventArgs e) {
            if (listTagFiles.SelectedItem == null || listTagFiles.SelectedItems.Count > 1) return;
            string fileName = listTagFiles.SelectedItem.ToString();

            DateTime? parsedDate;
            string parsedArtist, parsedShow, parsedLocation;

            if (_settings.ParseFromFilename(fileName, out parsedDate, out parsedArtist, out parsedShow, out parsedLocation)) {
                if (parsedDate.HasValue) dtpDate.Value = parsedDate.Value;
                if (!string.IsNullOrEmpty(parsedArtist)) SetArtistFields(new[] { parsedArtist });
                if (!string.IsNullOrEmpty(parsedShow)) txtShow.Text = parsedShow;
                if (!string.IsNullOrEmpty(parsedLocation)) cmbLocation.Text = parsedLocation;
                UpdateTitlePreview();
            } else {
                MessageBox.Show("Could not extract metadata from this filename using active naming conventions.", "Filename Parse", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        public static string SanitizeFilenameSegment(string segment) {
            if (string.IsNullOrWhiteSpace(segment)) return "";
            string s = segment.Trim();
            
            // Replace colon with dash
            s = s.Replace(" : ", " - ").Replace(": ", " - ").Replace(":", "-");
            
            // Replace slashes and backslashes with dash
            s = s.Replace("/", "-").Replace("\\", "-");
            
            // Replace double quotes with single quotes
            s = s.Replace("\"", "'");
            
            // Replace angle brackets with square brackets
            s = s.Replace("<", "[").Replace(">", "]");
            
            // Replace pipe with dash
            s = s.Replace("|", "-");

            // Replace ampersand with plus symbol
            s = s.Replace("&", "+");
            
            // Strip invalid characters like ? * and control characters
            char[] invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (char c in s) {
                if (!invalid.Contains(c)) {
                    sb.Append(c);
                }
            }
            
            // Collapse multiple consecutive spaces or dashes cleanly
            string result = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
            result = Regex.Replace(result, @"\s*-\s*-\s*", " - ");
            return result;
        }

        private void UpdateTitlePreview() {
            if (listTagFiles.SelectedItems.Count == 0) {
                lblTitlePreview.Text = "";
                return;
            }
            
            if (chkRenameMp3 != null && !chkRenameMp3.Checked) {
                lblTitlePreview.Text = "(Filenames will be preserved untouched)";
                return;
            }

            DateTime dt = dtpDate.Value;
            
            // Inclusion checkboxes determine filename parts
            List<string> enteredArtists = GetAllEnteredArtists();
            List<string> sanitizedArtists = new List<string>();
            foreach (var a in enteredArtists) {
                string s = SanitizeFilenameSegment(a);
                if (!string.IsNullOrWhiteSpace(s)) sanitizedArtists.Add(s);
            }
            string combinedArtist = FormatArtistsForFilename(sanitizedArtists, _settings);
            string artist = chkIncArtist.Checked ? combinedArtist : "";
            string show = chkIncShow.Checked ? SanitizeFilenameSegment(txtShow.Text.Trim()) : "";
            string location = chkIncLocation.Checked ? SanitizeFilenameSegment(cmbLocation.Text.Trim()) : "";
            
            if (chkIncArtist.Checked && string.IsNullOrWhiteSpace(artist)) artist = "[Artist]";
            if (chkIncLocation.Checked && string.IsNullOrWhiteSpace(location)) location = "[Location]";

            string previewName;
            if (listTagFiles.SelectedItems.Count > 1) {
                previewName = _settings.FormatFilename(dt, artist, show, location, "track_name");
            } else {
                previewName = _settings.FormatFilename(dt, artist, show, location, "");
            }

            // If date is omitted by checkbox, remove date segment
            if (!chkIncDate.Checked) {
                string formattedDate = dt.ToString(_settings.DateFormat ?? "yyyy-MM-dd");
                previewName = previewName.Replace(formattedDate, "").Trim();
                previewName = Regex.Replace(previewName, @"^(\s*-\s*|\s*_\s*|\s*\.\s*)+", "").Trim();
            }
            
            lblTitlePreview.Text = previewName + ".mp3";
        }

        private void UpdateM4bPreview() {
            if (lblM4bTitlePreview == null) return;

            DateTime dt = dtpM4bDate.Value;
            string artist = SanitizeFilenameSegment(cmbM4bArtist.Text.Trim());
            string album = SanitizeFilenameSegment(txtM4bAlbum.Text.Trim());
            string location = SanitizeFilenameSegment(cmbM4bLocation.Text.Trim());

            if (string.IsNullOrWhiteSpace(artist)) artist = "Various";
            if (string.IsNullOrWhiteSpace(album)) album = "[Album]";

            string formatted = _settings.FormatFilename(dt, artist, album, location, "");
            lblM4bTitlePreview.Text = formatted + ".m4b";
        }

        private string ExtractCleanTrackIdentifier(string fileName) {
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            // Check if name ends with parentheses: "... (Track01)"
            var match = Regex.Match(baseName, @"\((?<inner>[^()]+)\)$");
            if (match.Success) {
                return match.Groups["inner"].Value.Trim();
            }
            // If formatted as "YYYY-MM-DD - Artist - Show - Location", take the last segment or original
            var parts = baseName.Split(new[] { " - " }, StringSplitOptions.None);
            if (parts.Length >= 4) {
                return parts[parts.Length - 1].Trim();
            }
            return baseName;
        }

        private void BtnCopyTags_Click(object sender, EventArgs e) {
            _clipDate = dtpDate.Value;
            _clipArtists = GetAllEnteredArtists();
            _clipShow = txtShow.Text;
            _clipLocation = cmbLocation.Text;
            _hasClipboardTags = true;
            btnPasteTags.Enabled = true;
        }

        private void BtnPasteTags_Click(object sender, EventArgs e) {
            if (!_hasClipboardTags) return;
            dtpDate.Value = _clipDate;
            SetArtistFields(_clipArtists);
            txtShow.Text = _clipShow;
            cmbLocation.Text = _clipLocation;
            UpdateTitlePreview();
        }

        private void BtnClearTags_Click(object sender, EventArgs e) {
            int selectedCount = listTagFiles.SelectedItems.Count;
            if (selectedCount > 0) {
                string promptMsg = _settings.ClearErasesAllTags
                    ? string.Format("Do you want to strip all ID3 tags from the {0} selected file(s) and clear the input fields?", selectedCount)
                    : string.Format("Do you want to reset ComComTag ID3 tags (Artist, Show, Venue, Date, Cover) from the {0} selected file(s) and clear the input fields?", selectedCount);

                var dr = MessageBox.Show(promptMsg, "Clear Tags Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dr == DialogResult.Yes) {
                    List<string> selectedNames = new List<string>();
                    foreach (var item in listTagFiles.SelectedItems) selectedNames.Add(item.ToString());

                    int successCount = 0;
                    foreach (var fileName in selectedNames) {
                        string fullPath = Path.Combine(_selectedFolder, fileName);
                        if (File.Exists(fullPath)) {
                            try {
                                using (TagLib.File tagFile = CreateTagFileWithRetry(fullPath)) {
                                    if (_settings.ClearErasesAllTags) {
                                        tagFile.Tag.Clear();
                                        tagFile.RemoveTags(TagLib.TagTypes.AllTags);
                                    } else {
                                        // Wipe only ComComTag managed fields
                                        tagFile.Tag.Performers = new string[0];
                                        tagFile.Tag.AlbumArtists = new string[0];
                                        tagFile.Tag.Title = "";
                                        tagFile.Tag.Album = "";
                                        tagFile.Tag.Year = 0;
                                        tagFile.Tag.Pictures = new TagLib.IPicture[0];
                                    }
                                    SaveTagFileWithRetry(tagFile);
                                }
                                successCount++;
                            } catch (Exception ex) {
                                Debug.WriteLine("Failed to clear tags on " + fileName + ": " + ex.Message);
                            }
                        }
                    }

                    ResetTaggingFields();
                    
                    _isBatchUpdating = true;
                    try {
                        LoadDirectory(_selectedFolder);
                        for (int i = 0; i < listTagFiles.Items.Count; i++) {
                            string item = listTagFiles.Items[i].ToString();
                            if (selectedNames.Contains(item)) {
                                listTagFiles.SetSelected(i, true);
                            }
                        }
                    } finally {
                        _isBatchUpdating = false;
                    }

                    lblHelp.Text = string.Format("ID3 tags cleared on {0} file(s).", successCount);
                    MessageBox.Show(string.Format("Successfully cleared tags on {0} file(s)!", successCount), "Tags Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }

            // Reset fields only
            ResetTaggingFields();
        }

        private void ResetTaggingFields() {
            dtpDate.Value = DateTime.Today;
            SetArtistFields(new string[0]);
            txtShow.Text = "";
            cmbLocation.Text = "";
            SetTagCover("");
            btnTagInstaPrev.Visible = false;
            btnTagInstaNext.Visible = false;
            _tagInstaImages.Clear();
            _tagInstaImageIndex = 0;
            btnCopyTags.Enabled = false;
            UpdateTitlePreview();
        }

        private void BtnExecuteSave_Click(object sender, EventArgs e) {
            if (listTagFiles.SelectedItems.Count == 0) return;

            bool shouldRename = chkRenameMp3.Checked;

            DateTime dt = dtpDate.Value;
            string dateStr = dt.ToString("yyyy-MM-dd");
            List<string> artists = GetAllEnteredArtists();
            List<string> sanitizedArtists = new List<string>();
            foreach (var a in artists) {
                string s = SanitizeFilenameSegment(a);
                if (!string.IsNullOrWhiteSpace(s)) sanitizedArtists.Add(s);
            }
            string combinedArtist = FormatArtistsForFilename(sanitizedArtists, _settings);
            string show = txtShow.Text.Trim();
            string location = cmbLocation.Text.Trim();
            uint yearValue = (uint)dt.Year;

            if (artists.Count == 0 && string.IsNullOrWhiteSpace(location)) {
                MessageBox.Show("Please provide at least an Artist or a Venue/Location.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Auto-learn new Artists and Location into Settings if enabled
            if (_settings.AutoAddArtists) {
                foreach (var a in artists) {
                    _settings.AddArtistIfNew(a);
                }
            }
            if (_settings.AutoAddLocations && !string.IsNullOrWhiteSpace(location)) _settings.AddLocationIfNew(location);
            RefreshDropdownLists();

            // Prepare safe segments respecting inclusion checkboxes
            string safeArtist = chkIncArtist.Checked ? combinedArtist : "";
            string safeShow = chkIncShow.Checked ? SanitizeFilenameSegment(show) : "";
            string safeLocation = chkIncLocation.Checked ? SanitizeFilenameSegment(location) : "";

            List<string> selectedNames = new List<string>();
            foreach (var item in listTagFiles.SelectedItems) selectedNames.Add(item.ToString());

            int updatedCount = 0;
            List<string> finalNames = new List<string>();

            for (int i = 0; i < selectedNames.Count; i++) {
                string fileName = selectedNames[i];
                string sourcePath = Path.Combine(_selectedFolder, fileName);

                if (!File.Exists(sourcePath)) continue;

                string destPath = sourcePath;
                string newName = fileName;

                if (shouldRename) {
                    if (selectedNames.Count > 1) {
                        string trackIdentifier = ExtractCleanTrackIdentifier(fileName);
                        string safeTrack = SanitizeFilenameSegment(trackIdentifier);
                        newName = _settings.FormatFilename(dt, safeArtist, safeShow, safeLocation, safeTrack) + ".mp3";
                    } else {
                        newName = _settings.FormatFilename(dt, safeArtist, safeShow, safeLocation, "") + ".mp3";
                    }

                    // If date is omitted by checkbox, remove date segment
                    if (!chkIncDate.Checked) {
                        string formattedDate = dt.ToString(_settings.DateFormat ?? "yyyy-MM-dd");
                        newName = newName.Replace(formattedDate, "").Trim();
                        newName = Regex.Replace(newName, @"^(\s*-\s*|\s*_\s*|\s*\.\s*)+", "").Trim();
                        if (!newName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase)) newName += ".mp3";
                    }

                    destPath = Path.Combine(_selectedFolder, newName);

                    try {
                        // Physical Rename / Overwrite Protection
                        bool isSameFileDifferentCase = string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase) 
                                                   && !string.Equals(sourcePath, destPath, StringComparison.Ordinal);

                        if (isSameFileDifferentCase) {
                            string intermediatePath = sourcePath + ".tmp_" + Guid.NewGuid().ToString("N");
                            File.Move(sourcePath, intermediatePath);
                            File.Move(intermediatePath, destPath);
                        } else if (!string.Equals(sourcePath, destPath, StringComparison.OrdinalIgnoreCase)) {
                            if (File.Exists(destPath)) {
                                DialogResult dr = MessageBox.Show(
                                    string.Format("The file '{0}' already exists. Overwrite it?", newName), 
                                    "File Exists", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                if (dr == DialogResult.No) continue;
                                File.Delete(destPath);
                            }
                            File.Move(sourcePath, destPath);
                        }
                    } catch (Exception ex) {
                        MessageBox.Show(string.Format("Error renaming '{0}':\n{1}", fileName, ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }
                }

                try {
                    // Write ID3 Metadata via TagLib# with retry wrapper
                    using (TagLib.File tagFile = CreateTagFileWithRetry(destPath)) {
                        // Title: YYYY-MM-DD - [Show - ] Location
                        if (string.IsNullOrWhiteSpace(show)) {
                            tagFile.Tag.Title = string.Format("{0} - {1}", dateStr, location);
                        } else {
                            tagFile.Tag.Title = string.Format("{0} - {1} - {2}", dateStr, show, location);
                        }

                        // Artist / Performers & Album Artist
                        if (artists.Count > 0) {
                            tagFile.Tag.Performers = artists.ToArray();
                            if (artists.Count == 1) {
                                tagFile.Tag.AlbumArtists = new[] { artists[0] };
                            } else {
                                if (_settings.UseFirstArtistAsAlbumArtist) {
                                    tagFile.Tag.AlbumArtists = new[] { artists[0] };
                                } else {
                                    tagFile.Tag.AlbumArtists = new[] { "Various Artists" };
                                }
                            }
                        }

                        // Album / Show
                        tagFile.Tag.Album = show ?? "";
                        tagFile.Tag.Year = yearValue;

                        // Cover Art
                        if (!string.IsNullOrEmpty(_tagCoverPath) && File.Exists(_tagCoverPath)) {
                            tagFile.Tag.Pictures = new TagLib.IPicture[] { new TagLib.Picture(_tagCoverPath) };
                        }

                        SaveTagFileWithRetry(tagFile);
                    }

                    finalNames.Add(newName);
                    updatedCount++;
                } catch (Exception ex) {
                    MessageBox.Show(string.Format("Error writing tags to '{0}':\n{1}", fileName, ex.Message), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Reload and reselect modified files
            _isBatchUpdating = true;
            try {
                LoadDirectory(_selectedFolder);
                for (int i = 0; i < listTagFiles.Items.Count; i++) {
                    string item = listTagFiles.Items[i].ToString();
                    if (finalNames.Contains(item)) {
                        listTagFiles.SetSelected(i, true);
                    }
                }
            } finally {
                _isBatchUpdating = false;
            }

            string successMessage = shouldRename 
                ? string.Format("Successfully updated tags and renamed {0} file(s)!", updatedCount)
                : string.Format("Successfully updated ID3 tags for {0} file(s) (Filenames preserved)!", updatedCount);

            MessageBox.Show(successMessage, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                
                if (_settings.AppendChapterDuration) {
                    try {
                        long ms = M4bBuilder.GetDurationMs(fullPath);
                        if (ms > 0) {
                            TimeSpan d = TimeSpan.FromMilliseconds(ms);
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

        private void SetupChapterInlineEditing() {
            _inlineChapterEditBox = new TextBox {
                Visible = false,
                BorderStyle = BorderStyle.FixedSingle
            };
            _inlineChapterEditBox.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter) {
                    CommitChapterInlineEdit();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                } else if (e.KeyCode == Keys.Escape) {
                    CancelChapterInlineEdit();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            };
            _inlineChapterEditBox.Leave += (s, e) => {
                CommitChapterInlineEdit();
            };
            listChapters.Controls.Add(_inlineChapterEditBox);

            listChapters.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.F2 && listChapters.SelectedIndex >= 0) {
                    StartChapterInlineEdit(listChapters.SelectedIndex);
                    e.Handled = true;
                } else if ((e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back) && listChapters.SelectedItems.Count > 0) {
                    BtnRemoveChapter_Click(s, e);
                    e.Handled = true;
                } else if (e.Control && e.KeyCode == Keys.Up) {
                    BtnMoveUp_Click(s, e);
                    e.Handled = true;
                } else if (e.Control && e.KeyCode == Keys.Down) {
                    BtnMoveDown_Click(s, e);
                    e.Handled = true;
                }
            };

            listChapters.DoubleClick += (s, e) => {
                if (listChapters.SelectedIndex >= 0) {
                    StartChapterInlineEdit(listChapters.SelectedIndex);
                }
            };

            listChapters.MouseMove += (s, e) => {
                int idx = listChapters.IndexFromPoint(e.Location);
                if (idx >= 0 && idx < listChapters.Items.Count) {
                    ChapterItem item = listChapters.Items[idx] as ChapterItem;
                    if (item != null) {
                        string tip = "Original File: " + Path.GetFileName(item.FilePath);
                        if (toolTip.GetToolTip(listChapters) != tip) {
                            toolTip.SetToolTip(listChapters, tip);
                        }
                    }
                } else {
                    toolTip.SetToolTip(listChapters, "");
                }
            };
        }

        private void StartChapterInlineEdit(int index) {
            if (index < 0 || index >= listChapters.Items.Count) return;
            ChapterItem item = listChapters.Items[index] as ChapterItem;
            if (item == null) return;

            _editingChapterIndex = index;
            Rectangle rect = listChapters.GetItemRectangle(index);

            if (ThemeHelper.ShouldUseDarkMode(_settings.Theme)) {
                _inlineChapterEditBox.BackColor = ThemeHelper.DarkControlBg;
                _inlineChapterEditBox.ForeColor = ThemeHelper.DarkControlText;
            } else {
                _inlineChapterEditBox.BackColor = SystemColors.Window;
                _inlineChapterEditBox.ForeColor = SystemColors.WindowText;
            }

            _inlineChapterEditBox.Text = item.ChapterName;
            _inlineChapterEditBox.Bounds = new Rectangle(rect.X + 2, rect.Y + 1, rect.Width - 4, rect.Height + 2);
            _inlineChapterEditBox.Visible = true;
            _inlineChapterEditBox.BringToFront();
            _inlineChapterEditBox.Focus();
            _inlineChapterEditBox.SelectAll();
        }

        private void CommitChapterInlineEdit() {
            if (!_inlineChapterEditBox.Visible) return;
            if (_editingChapterIndex >= 0 && _editingChapterIndex < listChapters.Items.Count) {
                ChapterItem item = listChapters.Items[_editingChapterIndex] as ChapterItem;
                if (item != null && !string.IsNullOrWhiteSpace(_inlineChapterEditBox.Text)) {
                    item.ChapterName = _inlineChapterEditBox.Text.Trim();
                    listChapters.Items[_editingChapterIndex] = item;
                }
            }
            _editingChapterIndex = -1;
            _inlineChapterEditBox.Visible = false;
        }

        private void CancelChapterInlineEdit() {
            _editingChapterIndex = -1;
            if (_inlineChapterEditBox != null) _inlineChapterEditBox.Visible = false;
        }

        private void BtnRemoveChapter_Click(object sender, EventArgs e) {
            CancelChapterInlineEdit();
            while (listChapters.SelectedItems.Count > 0) {
                listChapters.Items.Remove(listChapters.SelectedItems[0]);
            }
        }

        private void BtnMoveUp_Click(object sender, EventArgs e) {
            CancelChapterInlineEdit();
            if (listChapters.SelectedItem == null || listChapters.SelectedIndex <= 0) return;
            int newIndex = listChapters.SelectedIndex - 1;
            object selected = listChapters.SelectedItem;
            listChapters.Items.Remove(selected);
            listChapters.Items.Insert(newIndex, selected);
            listChapters.SetSelected(newIndex, true);
        }

        private void BtnMoveDown_Click(object sender, EventArgs e) {
            CancelChapterInlineEdit();
            if (listChapters.SelectedItem == null || listChapters.SelectedIndex >= listChapters.Items.Count - 1) return;
            int newIndex = listChapters.SelectedIndex + 1;
            object selected = listChapters.SelectedItem;
            listChapters.Items.Remove(selected);
            listChapters.Items.Insert(newIndex, selected);
            listChapters.SetSelected(newIndex, true);
        }

        // --- Instagram Downloader Bridge (100% Native C#) ---
        private void HandleInstaDownload(List<string> imageList, Action<int> setIndex, Action<string> onSelect, Button prevBtn, Button nextBtn, Button downloadBtn) {
            string url = InputPromptDialog.Prompt(this, "Download Cover from Instagram", "Paste an Instagram Post URL:", "", _settings.Theme);
            if (string.IsNullOrWhiteSpace(url)) return;

            RunInstaDownloadAsync(url.Trim(), imageList, setIndex, onSelect, prevBtn, nextBtn, downloadBtn);
        }

        private async void RunInstaDownloadAsync(string url, List<string> imageList, Action<int> setIndex, Action<string> onSelect, Button prevBtn, Button nextBtn, Button downloadBtn) {
            this.Cursor = Cursors.WaitCursor;
            downloadBtn.Enabled = false;

            string targetDir = _settings.GetResolvedInstaTempDir();

            try {
                List<string> downloaded = await Task.Run(() => {
                    return InstagramDownloader.DownloadPostImages(url, targetDir);
                });

                imageList.Clear();
                imageList.AddRange(downloaded);

                if (imageList.Count > 0) {
                    setIndex(0);
                    onSelect(imageList[0]);

                    bool multi = imageList.Count > 1;
                    prevBtn.Visible = multi;
                    nextBtn.Visible = multi;

                    // Automatically open visual Carousel Picker modal if multiple images were found
                    if (multi) {
                        OpenCarouselModal(imageList, 0, onSelect, setIndex);
                    }
                }
            } catch (Exception ex) {
                MessageBox.Show("Error downloading Instagram art: " + ex.Message, "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } finally {
                this.Cursor = Cursors.Default;
                downloadBtn.Enabled = true;
            }
        }

        // --- Build M4B Action ---
        private void BtnBuildM4b_Click(object sender, EventArgs e) {
            if (listChapters.Items.Count == 0) {
                MessageBox.Show("Please add MP3s to the Audiobook Chapters list.", "No files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DateTime dt = dtpM4bDate.Value;
            string album = txtM4bAlbum.Text.Trim();
            string artist = cmbM4bArtist.Text.Trim();
            string m4bLocation = cmbM4bLocation.Text.Trim();

            if (string.IsNullOrWhiteSpace(album)) {
                MessageBox.Show("Please provide an Album name for the M4B.", "Missing Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Auto-learn artist and location if enabled
            if (_settings.AutoAddArtists && !string.IsNullOrWhiteSpace(artist)) _settings.AddArtistIfNew(artist);
            if (_settings.AutoAddLocations && !string.IsNullOrWhiteSpace(m4bLocation)) _settings.AddLocationIfNew(m4bLocation);
            RefreshDropdownLists();

            string safeM4bArtist = SanitizeFilenameSegment(artist);
            string safeM4bAlbum = SanitizeFilenameSegment(album);
            string safeM4bLocation = SanitizeFilenameSegment(m4bLocation);

            string defaultFileName = _settings.FormatFilename(dt, safeM4bArtist, safeM4bAlbum, safeM4bLocation, "") + ".m4b";

            using (var dialog = new SaveFileDialog { Filter = "Audiobook (*.m4b)|*.m4b", DefaultExt = "m4b", FileName = defaultFileName, InitialDirectory = _selectedFolder }) {
                if (dialog.ShowDialog(this) == DialogResult.OK) {
                    
                    List<string> mp3s = new List<string>();
                    List<string> chapterTitles = new List<string>();
                    foreach (ChapterItem item in listChapters.Items) {
                        mp3s.Add(item.FilePath);
                        chapterTitles.Add(item.ChapterName);
                    }

                    string fullFfmpeg = _settings.GetResolvedFFmpegPath();

                    if (!File.Exists(fullFfmpeg)) {
                        DialogResult dr = MessageBox.Show(
                            "FFmpeg is required to build M4B audiobooks. It was not found at:\n\n" + fullFfmpeg + 
                            "\n\nWould you like to locate ffmpeg.exe now? (If you don't have it, download from ffmpeg.org)",
                            "FFmpeg Required", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (dr == DialogResult.Yes) {
                            using (var ofd = new OpenFileDialog { Filter = "Executable Files (*.exe)|*.exe", Title = "Locate ffmpeg.exe" }) {
                                if (ofd.ShowDialog(this) == DialogResult.OK) {
                                    _settings.FFmpegPath = ofd.FileName;
                                    _settings.Save();
                                    fullFfmpeg = _settings.GetResolvedFFmpegPath();
                                } else {
                                    return;
                                }
                            }
                        } else {
                            return;
                        }
                    }

                    string bitrate = cmbBitrate.SelectedItem != null ? cmbBitrate.SelectedItem.ToString() : "128k";
                    string coverPath = _m4bCoverPath;
                    string outPath = dialog.FileName;

                    using (var progressDialog = new AudiobookProgressDialog(_settings.Theme)) {
                        string result = "";

                        progressDialog.Shown += async (snd, args) => {
                            try {
                                result = await Task.Run(() => {
                                    return M4bBuilder.Build(
                                        mp3s, chapterTitles, outPath, coverPath, album, artist, bitrate, fullFfmpeg,
                                        pct => progressDialog.SetProgress(pct),
                                        msg => progressDialog.SetStatusText(msg),
                                        progressDialog.CancellationToken
                                    );
                                });
                            } catch (Exception ex) {
                                result = "Error: " + ex.Message;
                            } finally {
                                progressDialog.DialogResult = DialogResult.OK;
                                progressDialog.Close();
                            }
                        };

                        progressDialog.ShowDialog(this);

                        if (result == "Success") {
                            MessageBox.Show("M4B Audiobook created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        } else if (result.Contains("cancelled") || progressDialog.CancellationToken.IsCancellationRequested) {
                            MessageBox.Show("Audiobook build was cancelled.", "Build Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        } else {
                            MessageBox.Show(result, "Audiobook Build Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e) {
            base.OnFormClosing(e);
            try {
                if (picTagCoverPreview != null && picTagCoverPreview.Image != null) {
                    picTagCoverPreview.Image.Dispose();
                    picTagCoverPreview.Image = null;
                }
                if (picCoverPreview != null && picCoverPreview.Image != null) {
                    picCoverPreview.Image.Dispose();
                    picCoverPreview.Image = null;
                }
            } catch { }

            if (_settings != null && _settings.FlushTempOnExit) {
                _settings.FlushInstaTempDirectory();
            }
        }
    }
}
