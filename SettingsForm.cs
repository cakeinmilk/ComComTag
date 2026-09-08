using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;

namespace ComComTag {
    public class SettingsForm : Form {
        private Settings _settings;
        
        private DarkTabControl tabControl;
        private TabPage tabGeneral;
        private TabPage tabAudiobookSettings;
        private TabPage tabConventions;

        // General Tab Controls
        private TextBox txtFfmpeg;
        private Button btnBrowseFFmpeg;
        private TextBox txtInstaTemp;
        private Button btnBrowseInstaTemp;
        private ComboBox cmbTheme;
        private CheckBox chkFlushTemp;
        private CheckBox chkAutoAddArtists;
        private CheckBox chkAutoAddLocations;
        private CheckBox chkUseFirstArtist;
        private CheckBox chkClearErasesAllTags;
        private CheckBox chkGroupAdditionalArtists;
        private Label lblGroupThreshold;
        private NumericUpDown numGroupThreshold;
        private TextBox txtGroupAdditionalArtistsText;
        private TextBox txtArtists;
        private TextBox txtLocations;

        // Audiobook Settings Tab Controls
        private ComboBox cmbDefaultBitrate;
        private Label lblBitrateNote;
        private TextBox txtDefaultM4bArtist;
        private Label lblDefaultM4bArtistNote;
        private CheckBox chkAppendDuration;

        // Conventions Tab Controls
        private ComboBox cmbDateFormat;
        private Label lblDateSample;
        private TextBox txtPattern;
        private Button btnInsertDate;
        private Button btnInsertArtist;
        private Button btnInsertShow;
        private Button btnInsertLocation;
        private Button btnInsertTrack;
        private Button btnInsertDash;
        private Button btnInsertUnderscore;
        private Label lblCaseInfo;

        // Playground Controls
        private DateTimePicker dtpPlayDate;
        private TextBox txtPlayArtist;
        private TextBox txtPlayShow;
        private TextBox txtPlayLocation;
        private Label lblPlayResult;

        // Bottom Dialog Buttons
        private Button btnSave;
        private Button btnCancel;

        public SettingsForm(Settings settings) {
            _settings = settings;
            InitializeComponent();
            ThemeHelper.ApplyTheme(this, ThemeHelper.ShouldUseDarkMode(_settings.Theme));
        }

        private void InitializeComponent() {
            this.Text = "Settings";
            this.Size = new Size(640, 680);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            tabControl = new DarkTabControl {
                Location = new Point(15, 12),
                Size = new Size(595, 575)
            };

            tabGeneral = new TabPage("General");
            tabAudiobookSettings = new TabPage("Audiobook");
            tabConventions = new TabPage("Conventions");

            tabControl.TabPages.Add(tabGeneral);
            tabControl.TabPages.Add(tabAudiobookSettings);
            tabControl.TabPages.Add(tabConventions);
            this.Controls.Add(tabControl);

            InitializeGeneralTab();
            InitializeAudiobookSettingsTab();
            InitializeConventionsTab();

            // Bottom Buttons
            btnSave = new Button { Location = new Point(420, 597), Size = new Size(90, 32), Text = "Save", Font = new Font(this.Font, FontStyle.Bold) };
            btnSave.Click += BtnSave_Click;
            
            btnCancel = new Button { Location = new Point(520, 597), Size = new Size(90, 32), Text = "Cancel" };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;

            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }

        private void InitializeGeneralTab() {
            int yOffset = 15;
            
            // FFmpeg Row
            Label lblFfmpeg = new Label { Location = new Point(15, yOffset), Size = new Size(110, 20), Text = "FFmpeg Path:" };
            txtFfmpeg = new TextBox { Location = new Point(135, yOffset - 2), Size = new Size(345, 22), Text = _settings.FFmpegPath };
            btnBrowseFFmpeg = new Button { Location = new Point(490, yOffset - 4), Size = new Size(80, 25), Text = "Browse..." };
            btnBrowseFFmpeg.Click += BtnBrowseFFmpeg_Click;

            // Instagram Temp Directory Row
            yOffset += 32;
            Label lblInstaTemp = new Label { Location = new Point(15, yOffset), Size = new Size(120, 20), Text = "Instagram Temp Dir:" };
            txtInstaTemp = new TextBox { Location = new Point(135, yOffset - 2), Size = new Size(345, 22), Text = _settings.InstaTempDirectory };
            btnBrowseInstaTemp = new Button { Location = new Point(490, yOffset - 4), Size = new Size(80, 25), Text = "Browse..." };
            btnBrowseInstaTemp.Click += BtnBrowseInstaTemp_Click;

            // Theme Row
            yOffset += 32;
            Label lblTheme = new Label { Location = new Point(15, yOffset), Size = new Size(110, 20), Text = "Theme:" };
            cmbTheme = new ComboBox { Location = new Point(135, yOffset - 2), Size = new Size(160, 22), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbTheme.Items.AddRange(new object[] { "System Default", "Light", "Dark" });
            if (string.Equals(_settings.Theme, "Dark", StringComparison.OrdinalIgnoreCase)) {
                cmbTheme.SelectedIndex = 2;
            } else if (string.Equals(_settings.Theme, "Light", StringComparison.OrdinalIgnoreCase)) {
                cmbTheme.SelectedIndex = 1;
            } else {
                cmbTheme.SelectedIndex = 0;
            }
            cmbTheme.SelectedIndexChanged += (s, e) => {
                string sel = cmbTheme.SelectedItem.ToString();
                string t = (sel == "Dark") ? "Dark" : ((sel == "Light") ? "Light" : "System");
                bool isDark = ThemeHelper.ShouldUseDarkMode(t);
                ThemeHelper.ApplyTheme(this, isDark);
            };

            // Options Checkboxes
            yOffset += 28;
            chkFlushTemp = new CheckBox {
                Location = new Point(15, yOffset),
                Size = new Size(540, 20),
                Text = "Flush temporary Instagram files on exit",
                Checked = _settings.FlushTempOnExit
            };

            yOffset += 22;
            chkAutoAddArtists = new CheckBox {
                Location = new Point(15, yOffset),
                Size = new Size(270, 20),
                Text = "Add new artists to dropdown",
                Checked = _settings.AutoAddArtists
            };

            chkAutoAddLocations = new CheckBox {
                Location = new Point(300, yOffset),
                Size = new Size(270, 20),
                Text = "Add new venues to dropdown",
                Checked = _settings.AutoAddLocations
            };

            yOffset += 22;
            chkUseFirstArtist = new CheckBox {
                Location = new Point(15, yOffset),
                Size = new Size(270, 20),
                Text = "Use first artist as album artist",
                Checked = _settings.UseFirstArtistAsAlbumArtist
            };

            chkClearErasesAllTags = new CheckBox {
                Location = new Point(300, yOffset),
                Size = new Size(270, 20),
                Text = "Clear button erases all ID3 tags",
                Checked = _settings.ClearErasesAllTags
            };

            yOffset += 22;
            chkGroupAdditionalArtists = new CheckBox {
                Location = new Point(15, yOffset),
                AutoSize = true,
                Text = "Group additional artists in filename",
                Checked = _settings.GroupAdditionalArtists
            };
            int chkWidth = TextRenderer.MeasureText(chkGroupAdditionalArtists.Text, chkGroupAdditionalArtists.Font).Width + 22;

            lblGroupThreshold = new Label {
                AutoSize = true,
                Location = new Point(15 + chkWidth + 4, yOffset + 2),
                Text = "if count \u2265"
            };
            int lblWidth = TextRenderer.MeasureText(lblGroupThreshold.Text, lblGroupThreshold.Font).Width;

            numGroupThreshold = new NumericUpDown {
                Location = new Point(15 + chkWidth + 4 + lblWidth + 4, yOffset - 1),
                Size = new Size(38, 22),
                Minimum = 2,
                Maximum = 99,
                Value = Math.Max(2, _settings.GroupAdditionalArtistsThreshold)
            };

            txtGroupAdditionalArtistsText = new TextBox {
                Location = new Point(15 + chkWidth + 4 + lblWidth + 4 + 38 + 6, yOffset - 1),
                Size = new Size(130, 22),
                Text = _settings.GroupAdditionalArtistsText ?? "and friends"
            };

            // Two-Column Layout for Artists and Locations
            yOffset += 24;
            int colWidth = 270;

            Label lblArt = new Label { Location = new Point(15, yOffset), Size = new Size(colWidth, 20), Text = "Artists (one per line):", Font = new Font(this.Font, FontStyle.Bold) };
            Label lblLoc = new Label { Location = new Point(300, yOffset), Size = new Size(colWidth, 20), Text = "Venues (one per line):", Font = new Font(this.Font, FontStyle.Bold) };

            yOffset += 22;
            int boxHeight = 314;

            txtArtists = new TextBox { 
                Location = new Point(15, yOffset), 
                Size = new Size(colWidth, boxHeight), 
                Multiline = true, 
                ScrollBars = ScrollBars.Vertical,
                Text = string.Join(Environment.NewLine, _settings.Artists)
            };

            txtLocations = new TextBox { 
                Location = new Point(300, yOffset), 
                Size = new Size(colWidth, boxHeight), 
                Multiline = true, 
                ScrollBars = ScrollBars.Vertical,
                Text = string.Join(Environment.NewLine, _settings.Locations)
            };

            tabGeneral.Controls.Add(lblFfmpeg);
            tabGeneral.Controls.Add(txtFfmpeg);
            tabGeneral.Controls.Add(btnBrowseFFmpeg);
            tabGeneral.Controls.Add(lblInstaTemp);
            tabGeneral.Controls.Add(txtInstaTemp);
            tabGeneral.Controls.Add(btnBrowseInstaTemp);
            tabGeneral.Controls.Add(lblTheme);
            tabGeneral.Controls.Add(cmbTheme);
            tabGeneral.Controls.Add(chkFlushTemp);
            tabGeneral.Controls.Add(chkAutoAddArtists);
            tabGeneral.Controls.Add(chkAutoAddLocations);
            tabGeneral.Controls.Add(chkUseFirstArtist);
            tabGeneral.Controls.Add(chkClearErasesAllTags);
            tabGeneral.Controls.Add(chkGroupAdditionalArtists);
            tabGeneral.Controls.Add(lblGroupThreshold);
            tabGeneral.Controls.Add(numGroupThreshold);
            tabGeneral.Controls.Add(txtGroupAdditionalArtistsText);
            tabGeneral.Controls.Add(lblArt);
            tabGeneral.Controls.Add(lblLoc);
            tabGeneral.Controls.Add(txtArtists);
            tabGeneral.Controls.Add(txtLocations);
        }

        private void InitializeAudiobookSettingsTab() {
            int yOffset = 20;

            // Default Bitrate Row
            Label lblBitrate = new Label { Location = new Point(15, yOffset), Size = new Size(130, 20), Text = "Default Bitrate:", Font = new Font(this.Font, FontStyle.Bold) };
            cmbDefaultBitrate = new ComboBox { Location = new Point(150, yOffset - 2), Size = new Size(85, 22), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDefaultBitrate.Items.AddRange(new object[] { "64k", "96k", "128k", "192k", "256k", "320k" });
            if (cmbDefaultBitrate.Items.Contains(_settings.DefaultBitrate)) {
                cmbDefaultBitrate.SelectedIndex = cmbDefaultBitrate.Items.IndexOf(_settings.DefaultBitrate);
            } else {
                cmbDefaultBitrate.SelectedIndex = 2; // Default 128k
            }
            lblBitrateNote = new Label { 
                Location = new Point(245, yOffset), 
                Size = new Size(320, 20), 
                ForeColor = SystemColors.GrayText, 
                Text = "(Audiobook AAC encoder bitrate)" 
            };

            // Default Audiobook Artist Row
            yOffset += 40;
            Label lblDefaultM4bArtist = new Label { Location = new Point(15, yOffset), Size = new Size(130, 20), Text = "Default Artist:", Font = new Font(this.Font, FontStyle.Bold) };
            txtDefaultM4bArtist = new TextBox { 
                Location = new Point(150, yOffset - 2), 
                Size = new Size(240, 22), 
                Text = _settings.DefaultM4bArtist ?? "Various Artists" 
            };
            lblDefaultM4bArtistNote = new Label { 
                Location = new Point(150, yOffset + 24), 
                Size = new Size(400, 20), 
                ForeColor = SystemColors.GrayText, 
                Text = "(Initial artist populated when creating a new audiobook)" 
            };

            // Chapter Options
            yOffset += 55;
            chkAppendDuration = new CheckBox {
                Location = new Point(15, yOffset),
                Size = new Size(350, 20),
                Text = "Append duration to chapter name",
                Checked = _settings.AppendChapterDuration
            };

            tabAudiobookSettings.Controls.Add(lblBitrate);
            tabAudiobookSettings.Controls.Add(cmbDefaultBitrate);
            tabAudiobookSettings.Controls.Add(lblBitrateNote);
            tabAudiobookSettings.Controls.Add(lblDefaultM4bArtist);
            tabAudiobookSettings.Controls.Add(txtDefaultM4bArtist);
            tabAudiobookSettings.Controls.Add(lblDefaultM4bArtistNote);
            tabAudiobookSettings.Controls.Add(chkAppendDuration);
        }

        private void InitializeConventionsTab() {
            int yOffset = 15;

            // Date Format Section
            Label lblDateTitle = new Label { Location = new Point(15, yOffset), Size = new Size(110, 20), Text = "Date Format:", Font = new Font(this.Font, FontStyle.Bold) };
            cmbDateFormat = new ComboBox { 
                Location = new Point(135, yOffset - 2), 
                Size = new Size(160, 22), 
                DropDownStyle = ComboBoxStyle.DropDownList 
            };
            cmbDateFormat.Items.AddRange(new object[] {
                "yyyy-MM-dd",
                "dd-MM-yyyy",
                "yyyy.MM.dd",
                "dd.MM.yyyy",
                "yyyy_MM_dd",
                "dd_MM_yyyy",
                "dd-MMM-yyyy",
                "dd-MMM-yy",
                "yyyyMMdd"
            });

            if (cmbDateFormat.Items.Contains(_settings.DateFormat)) {
                cmbDateFormat.SelectedItem = _settings.DateFormat;
            } else {
                cmbDateFormat.SelectedIndex = 0;
            }

            lblDateSample = new Label { 
                Location = new Point(305, yOffset), 
                Size = new Size(265, 20), 
                ForeColor = Color.DarkSlateBlue,
                Font = new Font(this.Font, FontStyle.Italic),
                Text = "Example: " + DateTime.Today.ToString(cmbDateFormat.SelectedItem.ToString()) 
            };
            cmbDateFormat.SelectedIndexChanged += (s, e) => {
                lblDateSample.Text = "Example: " + DateTime.Today.ToString(cmbDateFormat.SelectedItem.ToString());
                UpdatePlaygroundPreview();
            };

            // Filename Pattern Section
            yOffset += 38;
            Label lblPatternTitle = new Label { Location = new Point(15, yOffset), Size = new Size(130, 20), Text = "Naming Pattern:", Font = new Font(this.Font, FontStyle.Bold) };
            
            yOffset += 22;
            txtPattern = new TextBox { 
                Location = new Point(15, yOffset), 
                Size = new Size(550, 24), 
                Text = !string.IsNullOrWhiteSpace(_settings.FilenamePattern) ? _settings.FilenamePattern : "{Date} - {Artist} - {Show} - {Location}",
                Font = new Font("Consolas", 10F)
            };
            txtPattern.TextChanged += (s, e) => UpdatePlaygroundPreview();

            // Insert Buttons
            yOffset += 30;
            int btnX = 15;
            btnInsertDate = new Button { Location = new Point(btnX, yOffset), Size = new Size(72, 25), Text = "{Date}" };
            btnInsertDate.Click += (s, e) => InsertPlaceholder("{Date}");
            btnX += 78;

            btnInsertArtist = new Button { Location = new Point(btnX, yOffset), Size = new Size(76, 25), Text = "{Artist}" };
            btnInsertArtist.Click += (s, e) => InsertPlaceholder("{Artist}");
            btnX += 82;

            btnInsertShow = new Button { Location = new Point(btnX, yOffset), Size = new Size(72, 25), Text = "{Show}" };
            btnInsertShow.Click += (s, e) => InsertPlaceholder("{Show}");
            btnX += 78;

            btnInsertLocation = new Button { Location = new Point(btnX, yOffset), Size = new Size(86, 25), Text = "{Location}" };
            btnInsertLocation.Click += (s, e) => InsertPlaceholder("{Location}");
            btnX += 92;

            btnInsertTrack = new Button { Location = new Point(btnX, yOffset), Size = new Size(72, 25), Text = "({Track})" };
            btnInsertTrack.Click += (s, e) => InsertPlaceholder("({Track})");
            btnX += 78;

            btnInsertDash = new Button { Location = new Point(btnX, yOffset), Size = new Size(45, 25), Text = "\" - \"" };
            btnInsertDash.Click += (s, e) => InsertPlaceholder(" - ");
            btnX += 49;

            btnInsertUnderscore = new Button { Location = new Point(btnX, yOffset), Size = new Size(45, 25), Text = "\" _ \"" };
            btnInsertUnderscore.Click += (s, e) => InsertPlaceholder(" _ ");

            yOffset += 30;
            lblCaseInfo = new Label { 
                Location = new Point(15, yOffset), 
                Size = new Size(550, 32), 
                ForeColor = SystemColors.GrayText,
                Text = "Case Options: Use {ARTIST} for ALL CAPS, {artist} for lowercase, or {Artist} for original tag. Optional missing fields (e.g. Show) are cleanly omitted."
            };

            // Interactive Playground Box (Omit Track, Default to Wally Bazoom)
            yOffset += 36;
            GroupBox grpPlay = new GroupBox { 
                Location = new Point(15, yOffset), 
                Size = new Size(550, 205), 
                Text = "Experiment Playground (Live Real-Time Preview)" 
            };

            int py = 24;
            Label pLblDate = new Label { Location = new Point(15, py), Size = new Size(45, 20), Text = "Date:" };
            dtpPlayDate = new DateTimePicker { Location = new Point(65, py - 2), Size = new Size(125, 22), Format = DateTimePickerFormat.Short };
            DateTime parsedPDate;
            if (DateTime.TryParse(_settings.PlaygroundDate, out parsedPDate)) {
                dtpPlayDate.Value = parsedPDate;
            } else {
                dtpPlayDate.Value = new DateTime(1999, 6, 13);
            }
            dtpPlayDate.ValueChanged += (s, e) => UpdatePlaygroundPreview();

            Label pLblArt = new Label { Location = new Point(205, py), Size = new Size(45, 20), Text = "Artist:" };
            txtPlayArtist = new TextBox { 
                Location = new Point(255, py - 2), 
                Size = new Size(275, 22), 
                Text = !string.IsNullOrWhiteSpace(_settings.PlaygroundArtist) ? _settings.PlaygroundArtist : "Wally Bazoom" 
            };
            txtPlayArtist.TextChanged += (s, e) => UpdatePlaygroundPreview();

            py += 32;
            Label pLblShow = new Label { Location = new Point(15, py), Size = new Size(45, 20), Text = "Show:" };
            txtPlayShow = new TextBox { 
                Location = new Point(65, py - 2), 
                Size = new Size(180, 22), 
                Text = !string.IsNullOrWhiteSpace(_settings.PlaygroundShow) ? _settings.PlaygroundShow : "Smile Time" 
            };
            txtPlayShow.TextChanged += (s, e) => UpdatePlaygroundPreview();

            Label pLblLoc = new Label { Location = new Point(255, py), Size = new Size(50, 20), Text = "Venue:" };
            txtPlayLocation = new TextBox { 
                Location = new Point(310, py - 2), 
                Size = new Size(220, 22), 
                Text = !string.IsNullOrWhiteSpace(_settings.PlaygroundLocation) ? _settings.PlaygroundLocation : "The Aigburth Arms, Liverpool" 
            };
            txtPlayLocation.TextChanged += (s, e) => UpdatePlaygroundPreview();

            py += 38;
            Label pLblOut = new Label { Location = new Point(15, py), Size = new Size(110, 20), Text = "Sample Output:", Font = new Font(this.Font, FontStyle.Bold) };
            lblPlayResult = new Label { 
                Location = new Point(130, py), 
                Size = new Size(405, 55), 
                ForeColor = Color.DarkGreen, 
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Text = ""
            };

            grpPlay.Controls.Add(pLblDate);
            grpPlay.Controls.Add(dtpPlayDate);
            grpPlay.Controls.Add(pLblArt);
            grpPlay.Controls.Add(txtPlayArtist);
            grpPlay.Controls.Add(pLblShow);
            grpPlay.Controls.Add(txtPlayShow);
            grpPlay.Controls.Add(pLblLoc);
            grpPlay.Controls.Add(txtPlayLocation);
            grpPlay.Controls.Add(pLblOut);
            grpPlay.Controls.Add(lblPlayResult);

            tabConventions.Controls.Add(lblDateTitle);
            tabConventions.Controls.Add(cmbDateFormat);
            tabConventions.Controls.Add(lblDateSample);
            tabConventions.Controls.Add(lblPatternTitle);
            tabConventions.Controls.Add(txtPattern);
            tabConventions.Controls.Add(btnInsertDate);
            tabConventions.Controls.Add(btnInsertArtist);
            tabConventions.Controls.Add(btnInsertShow);
            tabConventions.Controls.Add(btnInsertLocation);
            tabConventions.Controls.Add(btnInsertTrack);
            tabConventions.Controls.Add(btnInsertDash);
            tabConventions.Controls.Add(btnInsertUnderscore);
            tabConventions.Controls.Add(lblCaseInfo);
            tabConventions.Controls.Add(grpPlay);

            UpdatePlaygroundPreview();
        }

        private void InsertPlaceholder(string placeholder) {
            int selStart = txtPattern.SelectionStart;
            txtPattern.Text = txtPattern.Text.Insert(selStart, placeholder);
            txtPattern.SelectionStart = selStart + placeholder.Length;
            txtPattern.Focus();
        }

        private void UpdatePlaygroundPreview() {
            if (lblPlayResult == null || txtPattern == null || cmbDateFormat == null) return;
            
            var dummySettings = new Settings {
                DateFormat = cmbDateFormat.SelectedItem != null ? cmbDateFormat.SelectedItem.ToString() : "yyyy-MM-dd",
                FilenamePattern = txtPattern.Text.Trim()
            };

            string formatted = dummySettings.FormatFilename(
                dtpPlayDate.Value, 
                txtPlayArtist.Text.Trim(), 
                txtPlayShow.Text.Trim(), 
                txtPlayLocation.Text.Trim(), 
                ""
            );

            lblPlayResult.Text = formatted + ".mp3";
        }

        private void BtnBrowseFFmpeg_Click(object sender, EventArgs e) {
            using (var ofd = new OpenFileDialog { Filter = "Executable Files (*.exe)|*.exe", Title = "Locate ffmpeg.exe" }) {
                if (ofd.ShowDialog(this) == DialogResult.OK) {
                    txtFfmpeg.Text = ofd.FileName;
                }
            }
        }

        private void BtnBrowseInstaTemp_Click(object sender, EventArgs e) {
            using (var fbd = new FolderBrowserDialog { Description = "Select folder for temporary Instagram downloads" }) {
                if (!string.IsNullOrEmpty(txtInstaTemp.Text) && Directory.Exists(txtInstaTemp.Text)) {
                    fbd.SelectedPath = txtInstaTemp.Text;
                }
                if (fbd.ShowDialog(this) == DialogResult.OK) {
                    txtInstaTemp.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e) {
            _settings.FFmpegPath = txtFfmpeg.Text.Trim();
            if (cmbDefaultBitrate.SelectedItem != null) {
                _settings.DefaultBitrate = cmbDefaultBitrate.SelectedItem.ToString();
            }
            _settings.InstaTempDirectory = txtInstaTemp.Text.Trim();
            _settings.FlushTempOnExit = chkFlushTemp.Checked;
            _settings.AutoAddArtists = chkAutoAddArtists.Checked;
            _settings.AutoAddLocations = chkAutoAddLocations.Checked;
            _settings.UseFirstArtistAsAlbumArtist = chkUseFirstArtist.Checked;
            _settings.ClearErasesAllTags = chkClearErasesAllTags.Checked;
            if (cmbTheme.SelectedItem != null) {
                string sel = cmbTheme.SelectedItem.ToString();
                _settings.Theme = (sel == "Dark") ? "Dark" : ((sel == "Light") ? "Light" : "System");
            }
            _settings.GroupAdditionalArtists = chkGroupAdditionalArtists.Checked;
            _settings.GroupAdditionalArtistsThreshold = (int)numGroupThreshold.Value;
            _settings.GroupAdditionalArtistsText = string.IsNullOrWhiteSpace(txtGroupAdditionalArtistsText.Text) ? "and friends" : txtGroupAdditionalArtistsText.Text.Trim();
            _settings.DefaultM4bArtist = string.IsNullOrWhiteSpace(txtDefaultM4bArtist.Text) ? "Various Artists" : txtDefaultM4bArtist.Text.Trim();
            _settings.AppendChapterDuration = chkAppendDuration.Checked;
            
            // Save Conventions
            if (cmbDateFormat.SelectedItem != null) {
                _settings.DateFormat = cmbDateFormat.SelectedItem.ToString();
            }
            if (!string.IsNullOrWhiteSpace(txtPattern.Text)) {
                _settings.FilenamePattern = txtPattern.Text.Trim();
            }

            // Save Playground User Values
            _settings.PlaygroundDate = dtpPlayDate.Value.ToString("yyyy-MM-dd");
            _settings.PlaygroundArtist = txtPlayArtist.Text.Trim();
            _settings.PlaygroundShow = txtPlayShow.Text.Trim();
            _settings.PlaygroundLocation = txtPlayLocation.Text.Trim();

            // Save Artists
            _settings.Artists.Clear();
            var artLines = txtArtists.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in artLines) {
                if (!string.IsNullOrWhiteSpace(line)) {
                    _settings.Artists.Add(line.Trim());
                }
            }
            if (_settings.Artists.Count == 0) _settings.Artists.Add("Various Artists");

            // Save Locations
            _settings.Locations.Clear();
            var locLines = txtLocations.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in locLines) {
                if (!string.IsNullOrWhiteSpace(line)) {
                    _settings.Locations.Add(line.Trim());
                }
            }
            if (_settings.Locations.Count == 0) _settings.Locations.Add("Default Location");
            
            _settings.Save();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
