using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;

namespace ComComTag {
    public class SettingsForm : Form {
        private Settings _settings;
        
        private TextBox txtFfmpeg;
        private Button btnBrowseFFmpeg;
        private ComboBox cmbDefaultBitrate;
        private TextBox txtLocations;
        private Button btnSave;
        private Button btnCancel;

        public SettingsForm(Settings settings) {
            _settings = settings;
            InitializeComponent();
        }

        private void InitializeComponent() {
            this.Text = "Settings";
            this.Size = new Size(400, 350);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int yOffset = 20;
            
            Label lblFfmpeg = new Label { Location = new Point(15, yOffset), Size = new Size(100, 20), Text = "FFmpeg Path:" };
            txtFfmpeg = new TextBox { Location = new Point(15, yOffset + 20), Size = new Size(280, 20), Text = _settings.FFmpegPath };
            btnBrowseFFmpeg = new Button { Location = new Point(300, yOffset + 18), Size = new Size(70, 24), Text = "Browse" };
            btnBrowseFFmpeg.Click += BtnBrowseFFmpeg_Click;

            yOffset += 40;
            Label lblBitrate = new Label { Location = new Point(15, yOffset), Size = new Size(100, 20), Text = "Default Bitrate:" };
            cmbDefaultBitrate = new ComboBox { Location = new Point(115, yOffset - 2), Size = new Size(180, 20), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDefaultBitrate.Items.AddRange(new object[] { "64k", "96k", "128k", "192k", "256k", "320k" });
            if (cmbDefaultBitrate.Items.Contains(_settings.DefaultBitrate)) {
                cmbDefaultBitrate.SelectedIndex = cmbDefaultBitrate.Items.IndexOf(_settings.DefaultBitrate);
            } else {
                cmbDefaultBitrate.SelectedIndex = 2; // Default 128k
            }

            yOffset += 40;
            Label lblLoc = new Label { Location = new Point(15, yOffset), Size = new Size(300, 20), Text = "Common Locations (one per line):" };
            txtLocations = new TextBox { Location = new Point(15, yOffset + 20), Size = new Size(355, 120), Multiline = true, ScrollBars = ScrollBars.Vertical };
            txtLocations.Text = string.Join(Environment.NewLine, _settings.Locations);

            yOffset += 160;
            btnSave = new Button { Location = new Point(210, yOffset), Size = new Size(75, 30), Text = "Save" };
            btnSave.Click += BtnSave_Click;
            
            btnCancel = new Button { Location = new Point(295, yOffset), Size = new Size(75, 30), Text = "Cancel" };
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            this.Controls.Add(lblFfmpeg);
            this.Controls.Add(txtFfmpeg);
            this.Controls.Add(btnBrowseFFmpeg);
            this.Controls.Add(lblBitrate);
            this.Controls.Add(cmbDefaultBitrate);
            this.Controls.Add(lblLoc);
            this.Controls.Add(txtLocations);
            this.Controls.Add(btnSave);
            this.Controls.Add(btnCancel);
        }

        private void BtnBrowseFFmpeg_Click(object sender, EventArgs e) {
            using (var ofd = new OpenFileDialog { Filter = "Executable Files (*.exe)|*.exe", Title = "Locate ffmpeg.exe" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    txtFfmpeg.Text = ofd.FileName;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e) {
            _settings.FFmpegPath = txtFfmpeg.Text.Trim();
            if (cmbDefaultBitrate.SelectedItem != null) {
                _settings.DefaultBitrate = cmbDefaultBitrate.SelectedItem.ToString();
            }
            
            _settings.Locations.Clear();
            var lines = txtLocations.Text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines) {
                if (!string.IsNullOrWhiteSpace(line)) {
                    _settings.Locations.Add(line.Trim());
                }
            }
            
            _settings.Save();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
