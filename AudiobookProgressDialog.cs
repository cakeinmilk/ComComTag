using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Reflection;

namespace ComComTag {
    public class AudiobookProgressDialog : Form {
        private Label lblTitle;
        private Label lblStatus;
        private ProgressBar progressBar;
        private Button btnCancel;
        private CancellationTokenSource _cts;
        private bool _isCancelled = false;

        public CancellationToken CancellationToken {
            get { return _cts != null ? _cts.Token : CancellationToken.None; }
        }

        public AudiobookProgressDialog(string theme = "System") {
            _cts = new CancellationTokenSource();
            InitializeComponent();
            ThemeHelper.ApplyTheme(this, ThemeHelper.ShouldUseDarkMode(theme));
        }

        private void InitializeComponent() {
            this.Text = "Building Audiobook...";
            this.ClientSize = new Size(440, 160);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;

            // Set Application Icon if embedded
            try {
                var assembly = Assembly.GetExecutingAssembly();
                using (var iconStream = assembly.GetManifestResourceStream("ComComTag.icon.ico")) {
                    if (iconStream != null) {
                        this.Icon = new Icon(iconStream);
                    }
                }
            } catch { }

            lblTitle = new Label {
                Location = new Point(20, 16),
                Size = new Size(400, 22),
                Font = new Font(this.Font, FontStyle.Bold),
                Text = "Creating M4B Audiobook..."
            };

            lblStatus = new Label {
                Location = new Point(20, 38),
                Size = new Size(400, 20),
                ForeColor = SystemColors.GrayText,
                Text = "Starting build..."
            };

            progressBar = new ProgressBar {
                Location = new Point(20, 64),
                Size = new Size(400, 22),
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            btnCancel = new Button {
                Location = new Point(165, 105),
                Size = new Size(110, 32),
                Text = "Cancel Build",
                Font = new Font(this.Font, FontStyle.Regular),
                DialogResult = DialogResult.None
            };
            btnCancel.Click += (s, e) => RequestCancel();

            this.Controls.Add(lblTitle);
            this.Controls.Add(lblStatus);
            this.Controls.Add(progressBar);
            this.Controls.Add(btnCancel);

            this.FormClosing += (s, e) => {
                if (!_isCancelled && this.DialogResult == DialogResult.None) {
                    RequestCancel();
                    e.Cancel = true;
                }
            };
        }

        public void SetProgress(int percent) {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            this.BeginInvoke((Action)(() => {
                progressBar.Value = Math.Max(0, Math.Min(100, percent));
                if (percent > 0 && percent < 100) {
                    lblStatus.Text = string.Format("Transcoding audio and muxing chapters... {0}%", percent);
                } else if (percent >= 100) {
                    lblStatus.Text = "Finalizing metadata and cover art...";
                }
            }));
        }

        public void SetStatusText(string message) {
            if (this.IsDisposed || !this.IsHandleCreated) return;
            this.BeginInvoke((Action)(() => {
                lblStatus.Text = message;
            }));
        }

        private void RequestCancel() {
            if (_isCancelled) return;
            _isCancelled = true;
            lblStatus.Text = "Cancelling build...";
            btnCancel.Enabled = false;
            try {
                if (_cts != null && !_cts.IsCancellationRequested) {
                    _cts.Cancel();
                }
            } catch { }
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                if (_cts != null) {
                    _cts.Dispose();
                    _cts = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
