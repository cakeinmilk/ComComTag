using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;

namespace ComComTag {
    public class ImageCarouselDialog : Form {
        private List<string> _imagePaths;
        private int _currentIndex = 0;

        private PictureBox picPreview;
        private Label lblCounter;
        private Label lblInfo;
        private Button btnPrev;
        private Button btnNext;
        private Button btnSelect;
        private Button btnCancel;

        public string SelectedImagePath {
            get {
                if (_imagePaths != null && _currentIndex >= 0 && _currentIndex < _imagePaths.Count) {
                    return _imagePaths[_currentIndex];
                }
                return null;
            }
        }

        public int SelectedIndex {
            get { return _currentIndex; }
        }

        public ImageCarouselDialog(List<string> imagePaths, int initialIndex = 0) {
            _imagePaths = imagePaths ?? new List<string>();
            _currentIndex = Math.Max(0, Math.Min(initialIndex, _imagePaths.Count - 1));
            InitializeComponent();
            UpdateImageDisplay();
        }

        private void InitializeComponent() {
            this.Text = "Instagram Artwork Picker - Choose Image";
            this.Size = new Size(580, 620);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.KeyPreview = true;

            picPreview = new PictureBox {
                Location = new Point(25, 20),
                Size = new Size(515, 430),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.Black
            };

            btnPrev = new Button {
                Location = new Point(25, 465),
                Size = new Size(80, 32),
                Text = "< Previous",
                Font = new Font(this.Font, FontStyle.Bold)
            };
            btnPrev.Click += (s, e) => Navigate(-1);

            lblCounter = new Label {
                Location = new Point(115, 472),
                Size = new Size(335, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(this.Font, FontStyle.Bold),
                Text = "Image 0 of 0"
            };

            btnNext = new Button {
                Location = new Point(460, 465),
                Size = new Size(80, 32),
                Text = "Next >",
                Font = new Font(this.Font, FontStyle.Bold)
            };
            btnNext.Click += (s, e) => Navigate(1);

            lblInfo = new Label {
                Location = new Point(25, 505),
                Size = new Size(515, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = SystemColors.GrayText,
                Text = ""
            };

            btnSelect = new Button {
                Location = new Point(280, 535),
                Size = new Size(160, 36),
                Text = "Select This Cover Art",
                Font = new Font(this.Font, FontStyle.Bold),
                DialogResult = DialogResult.OK
            };

            btnCancel = new Button {
                Location = new Point(450, 535),
                Size = new Size(90, 36),
                Text = "Cancel",
                DialogResult = DialogResult.Cancel
            };

            this.AcceptButton = btnSelect;
            this.CancelButton = btnCancel;

            this.KeyDown += ImageCarouselDialog_KeyDown;

            this.Controls.Add(picPreview);
            this.Controls.Add(btnPrev);
            this.Controls.Add(lblCounter);
            this.Controls.Add(btnNext);
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnSelect);
            this.Controls.Add(btnCancel);
        }

        private void ImageCarouselDialog_KeyDown(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Left) {
                Navigate(-1);
                e.Handled = true;
            } else if (e.KeyCode == Keys.Right) {
                Navigate(1);
                e.Handled = true;
            }
        }

        private void Navigate(int delta) {
            if (_imagePaths.Count <= 1) return;
            _currentIndex = (_currentIndex + delta + _imagePaths.Count) % _imagePaths.Count;
            UpdateImageDisplay();
        }

        private void UpdateImageDisplay() {
            if (_imagePaths.Count == 0) {
                picPreview.Image = null;
                lblCounter.Text = "No images available";
                lblInfo.Text = "";
                btnPrev.Enabled = false;
                btnNext.Enabled = false;
                btnSelect.Enabled = false;
                return;
            }

            btnPrev.Enabled = _imagePaths.Count > 1;
            btnNext.Enabled = _imagePaths.Count > 1;
            btnSelect.Enabled = true;

            lblCounter.Text = string.Format("Image {0} of {1}", _currentIndex + 1, _imagePaths.Count);

            string currentPath = _imagePaths[_currentIndex];
            if (File.Exists(currentPath)) {
                try {
                    // Load image into memory to avoid file locks
                    using (var original = Image.FromFile(currentPath)) {
                        if (picPreview.Image != null) picPreview.Image.Dispose();
                        picPreview.Image = new Bitmap(original);
                        lblInfo.Text = string.Format("{0} ({1} x {2})", Path.GetFileName(currentPath), original.Width, original.Height);
                    }
                } catch {
                    picPreview.Image = null;
                    lblInfo.Text = Path.GetFileName(currentPath);
                }
            } else {
                picPreview.Image = null;
                lblInfo.Text = "File not found: " + currentPath;
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e) {
            if (picPreview.Image != null) {
                picPreview.Image.Dispose();
                picPreview.Image = null;
            }
            base.OnFormClosed(e);
        }
    }
}
