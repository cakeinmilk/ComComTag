using System;
using System.Drawing;
using System.Windows.Forms;

namespace ComComTag {
    public class InputPromptDialog : Form {
        private TextBox txtInput;
        private Button btnOk;
        private Button btnCancel;
        private Label lblPrompt;

        public string InputText {
            get { return txtInput.Text; }
            set { txtInput.Text = value; }
        }

        public InputPromptDialog(string title, string prompt, string defaultText = "", string themeSetting = "System") {
            InitializeComponent(title, prompt, defaultText);
            ThemeHelper.ApplyTheme(this, ThemeHelper.ShouldUseDarkMode(themeSetting));
        }

        private void InitializeComponent(string title, string prompt, string defaultText) {
            this.Text = title;
            this.Size = new Size(460, 180);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;

            lblPrompt = new Label {
                Location = new Point(20, 15),
                Size = new Size(405, 35),
                Text = prompt
            };

            txtInput = new TextBox {
                Location = new Point(20, 55),
                Size = new Size(405, 22),
                Text = defaultText
            };

            btnOk = new Button {
                Location = new Point(255, 95),
                Size = new Size(80, 28),
                Text = "OK",
                DialogResult = DialogResult.OK
            };

            btnCancel = new Button {
                Location = new Point(345, 95),
                Size = new Size(80, 28),
                Text = "Cancel",
                DialogResult = DialogResult.Cancel
            };

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtInput);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
        }

        public static string Prompt(IWin32Window owner, string title, string prompt, string defaultText = "", string themeSetting = "System") {
            return ShowPrompt(owner, title, prompt, defaultText, themeSetting);
        }

        public static string ShowPrompt(IWin32Window owner, string title, string prompt, string defaultText = "", string themeSetting = "System") {
            using (var dlg = new InputPromptDialog(title, prompt, defaultText, themeSetting)) {
                if (dlg.ShowDialog(owner) == DialogResult.OK) {
                    return dlg.InputText;
                }
            }
            return null;
        }
    }
}
