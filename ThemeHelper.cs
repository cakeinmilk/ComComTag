using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ComComTag {
    public static class ThemeHelper {
        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        [DllImport("uxtheme.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        // Dark Theme Palette
        public static readonly Color DarkFormBg = Color.FromArgb(28, 29, 31);
        public static readonly Color DarkPanelBg = Color.FromArgb(36, 37, 40);
        public static readonly Color DarkControlBg = Color.FromArgb(44, 45, 48);
        public static readonly Color DarkControlText = Color.FromArgb(240, 240, 240);
        public static readonly Color DarkSubduedText = Color.FromArgb(160, 165, 170);
        public static readonly Color DarkBorder = Color.FromArgb(65, 68, 72);
        public static readonly Color DarkButton = Color.FromArgb(52, 54, 58);
        public static readonly Color DarkButtonHover = Color.FromArgb(68, 71, 76);
        public static readonly Color DarkAccent = Color.FromArgb(66, 133, 244);
        public static readonly Color DarkAccentHover = Color.FromArgb(85, 145, 245);
        public static readonly Color DarkAccentText = Color.White;
        public static readonly Color DarkBannerBg = Color.FromArgb(30, 42, 56);
        public static readonly Color DarkBannerFore = Color.FromArgb(200, 225, 255);

        // Light Theme Palette
        public static readonly Color LightBannerBg = Color.FromArgb(240, 244, 248);
        public static readonly Color LightBannerFore = Color.FromArgb(50, 60, 70);

        public static bool IsWindowsInDarkMode() {
            try {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")) {
                    if (key != null) {
                        object val = key.GetValue("AppsUseLightTheme");
                        if (val != null && val is int) {
                            return ((int)val) == 0;
                        }
                    }
                }
            } catch { }
            return false;
        }

        public static bool ShouldUseDarkMode(string themeSetting) {
            if (string.Equals(themeSetting, "Dark", StringComparison.OrdinalIgnoreCase)) return true;
            if (string.Equals(themeSetting, "Light", StringComparison.OrdinalIgnoreCase)) return false;
            return IsWindowsInDarkMode();
        }

        public static void ApplyImmersiveDarkMode(Form form, bool dark) {
            if (form == null) return;
            try {
                if (!form.IsHandleCreated) {
                    IntPtr dummy = form.Handle;
                }
                int darkMode = dark ? 1 : 0;
                // 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Windows 11 / Windows 10 20H1+), 19 for older Windows 10
                int hr = DwmSetWindowAttribute(form.Handle, 20, ref darkMode, sizeof(int));
                if (hr != 0) {
                    DwmSetWindowAttribute(form.Handle, 19, ref darkMode, sizeof(int));
                }
            } catch { }
        }

        public static void ApplyTheme(Form form, bool dark) {
            if (form == null) return;

            form.BackColor = dark ? DarkFormBg : SystemColors.Control;
            form.ForeColor = dark ? DarkControlText : SystemColors.ControlText;

            ApplyImmersiveDarkMode(form, dark);

            // Re-apply when form handle is created or displayed to guarantee title bar dark mode
            form.Shown -= Form_ApplyDwm;
            form.Shown += Form_ApplyDwm;
            form.Tag = dark ? "dark" : "light";

            ApplyControlTheme(form, dark);
        }

        private static void Form_ApplyDwm(object sender, EventArgs e) {
            Form form = sender as Form;
            if (form != null) {
                bool isDark = (form.Tag != null && form.Tag.ToString() == "dark");
                ApplyImmersiveDarkMode(form, isDark);
            }
        }

        public static void ApplyControlTheme(Control parent, bool dark) {
            if (parent == null) return;

            foreach (Control c in parent.Controls) {
                // Apply Windows 10/11 native DarkMode theme to win32 handles
                try {
                    if (c.IsHandleCreated) {
                        SetWindowTheme(c.Handle, dark ? "DarkMode_Explorer" : "Explorer", null);
                    }
                } catch { }

                MenuStrip ms = c as MenuStrip;
                if (ms != null) {
                    ms.Renderer = dark ? (ToolStripRenderer)new DarkToolStripRenderer() : (ToolStripRenderer)new ToolStripProfessionalRenderer();
                    ms.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                    ms.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    continue;
                }

                ToolStrip ts = c as ToolStrip;
                if (ts != null) {
                    ts.Renderer = dark ? (ToolStripRenderer)new DarkToolStripRenderer() : (ToolStripRenderer)new ToolStripProfessionalRenderer();
                    ts.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                    ts.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    continue;
                }

                DarkTabControl dtc = c as DarkTabControl;
                if (dtc != null) {
                    dtc.IsDark = dark;
                    foreach (TabPage page in dtc.TabPages) {
                        page.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                        page.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                        ApplyControlTheme(page, dark);
                    }
                    continue;
                }

                TabControl tc = c as TabControl;
                if (tc != null) {
                    tc.DrawMode = TabDrawMode.OwnerDrawFixed;
                    tc.DrawItem -= TabControl_DrawItem;
                    tc.DrawItem += TabControl_DrawItem;
                    foreach (TabPage page in tc.TabPages) {
                        page.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                        page.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                        ApplyControlTheme(page, dark);
                    }
                    continue;
                }

                TabPage tp = c as TabPage;
                if (tp != null) {
                    tp.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                    tp.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    ApplyControlTheme(tp, dark);
                    continue;
                }

                Panel p = c as Panel;
                if (p != null) {
                    if (p.Name == "pnlHelpBanner" || (p.Tag != null && p.Tag.ToString() == "banner")) {
                        p.BackColor = dark ? DarkBannerBg : LightBannerBg;
                    } else {
                        p.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                    }
                    p.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    ApplyControlTheme(p, dark);
                    continue;
                }

                GroupBox gb = c as GroupBox;
                if (gb != null) {
                    gb.BackColor = dark ? DarkPanelBg : SystemColors.Control;
                    gb.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    ApplyControlTheme(gb, dark);
                    continue;
                }

                TextBox txt = c as TextBox;
                if (txt != null) {
                    txt.BackColor = dark ? DarkControlBg : SystemColors.Window;
                    txt.ForeColor = dark ? DarkControlText : SystemColors.WindowText;
                    txt.BorderStyle = BorderStyle.FixedSingle;
                    continue;
                }

                ListBox lb = c as ListBox;
                if (lb != null) {
                    lb.BackColor = dark ? DarkControlBg : SystemColors.Window;
                    lb.ForeColor = dark ? DarkControlText : SystemColors.WindowText;
                    lb.BorderStyle = BorderStyle.FixedSingle;
                    continue;
                }

                ComboBox cmb = c as ComboBox;
                if (cmb != null) {
                    cmb.BackColor = dark ? DarkControlBg : SystemColors.Window;
                    cmb.ForeColor = dark ? DarkControlText : SystemColors.WindowText;
                    cmb.FlatStyle = dark ? FlatStyle.Flat : FlatStyle.Standard;
                    if (dark) {
                        cmb.DrawMode = DrawMode.OwnerDrawFixed;
                        cmb.DrawItem -= ComboBox_DrawItem;
                        cmb.DrawItem += ComboBox_DrawItem;
                    } else {
                        cmb.DrawMode = DrawMode.Normal;
                        cmb.DrawItem -= ComboBox_DrawItem;
                    }
                    continue;
                }

                NumericUpDown num = c as NumericUpDown;
                if (num != null) {
                    num.BackColor = dark ? DarkControlBg : SystemColors.Window;
                    num.ForeColor = dark ? DarkControlText : SystemColors.WindowText;
                    continue;
                }

                Button btn = c as Button;
                if (btn != null) {
                    bool isAccent = btn.Font != null && btn.Font.Bold && (btn.Text.Contains("Save") || btn.Text.Contains("Build") || btn.Text.Contains("OK") || btn.Text.Contains("Select"));
                    btn.FlatStyle = FlatStyle.Flat;
                    if (dark) {
                        if (isAccent) {
                            btn.BackColor = DarkAccent;
                            btn.ForeColor = DarkAccentText;
                            btn.FlatAppearance.BorderColor = DarkAccentHover;
                        } else {
                            btn.BackColor = DarkButton;
                            btn.ForeColor = DarkControlText;
                            btn.FlatAppearance.BorderColor = DarkBorder;
                        }
                    } else {
                        if (isAccent) {
                            btn.BackColor = Color.FromArgb(235, 242, 250);
                            btn.ForeColor = Color.FromArgb(25, 103, 210);
                            btn.FlatAppearance.BorderColor = Color.FromArgb(170, 200, 240);
                        } else {
                            btn.BackColor = SystemColors.Control;
                            btn.ForeColor = SystemColors.ControlText;
                            btn.FlatAppearance.BorderColor = SystemColors.ControlDark;
                        }
                    }
                    continue;
                }

                Label lbl = c as Label;
                if (lbl != null) {
                    if (lbl.Name == "lblHelp" || (lbl.Tag != null && lbl.Tag.ToString() == "banner")) {
                        lbl.ForeColor = dark ? DarkBannerFore : LightBannerFore;
                    } else if (lbl.ForeColor == SystemColors.GrayText || (lbl.Tag != null && lbl.Tag.ToString() == "subdued")) {
                        lbl.ForeColor = dark ? DarkSubduedText : SystemColors.GrayText;
                    } else {
                        lbl.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    }
                    continue;
                }

                CheckBox chk = c as CheckBox;
                if (chk != null) {
                    chk.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    continue;
                }

                RadioButton rb = c as RadioButton;
                if (rb != null) {
                    rb.ForeColor = dark ? DarkControlText : SystemColors.ControlText;
                    continue;
                }

                DateTimePicker dtp = c as DateTimePicker;
                if (dtp != null) {
                    dtp.CalendarMonthBackground = dark ? DarkControlBg : SystemColors.Window;
                    dtp.CalendarForeColor = dark ? DarkControlText : SystemColors.WindowText;
                    dtp.CalendarTitleBackColor = dark ? DarkPanelBg : SystemColors.ActiveCaption;
                    dtp.CalendarTitleForeColor = dark ? DarkControlText : SystemColors.ActiveCaptionText;
                    dtp.CalendarTrailingForeColor = dark ? DarkSubduedText : SystemColors.GrayText;
                    continue;
                }

                PictureBox pic = c as PictureBox;
                if (pic != null) {
                    pic.BackColor = dark ? DarkControlBg : SystemColors.Window;
                    continue;
                }

                if (c.HasChildren) {
                    ApplyControlTheme(c, dark);
                }
            }
        }

        private static void ComboBox_DrawItem(object sender, DrawItemEventArgs e) {
            ComboBox cmb = sender as ComboBox;
            if (cmb == null || e.Index < 0 || e.Index >= cmb.Items.Count) return;

            bool isDark = cmb.BackColor == DarkControlBg || cmb.BackColor == DarkPanelBg;
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            Color backColor = isDark ? (isSelected ? DarkButtonHover : DarkControlBg) : (isSelected ? SystemColors.Highlight : SystemColors.Window);
            Color textColor = isDark ? DarkControlText : (isSelected ? SystemColors.HighlightText : SystemColors.WindowText);

            using (var brush = new SolidBrush(backColor)) {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            string text = cmb.Items[e.Index].ToString();
            TextRenderer.DrawText(
                e.Graphics,
                text,
                cmb.Font,
                e.Bounds,
                textColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter
            );
        }

        private static void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
            TabControl tc = sender as TabControl;
            if (tc == null) return;
            if (e.Index < 0 || e.Index >= tc.TabPages.Count) return;

            TabPage page = tc.TabPages[e.Index];
            bool isDark = page.BackColor == DarkPanelBg || page.BackColor == DarkFormBg;
            bool isSelected = tc.SelectedIndex == e.Index;

            Color tabBack = isDark 
                ? (isSelected ? DarkPanelBg : DarkFormBg) 
                : (isSelected ? SystemColors.Control : SystemColors.ControlLight);
            Color tabText = isDark ? DarkControlText : SystemColors.ControlText;

            using (var brush = new SolidBrush(tabBack)) {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            if (isSelected) {
                using (var pen = new Pen(isDark ? DarkAccent : SystemColors.Highlight, 2)) {
                    e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Top + 1, e.Bounds.Right, e.Bounds.Top + 1);
                }
            }

            TextRenderer.DrawText(
                e.Graphics,
                page.Text,
                tc.Font,
                e.Bounds,
                tabText,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
            );
        }
    }

    public class DarkTabControl : TabControl {
        private bool _isDark;

        public bool IsDark {
            get { return _isDark; }
            set {
                if (_isDark != value) {
                    _isDark = value;
                    SetStyle(ControlStyles.UserPaint, _isDark);
                    Invalidate();
                }
            }
        }

        public DarkTabControl() {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
        }

        protected override void OnPaint(PaintEventArgs e) {
            if (!_isDark) {
                base.OnPaint(e);
                return;
            }

            // Fill header strip background with DarkFormBg
            using (var brush = new SolidBrush(ThemeHelper.DarkFormBg)) {
                e.Graphics.FillRectangle(brush, ClientRectangle);
            }

            // Fill body area where tab pages sit with DarkPanelBg
            if (SelectedTab != null) {
                Rectangle pageRect = SelectedTab.Bounds;
                Rectangle bodyRect = new Rectangle(0, pageRect.Top - 2, Width, Height - pageRect.Top + 2);
                using (var brush = new SolidBrush(ThemeHelper.DarkPanelBg)) {
                    e.Graphics.FillRectangle(brush, bodyRect);
                }
                using (var pen = new Pen(ThemeHelper.DarkBorder)) {
                    e.Graphics.DrawRectangle(pen, 0, pageRect.Top - 2, Width - 1, Height - pageRect.Top + 1);
                }
            }

            // Draw Tabs
            for (int i = 0; i < TabCount; i++) {
                Rectangle tabRect = GetTabRect(i);
                bool isSelected = (SelectedIndex == i);

                Color backColor = isSelected ? ThemeHelper.DarkPanelBg : ThemeHelper.DarkFormBg;
                Color textColor = isSelected ? ThemeHelper.DarkControlText : ThemeHelper.DarkSubduedText;

                using (var brush = new SolidBrush(backColor)) {
                    e.Graphics.FillRectangle(brush, tabRect);
                }

                if (isSelected) {
                    using (var pen = new Pen(ThemeHelper.DarkAccent, 2)) {
                        e.Graphics.DrawLine(pen, tabRect.Left, tabRect.Top, tabRect.Right, tabRect.Top);
                    }
                } else {
                    using (var pen = new Pen(ThemeHelper.DarkBorder, 1)) {
                        e.Graphics.DrawLine(pen, tabRect.Left, tabRect.Bottom, tabRect.Right, tabRect.Bottom);
                    }
                }

                TextRenderer.DrawText(
                    e.Graphics,
                    TabPages[i].Text,
                    Font,
                    tabRect,
                    textColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                );
            }
        }
    }

    public class DarkToolStripRenderer : ToolStripProfessionalRenderer {
        public DarkToolStripRenderer() : base(new DarkColorTable()) { }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e) {
            e.TextColor = ThemeHelper.DarkControlText;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e) {
            e.ArrowColor = ThemeHelper.DarkControlText;
            base.OnRenderArrow(e);
        }
    }

    public class DarkColorTable : ProfessionalColorTable {
        public override Color MenuStripGradientBegin { get { return ThemeHelper.DarkPanelBg; } }
        public override Color MenuStripGradientEnd { get { return ThemeHelper.DarkPanelBg; } }
        public override Color ToolStripDropDownBackground { get { return ThemeHelper.DarkPanelBg; } }
        public override Color ImageMarginGradientBegin { get { return ThemeHelper.DarkPanelBg; } }
        public override Color ImageMarginGradientMiddle { get { return ThemeHelper.DarkPanelBg; } }
        public override Color ImageMarginGradientEnd { get { return ThemeHelper.DarkPanelBg; } }
        public override Color MenuItemSelected { get { return ThemeHelper.DarkButtonHover; } }
        public override Color MenuItemSelectedGradientBegin { get { return ThemeHelper.DarkButtonHover; } }
        public override Color MenuItemSelectedGradientEnd { get { return ThemeHelper.DarkButtonHover; } }
        public override Color MenuItemPressedGradientBegin { get { return ThemeHelper.DarkButton; } }
        public override Color MenuItemPressedGradientMiddle { get { return ThemeHelper.DarkButton; } }
        public override Color MenuItemPressedGradientEnd { get { return ThemeHelper.DarkButton; } }
        public override Color MenuBorder { get { return ThemeHelper.DarkBorder; } }
        public override Color MenuItemBorder { get { return ThemeHelper.DarkBorder; } }
        public override Color SeparatorDark { get { return ThemeHelper.DarkBorder; } }
        public override Color SeparatorLight { get { return ThemeHelper.DarkPanelBg; } }
    }
}
