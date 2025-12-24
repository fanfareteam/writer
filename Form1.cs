using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Globalization;

namespace ProjectWriter // Renamed for GitHub release
{
    public partial class Form1 : Form
    {
        // --- UI Groups ---
        private Panel welcomePanel;
        private Panel editorPanel;

        // --- Controls ---
        private MenuStrip menuStrip;
        private ToolStrip toolStrip;
        private RichTextBox richTextBox;
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private ToolStripStatusLabel pageLabel; // New in 0.6

        // --- Toolbar Items ---
        private ToolStripComboBox cmbFontName;
        private ToolStripComboBox cmbFontSize;
        private ToolStripButton btnBold, btnItalic, btnHeading;
        private ToolStripLabel zoomLabel;
        private ToolStripMenuItem recentFilesMenu;

        // --- Logic & Config ---
        private List<string> recentFiles = new List<string>();
        private string currentLang = "en"; // New in 0.6
        private const string configPath = "project.dat"; // Consolidated config
        private bool isUpdatingFont = false;

        // --- Aero API ---
        [DllImport("dwmapi.dll")]
        private static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMarInset);
        [DllImport("dwmapi.dll")]
        private static extern int DwmIsCompositionEnabled(out bool enabled);
        [StructLayout(LayoutKind.Sequential)]
        private struct MARGINS { public int cxLeftWidth, cxRightWidth, cyTopHeight, cyBottomHeight; }

        public Form1(string[] args)
        {
            this.Text = "Project Writer 0.6.1 \"Glasswave\"";
            this.Size = new Size(1000, 850);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = SystemColors.Control;

            LoadSettings(); // Load language/history first
            InitializeEditorUI();
            InitializeWelcomeUI();

            // Detect OS and Set Initial Theme
            DetectInitialTheme();

            if (args != null && args.Length > 0 && File.Exists(args[0]))
            {
                ShowEditor();
                OpenFile(args[0]);
            }
            else
            {
                ShowWelcome();
            }

            this.Load += (s, e) => { if (IsGlassSupported() && this.BackColor == Color.Black) ApplyAero(); };
        }

        private void DetectInitialTheme()
        {
            double osVer = Environment.OSVersion.Version.Major + (Environment.OSVersion.Version.Minor / 10.0);
            if (osVer == 5.1) SetTheme("Luna");
            else if (osVer >= 10.0) SetTheme(IsDarkMode() ? "UWP Dark" : "UWP");
            else if (osVer >= 6.0) SetTheme("Aero");
            else SetTheme("Classic (2009)");
        }

        private bool IsDarkMode()
        {
            return SystemColors.Window.GetBrightness() < 0.5f;
        }

        private void InitializeWelcomeUI()
        {
            welcomePanel = new Panel { Dock = DockStyle.Fill };
            welcomePanel.Paint += (s, e) =>
            {
                // Dynamic header based on language/theme
                Rectangle hRect = new Rectangle(0, 0, welcomePanel.Width, 100);
                Color c1 = Color.FromArgb(0, 50, 120);
                Color c2 = Color.FromArgb(0, 80, 180);

                using (LinearGradientBrush b = new LinearGradientBrush(hRect, c1, c2, 90f))
                    e.Graphics.FillRectangle(b, hRect);

                Rectangle bRect = new Rectangle(0, 100, welcomePanel.Width, welcomePanel.Height - 100);
                using (LinearGradientBrush b = new LinearGradientBrush(bRect, Color.White, Color.FromArgb(240, 240, 240), 90f))
                    e.Graphics.FillRectangle(b, bRect);

                using (Pen p = new Pen(Color.FromArgb(255, 180, 60), 2)) e.Graphics.DrawLine(p, 0, 100, welcomePanel.Width, 100);

                string title = currentLang == "de" ? "Willkommen bei Project Writer" : "Welcome to Project Writer";
                e.Graphics.DrawString(title, new Font("Segoe UI", 24, FontStyle.Bold), Brushes.White, 30, 25);
                e.Graphics.DrawString("Version 0.6.1 \"Glasswave\"", new Font("Segoe UI", 10), Brushes.LightBlue, 35, 70);
            };

            int startY = 150;
            welcomePanel.Controls.Add(CreateWelcomeButton(currentLang == "de" ? "Neues Dokument" : "Create New Document", 50, startY, (s, e) => { ShowEditor(); richTextBox.Clear(); }));
            welcomePanel.Controls.Add(CreateWelcomeButton(currentLang == "de" ? "Öffnen..." : "Open Existing File...", 50, startY + 50, (s, e) => OnOpenClick(null, null)));
            welcomePanel.Controls.Add(CreateWelcomeButton(currentLang == "de" ? "Themen..." : "Configure Theme...", 50, startY + 100, (s, e) =>
            {
                ShowEditor();
                if (menuStrip.Items.ContainsKey("More")) menuStrip.Items["More"].PerformClick();
            }));

            welcomePanel.Visible = false;
            this.Controls.Add(welcomePanel);
        }

        private Button CreateWelcomeButton(string text, int x, int y, EventHandler onClick)
        {
            Button btn = new Button { Text = "  " + text, Location = new Point(x, y), Size = new Size(300, 40), FlatStyle = FlatStyle.Flat, BackColor = Color.Transparent, TextAlign = ContentAlignment.MiddleLeft, Font = new Font("Segoe UI", 12), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            btn.Click += onClick;
            btn.MouseEnter += (s, e) => { btn.BackColor = Color.FromArgb(225, 240, 255); btn.ForeColor = Color.DarkBlue; };
            btn.MouseLeave += (s, e) => { btn.BackColor = Color.Transparent; btn.ForeColor = Color.Black; };
            return btn;
        }

        private void InitializeEditorUI()
        {
            editorPanel = new Panel { Dock = DockStyle.Fill, Visible = false };

            // 1. Menu Strip
            menuStrip = new MenuStrip { Dock = DockStyle.Top, RenderMode = ToolStripRenderMode.Professional, ForeColor = Color.Black };

            var fileMenu = new ToolStripMenuItem(currentLang == "de" ? "Datei" : "File");
            fileMenu.DropDownItems.Add(currentLang == "de" ? "Neu" : "New", null, (s, e) => { ShowWelcome(); richTextBox.Clear(); });
            fileMenu.DropDownItems.Add(currentLang == "de" ? "Öffnen..." : "Open...", null, OnOpenClick);
            recentFilesMenu = new ToolStripMenuItem(currentLang == "de" ? "Zuletzt verwendet" : "Recent Files");
            fileMenu.DropDownItems.Add(recentFilesMenu);
            fileMenu.DropDownItems.Add(currentLang == "de" ? "Speichern..." : "Save...", null, OnSaveClick);
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add(currentLang == "de" ? "Beenden" : "Exit", null, (s, e) => Application.Exit());

            var formatMenu = new ToolStripMenuItem(currentLang == "de" ? "Format" : "Format");
            formatMenu.DropDownItems.Add("Font Dialog...", null, OnFontClick);
            formatMenu.DropDownItems.Add("Color Dialog...", null, OnColorClick);

            var moreMenu = new ToolStripMenuItem(currentLang == "de" ? "Mehr" : "More") { Name = "More" };

            var toolbarConfig = new ToolStripMenuItem("Toolbar Configuration");
            var itemShowFonts = new ToolStripMenuItem("Show Font Controls") { Checked = true, CheckOnClick = true };
            itemShowFonts.Click += (s, e) => { cmbFontName.Visible = itemShowFonts.Checked; cmbFontSize.Visible = itemShowFonts.Checked; };
            toolbarConfig.DropDownItems.Add(itemShowFonts);

            var themeConfig = new ToolStripMenuItem("Theme Configuration");
            themeConfig.DropDownItems.Add("Windows 10 (UWP)", null, (s, e) => SetTheme("UWP"));
            themeConfig.DropDownItems.Add("UWP Dark Mode", null, (s, e) => SetTheme("UWP Dark"));
            themeConfig.DropDownItems.Add("Aero (Glass)", null, (s, e) => SetTheme("Aero"));
            themeConfig.DropDownItems.Add("Luna (XP)", null, (s, e) => SetTheme("Luna"));
            themeConfig.DropDownItems.Add("Classic (2009)", null, (s, e) => SetTheme("Blue (2009)"));

            var langConfig = new ToolStripMenuItem("Language / Sprache");
            langConfig.DropDownItems.Add("English", null, (s, e) => SwitchLang("en"));
            langConfig.DropDownItems.Add("Deutsch (Beta)", null, (s, e) => SwitchLang("de"));

            moreMenu.DropDownItems.Add(toolbarConfig);
            moreMenu.DropDownItems.Add(themeConfig);
            moreMenu.DropDownItems.Add(langConfig);
            moreMenu.DropDownItems.Add(new ToolStripSeparator());
            moreMenu.DropDownItems.Add("About", null, OnAboutClick);

            menuStrip.Items.AddRange(new ToolStripItem[] { fileMenu, formatMenu, new ToolStripMenuItem("Search", null, OnSearchClick), moreMenu });

            // 2. Toolbar
            toolStrip = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, ForeColor = Color.Black };

            cmbFontName = new ToolStripComboBox { Width = 150, AutoCompleteMode = AutoCompleteMode.SuggestAppend, AutoCompleteSource = AutoCompleteSource.ListItems };
            cmbFontName.ComboBox.DrawMode = DrawMode.OwnerDrawVariable;
            cmbFontName.ComboBox.DrawItem += (s, e) =>
            {
                if (e.Index < 0) return;
                string name = cmbFontName.Items[e.Index].ToString();
                e.DrawBackground();
                using (Font f = new Font(name, 10)) e.Graphics.DrawString(name, f, Brushes.Black, e.Bounds.X, e.Bounds.Y);
            };
            foreach (FontFamily ff in FontFamily.Families) cmbFontName.Items.Add(ff.Name);
            cmbFontName.SelectedIndexChanged += (s, e) => ApplyToolbarFont();

            cmbFontSize = new ToolStripComboBox { Width = 50 };
            cmbFontSize.Items.AddRange(new object[] { "8", "9", "10", "11", "12", "14", "16", "18", "24", "36", "48", "72" });
            cmbFontSize.SelectedIndexChanged += (s, e) => ApplyToolbarFont();

            btnBold = new ToolStripButton("B", null, (s, e) => ToggleStyle(FontStyle.Bold)) { Font = new Font("Times", 10, FontStyle.Bold), CheckOnClick = true };
            btnItalic = new ToolStripButton("I", null, (s, e) => ToggleStyle(FontStyle.Italic)) { Font = new Font("Times", 10, FontStyle.Italic), CheckOnClick = true };
            btnHeading = new ToolStripButton("H", null, (s, e) => ToggleHeading()) { Font = new Font("Arial", 11, FontStyle.Bold), CheckOnClick = true };

            toolStrip.Items.AddRange(new ToolStripItem[] { 
                new ToolStripButton(currentLang == "de" ? "Öffnen" : "Open", null, OnOpenClick), 
                new ToolStripSeparator(),
                cmbFontName, cmbFontSize,
                new ToolStripSeparator(),
                btnBold, btnItalic, btnHeading 
            });

            // 3. Status Strip
            statusStrip = new StatusStrip();
            statusLabel = new ToolStripStatusLabel("Ready");
            pageLabel = new ToolStripStatusLabel(currentLang == "de" ? "Seite: 1" : "Page: 1");
            zoomLabel = new ToolStripLabel("Zoom: 100%");
            statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel, new ToolStripSeparator(), pageLabel, new ToolStripSeparator(), zoomLabel });

            // 4. Editor (A4 Layout)
            richTextBox = new RichTextBox { Size = new Size(794, 1123), Location = new Point(50, 50), BorderStyle = BorderStyle.None, Font = new Font("Calibri", 11.5f) };
            richTextBox.SelectionChanged += (s, e) => UpdateUIStates();
            richTextBox.VScroll += (s, e) =>
            {
                int line = richTextBox.GetLineFromCharIndex(richTextBox.GetCharIndexFromPosition(new Point(0, 0)));
                pageLabel.Text = (currentLang == "de" ? "Seite: " : "Page: ") + ((line / 60) + 1);
            };

            Panel canvasContainer = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(80, 80, 80) };
            canvasContainer.Controls.Add(richTextBox);

            editorPanel.Controls.AddRange(new Control[] { canvasContainer, statusStrip, toolStrip, menuStrip });
            this.Controls.Add(editorPanel);
            UpdateRecent();
        }

        private void ApplyToolbarFont()
        {
            if (isUpdatingFont || richTextBox.SelectionFont == null) return;
            float size;
            if (!float.TryParse(cmbFontSize.Text, out size)) size = richTextBox.SelectionFont.Size;
            try { richTextBox.SelectionFont = new Font(cmbFontName.Text, size, richTextBox.SelectionFont.Style); }
            catch { }
            richTextBox.Focus();
        }

        private void UpdateUIStates()
        {
            isUpdatingFont = true;
            if (richTextBox.SelectionFont != null)
            {
                btnBold.Checked = richTextBox.SelectionFont.Bold;
                btnItalic.Checked = richTextBox.SelectionFont.Italic;
                cmbFontName.Text = richTextBox.SelectionFont.Name;
                cmbFontSize.Text = ((int)richTextBox.SelectionFont.Size).ToString();
            }
            isUpdatingFont = false;
        }

        private void SetTheme(string theme)
        {
            this.BackColor = SystemColors.Control;
            foreach (ToolStripItem i in menuStrip.Items) i.ForeColor = Color.Black;
            foreach (ToolStripItem i in toolStrip.Items) i.ForeColor = Color.Black;

            switch (theme)
            {
                case "Aero":
                    if (IsGlassSupported() && Environment.OSVersion.Version.Major < 10) { this.BackColor = Color.Black; ApplyAero(); }
                    else this.BackColor = Color.FromArgb(200, 220, 240);
                    toolStrip.Renderer = new StateAwareRenderer(); menuStrip.Renderer = new StateAwareRenderer();
                    break;
                case "Luna":
                    this.BackColor = Color.FromArgb(163, 189, 227);
                    toolStrip.Renderer = new LunaRenderer(); menuStrip.Renderer = new LunaRenderer();
                    break;
                case "Blue (2009)":
                    toolStrip.Renderer = new BlueGradientRenderer(); menuStrip.Renderer = new BlueGradientRenderer();
                    break;
                case "UWP":
                    this.BackColor = Color.White;
                    toolStrip.Renderer = new UWPRenderer(); menuStrip.Renderer = new UWPRenderer();
                    break;
                case "UWP Dark":
                    this.BackColor = Color.FromArgb(32, 32, 32);
                    toolStrip.Renderer = new DarkRenderer(); menuStrip.Renderer = new DarkRenderer();
                    foreach (ToolStripItem i in menuStrip.Items) i.ForeColor = Color.White;
                    break;
            }
            this.Invalidate();
        }

        // --- Core Helpers ---
        private void ShowWelcome() { this.BackColor = SystemColors.Control; editorPanel.Visible = false; welcomePanel.Visible = true; }
        private void ShowEditor() { welcomePanel.Visible = false; editorPanel.Visible = true; }
        private void SwitchLang(string l) { currentLang = l; SaveSettings(); Application.Restart(); }

        private void LoadSettings()
        {
            if (File.Exists(configPath))
            {
                string[] lines = File.ReadAllLines(configPath);
                if (lines.Length > 0) currentLang = lines[0];
                for (int i = 1; i < lines.Length; i++) recentFiles.Add(lines[i]);
            }
            else
            {
                currentLang = CultureInfo.CurrentCulture.TwoLetterISOLanguageName == "de" ? "de" : "en";
            }
        }

        private void SaveSettings()
        {
            List<string> output = new List<string> { currentLang };
            output.AddRange(recentFiles);
            File.WriteAllLines(configPath, output.ToArray());
        }

        private void OnOpenClick(object s, EventArgs e)
        {
            using (OpenFileDialog o = new OpenFileDialog { Filter = "Supported|*.rtf;*.txt;*.docx" })
            {
                if (o.ShowDialog() == DialogResult.OK)
                {
                    if (o.FileName.EndsWith(".docx")) { MessageBox.Show("Docx support is in Beta. Some WordArt features may be missing.", "Project Writer Beta"); }
                    ShowEditor(); OpenFile(o.FileName);
                }
            }
        }

        private void OpenFile(string p) { try { if (p.EndsWith(".rtf")) richTextBox.LoadFile(p); else richTextBox.LoadFile(p, RichTextBoxStreamType.PlainText); this.Text = "Project Writer - " + Path.GetFileName(p); AddRecent(p); } catch { } }
        private void AddRecent(string p) { if (recentFiles.Contains(p)) recentFiles.Remove(p); recentFiles.Insert(0, p); if (recentFiles.Count > 10) recentFiles.RemoveAt(10); SaveSettings(); UpdateRecent(); }
        private void UpdateRecent() { recentFilesMenu.DropDownItems.Clear(); foreach (string p in recentFiles) { var i = new ToolStripMenuItem(Path.GetFileName(p)); i.Click += (s, e) => { ShowEditor(); OpenFile(p); }; recentFilesMenu.DropDownItems.Add(i); } }
        private void OnSaveClick(object s, EventArgs e) { using (SaveFileDialog sv = new SaveFileDialog { Filter = "RTF|*.rtf" }) if (sv.ShowDialog() == DialogResult.OK) { richTextBox.SaveFile(sv.FileName); AddRecent(sv.FileName); } }
        private void ToggleStyle(FontStyle s) { if (richTextBox.SelectionFont != null) richTextBox.SelectionFont = new Font(richTextBox.SelectionFont, richTextBox.SelectionFont.Style ^ s); }
        private void ToggleHeading() { if (btnHeading.Checked) { richTextBox.SelectionFont = new Font("Segoe UI", 16, FontStyle.Bold); richTextBox.SelectionColor = Color.FromArgb(0, 120, 215); } else { richTextBox.SelectionFont = new Font("Calibri", 11.5f); richTextBox.SelectionColor = Color.Black; } }
        private void OnFontClick(object s, EventArgs e) { using (FontDialog fd = new FontDialog { Font = richTextBox.SelectionFont }) if (fd.ShowDialog() == DialogResult.OK) richTextBox.SelectionFont = fd.Font; }
        private void OnColorClick(object s, EventArgs e) { using (ColorDialog cd = new ColorDialog { Color = richTextBox.SelectionColor }) if (cd.ShowDialog() == DialogResult.OK) richTextBox.SelectionColor = cd.Color; }
        private void OnSearchClick(object s, EventArgs e) { string t = Interaction.InputBox("Find:", "Search"); if (!string.IsNullOrEmpty(t)) { int i = richTextBox.Find(t); if (i > -1) richTextBox.Select(i, t.Length); } }
        private void OnAboutClick(object s, EventArgs e) { MessageBox.Show("Project Writer 0.6.1 \"Glasswave\"\nGitHub Release\nMade by fanfare.", "About"); }
        private bool IsGlassSupported() { bool e; if (Environment.OSVersion.Version.Major < 6) return false; DwmIsCompositionEnabled(out e); return e; }
        private void ApplyAero() { try { MARGINS m = new MARGINS { cyTopHeight = menuStrip.Height + toolStrip.Height }; DwmExtendFrameIntoClientArea(this.Handle, ref m); } catch { } }
    }

    // --- RENDERERS ---
    public class StateAwareRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
        {
            if (e.Item.Selected || (e.Item is ToolStripButton && ((ToolStripButton)e.Item).Checked))
            {
                Rectangle r = new Rectangle(0, 0, e.Item.Width, e.Item.Height);
                using (LinearGradientBrush b = new LinearGradientBrush(r, Color.FromArgb(255, 240, 190), Color.FromArgb(255, 210, 80), 90f)) e.Graphics.FillRectangle(b, r);
                e.Graphics.DrawRectangle(new Pen(Color.FromArgb(230, 160, 50)), 0, 0, r.Width - 1, r.Height - 1);
            }
            else base.OnRenderButtonBackground(e);
        }
    }
    public class LunaRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { using (LinearGradientBrush b = new LinearGradientBrush(e.AffectedBounds, Color.FromArgb(0, 70, 213), Color.FromArgb(110, 160, 255), 90f)) e.Graphics.FillRectangle(b, e.AffectedBounds); }
        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected || (e.Item is ToolStripButton && ((ToolStripButton)e.Item).Checked)) { e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(61, 149, 38)), new Rectangle(0, 0, e.Item.Width, e.Item.Height)); e.Graphics.DrawRectangle(Pens.White, 0, 0, e.Item.Width - 1, e.Item.Height - 1); } }
    }
    public class BlueGradientRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { using (LinearGradientBrush b = new LinearGradientBrush(e.AffectedBounds, Color.FromArgb(215, 230, 250), Color.FromArgb(170, 195, 230), 90f)) e.Graphics.FillRectangle(b, e.AffectedBounds); }
    }
    public class UWPRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { e.Graphics.Clear(Color.White); using (Pen p = new Pen(Color.FromArgb(230, 230, 230))) e.Graphics.DrawLine(p, 0, e.ToolStrip.Height - 1, e.ToolStrip.Width, e.ToolStrip.Height - 1); }
        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected || (e.Item is ToolStripButton && ((ToolStripButton)e.Item).Checked)) e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(230, 240, 255)), new Rectangle(0, 0, e.Item.Width, e.Item.Height)); }
    }
    public class DarkRenderer : ToolStripProfessionalRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { e.Graphics.Clear(Color.FromArgb(45, 45, 45)); }
        protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected) e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(80, 80, 80)), new Rectangle(0, 0, e.Item.Width, e.Item.Height)); }
    }
}