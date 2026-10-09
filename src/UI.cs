using System;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.Web.Script.Serialization;

namespace LetterMerger
{
    public sealed class Preferences
    {
        public int Version = 1, StartRow = 1, RowCount = 5;
        public bool AllRows = false, CreatePdf = false, RepairText = true, SeparateMissingPhotos = false;
        public string CsvPath = "", TemplatePath = "", PhotoMode = "Fit entire photo";
        public Dictionary<string, string> Columns = new Dictionary<string, string>
        {
            {
                "Name",
                ""
            },
            {
                "Scholarship",
                ""
            },
            {
                "Message",
                ""
            },
            {
                "Photo",
                ""
            }
        };
        public Overlay Overlay = new Overlay();
    }

    public static class Settings
    {
        public static Preferences Read(string path)
        {
            var p = new Preferences();
            if (!File.Exists(path))
                return p;
            Core.RejectLink(path);
            if (new FileInfo(path).Length > 1048576)
                throw new InvalidDataException("Saved settings are too large.");
            var serializer = new JavaScriptSerializer
            {
                MaxJsonLength = 1048576,
                RecursionLimit = 20
            };
            var d = serializer.DeserializeObject(File.ReadAllText(path)) as Dictionary<string, object>;
            if (d == null || !d.ContainsKey("Version") || Convert.ToString(d["Version"], CultureInfo.InvariantCulture) != "1")
                throw new InvalidDataException("Saved settings have an unsupported format.");
            foreach (string key in new[]
            {
                "CsvPath",
                "TemplatePath",
                "PhotoMode"
            }

            )
            {
                object value;
                if (d.TryGetValue(key, out value) && value is string)
                {
                    if (key == "CsvPath")
                        p.CsvPath = (string)value;
                    if (key == "TemplatePath")
                        p.TemplatePath = (string)value;
                    if (key == "PhotoMode" && Core.Modes.Contains((string)value))
                        p.PhotoMode = (string)value;
                }
            }

            foreach (string key in new[]
            {
                "StartRow",
                "RowCount"
            }

            )
            {
                object value;
                int number;
                if (d.TryGetValue(key, out value) && Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out number) && number >= 1 && number <= 1000000)
                {
                    if (key == "StartRow")
                        p.StartRow = number;
                    else
                        p.RowCount = number;
                }
            }

            object item;
            if (d.TryGetValue("AllRows", out item) && item is bool)
                p.AllRows = (bool)item;
            if (d.TryGetValue("CreatePdf", out item) && item is bool)
                p.CreatePdf = (bool)item;
            if (d.TryGetValue("RepairText", out item) && item is bool)
                p.RepairText = (bool)item;
            if (d.TryGetValue("SeparateMissingPhotos", out item) && item is bool)
                p.SeparateMissingPhotos = (bool)item;
            if (d.TryGetValue("Columns", out item) && item is Dictionary<string, object>)
            {
                var cols = (Dictionary<string, object>)item;
                foreach (string key in p.Columns.Keys.ToArray())
                    if (cols.TryGetValue(key, out item) && item is string)
                        p.Columns[key] = (string)item;
            }

            if (d.TryGetValue("Overlay", out item) && item is Dictionary<string, object>)
            {
                var overlay = (Dictionary<string, object>)item;
                foreach (string key in new[]
                {
                    "Width",
                    "Height",
                    "X",
                    "Y"
                }

                )
                {
                    double v;
                    if (overlay.TryGetValue(key, out item) && Double.TryParse(Convert.ToString(item, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out v) && !Double.IsNaN(v) && !Double.IsInfinity(v) && ((key == "Width" || key == "Height") ? v >= .25 && v <= 10 : Math.Abs(v) <= .5))
                    {
                        if (key == "Width")
                            p.Overlay.Width = v;
                        if (key == "Height")
                            p.Overlay.Height = v;
                        if (key == "X")
                            p.Overlay.X = v;
                        if (key == "Y")
                            p.Overlay.Y = v;
                    }
                }
            }

            return p;
        }

        public static void Write(string path, Preferences p)
        {
            Core.RejectLink(path);
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, new JavaScriptSerializer().Serialize(p), new UTF8Encoding(false));
                if (File.Exists(path))
                {
                    Core.RejectLink(path);
                    File.Replace(temporary, path, null);
                }
                else
                    File.Move(temporary, path);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }

        public static string StorePath(string root, string path)
        {
            if (String.IsNullOrEmpty(path))
                return "";
            string prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string full = Path.GetFullPath(path);
            return full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? full.Substring(prefix.Length) : full;
        }

        public static string RestorePath(string root, string path, string ext)
        {
            if (String.IsNullOrEmpty(path) || path.Any(c => c < 32))
                return "";
            try
            {
                string full = Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(root, path.Replace('\\', Path.DirectorySeparatorChar)));
                return File.Exists(full) && Path.GetExtension(full).Equals(ext, StringComparison.OrdinalIgnoreCase) ? full : "";
            }
            catch
            {
                return "";
            }
        }
    }

    public sealed class MainForm : Form
    {
        public readonly string Root;
        readonly Dictionary<string, string> folders = new Dictionary<string, string>();
        readonly Panel body = new Panel();
        readonly ToolTip tips = new ToolTip();
        readonly Timer saveTimer = new Timer();
        readonly ComboBox csvBox, templateBox, modeBox;
        readonly Dictionary<string, ComboBox> fields = new Dictionary<string, ComboBox>();
        readonly NumericUpDown startBox, countBox;
        readonly CheckBox allBox, pdfBox, repairBox, separateBox;
        readonly Label rangeLabel;
        readonly TextBox status;
        readonly Button previewButton;
        readonly List<string> extraCsv = new List<string>(), extraTemplates = new List<string>();
        Preferences preferences = new Preferences();
        bool ready = false, busy = false, loading = false;
        CsvData loaded;
        PhotoPreview activePreview;
        public static readonly Color Ink = ColorTranslator.FromHtml("#1F2937"), SoftAccent = ColorTranslator.FromHtml("#D1D5DB"), Accent = ColorTranslator.FromHtml("#2563EB");
        static Label LabelAt(Control parent, string text, int x, int y, int w, int h)
        {
            var c = new Label
            {
                Text = text,
                ForeColor = Ink
            };
            c.SetBounds(x, y, w, h);
            parent.Controls.Add(c);
            return c;
        }

        static Button ButtonAt(Control parent, string text, int x, int y, int w, Action action)
        {
            var c = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Ink
            };
            c.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#CBD5E1");
            c.SetBounds(x, y, w, 34);
            c.Click += (s, e) =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.GetBaseException().Message, "Letter Merger", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            parent.Controls.Add(c);
            return c;
        }

        static ComboBox ComboAt(Control parent, int x, int y, int w)
        {
            var c = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                MaxDropDownItems = 10,
                IntegralHeight = false,
                DropDownWidth = w
            };
            c.SetBounds(x, y, w, 28);
            c.DropDown += (s, e) => c.DropDownWidth = c.Width;
            c.SizeChanged += (s, e) => c.DropDownWidth = Math.Max(1, c.Width);
            parent.Controls.Add(c);
            return c;
        }

        static NumericUpDown NumberAt(Control parent, int x, int y, int width, int maximum, int value)
        {
            var c = new NumericUpDown
            {
                Minimum = 1,
                Maximum = maximum,
                Value = value
            };
            c.SetBounds(x, y, width, 30);
            parent.Controls.Add(c);
            return c;
        }

        public MainForm(string root)
        {
            Root = Path.GetFullPath(root);
            Text = "Letter Merger";
            Font = new Font("Segoe UI", 10);
            AutoScaleMode = AutoScaleMode.Dpi;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96, 96);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = MinimizeBox = true;
            MinimumSize = new Size(790, 510);
            ClientSize = new Size(950, Math.Min(830, Screen.PrimaryScreen.WorkingArea.Height - 70));
            BackColor = ColorTranslator.FromHtml("#F8FAFC");
            foreach (string name in new[]
            {
                "Data",
                "Photos",
                "Templates",
                "Output"
            }

            )
            {
                string path = Path.Combine(Root, name);
                if (Directory.Exists(path))
                    Core.RejectLink(path);
                Directory.CreateDirectory(path);
                folders[name] = path;
            }

            var shell = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = Padding.Empty };
            shell.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            shell.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));
            shell.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            shell.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            Controls.Add(shell);
            var header = new Panel { Dock = DockStyle.Fill, BackColor = Ink, Margin = Padding.Empty };
            shell.Controls.Add(header, 0, 0);
            var title = LabelAt(header, "LETTER MERGER", 22, 9, 690, 29);
            title.Font = new Font("Segoe UI", 17, FontStyle.Bold);
            title.ForeColor = Color.White;
            title.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            var subtitle = LabelAt(header, "A7 photo letters from CSV data", 22, 43, 690, 23);
            subtitle.ForeColor = SoftAccent;
            subtitle.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            header.Controls.Add(new Panel { Dock = DockStyle.Bottom, Height = 4, BackColor = Accent });
            var footer = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
            shell.Controls.Add(footer, 0, 2);
            var signature = LabelAt(footer, "-Kevin Le", 22, 4, 180, 20);
            signature.Font = new Font("Segoe UI", 9);
            signature.ForeColor = Color.DimGray;
            body.Dock = DockStyle.Fill;
            body.AutoScroll = true;
            body.Padding = new Padding(20, 14, 20, 10);
            body.Margin = Padding.Empty;
            shell.Controls.Add(body, 0, 1);
            var content = ViewLayout.Rows();
            content.MinimumSize = new Size(720, 0);
            body.Controls.Add(content);
            body.SizeChanged += (s, e) => content.MinimumSize = new Size(720, Math.Max(0, body.ClientSize.Height - body.Padding.Vertical));
            ViewLayout.Add(content, ViewLayout.Text("1. Choose files and map columns     2. Preview photos     3. Generate and review"), false);
            var folderButtons = new List<Control>();
            foreach (string name in new[] { "Data", "Photos", "Templates", "Output" })
            {
                string target = folders[name];
                folderButtons.Add(ButtonAt(body, "Open " + name, 0, 0, 158, () => OpenPath(target)));
            }
            ViewLayout.Add(content, ViewLayout.Flow(folderButtons.ToArray()), false);
            var dataLabel = ViewLayout.Text("Data folder: " + folders["Data"]);
            dataLabel.Font = new Font("Segoe UI", 8);
            dataLabel.AutoEllipsis = true;
            tips.SetToolTip(dataLabel, folders["Data"]);
            ViewLayout.Add(content, dataLabel, false);
            var files = new TableLayoutPanel { ColumnCount = 4, AutoSize = true, Dock = DockStyle.Fill, Margin = Padding.Empty };
            files.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            files.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            files.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 135));
            files.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 125));
            csvBox = ComboAt(body, 0, 0, 350);
            csvBox.DisplayMember = "Name";
            templateBox = ComboAt(body, 0, 0, 350);
            templateBox.DisplayMember = "Name";
            Action<Control, int, int, int> addFile = (control, column, rowIndex, span) =>
            {
                control.Dock = DockStyle.Fill;
                control.Margin = new Padding(0, 3, 8, 8);
                files.Controls.Add(control, column, rowIndex);
                files.SetColumnSpan(control, span);
            };
            addFile(ViewLayout.Text("1. CSV data"), 0, 0, 1);
            addFile(csvBox, 1, 0, 1);
            addFile(ButtonAt(body, "Browse CSV", 0, 0, 125, () => Browse(false)), 2, 0, 1);
            addFile(ButtonAt(body, "Refresh files", 0, 0, 115, RefreshFiles), 3, 0, 1);
            addFile(ViewLayout.Text("A7 Word template"), 0, 1, 1);
            addFile(templateBox, 1, 1, 1);
            addFile(ButtonAt(body, "Browse template", 0, 0, 125, () => Browse(true)), 2, 1, 1);
            addFile(ButtonAt(body, "Load columns", 0, 0, 115, () => LoadColumns(false)), 3, 1, 1);
            int fieldRow = 2;
            foreach (var pair in new[] { new[] { "Name", "Student name" }, new[] { "Scholarship", "Scholarship" }, new[] { "Message", "Thank-you text" }, new[] { "Photo", "Photo path" } })
            {
                addFile(ViewLayout.Text(pair[1]), 0, fieldRow, 1);
                fields[pair[0]] = ComboAt(body, 0, 0, 560);
                addFile(fields[pair[0]], 1, fieldRow++, 3);
            }
            ViewLayout.Add(content, files, false);
            startBox = NumberAt(body, 0, 0, 85, 1000000, 1);
            countBox = NumberAt(body, 0, 0, 85, 1000000, 5);
            allBox = new CheckBox { Text = "All rows", AutoSize = true };
            ViewLayout.Add(content, ViewLayout.Flow(ViewLayout.Text("Start row"), startBox, ViewLayout.Text("Rows to merge"), countBox, allBox), false);
            rangeLabel = ViewLayout.Text("Load columns to see the selected range.");
            ViewLayout.Add(content, rangeLabel, false);
            modeBox = ComboAt(body, 0, 0, 240);
            modeBox.Items.AddRange(Core.Modes);
            modeBox.SelectedIndex = 0;
            previewButton = ButtonAt(body, "Preview photo", 0, 0, 150, Preview);
            previewButton.BackColor = Ink;
            previewButton.ForeColor = Color.White;
            ViewLayout.Add(content, ViewLayout.Flow(ViewLayout.Text("2. Photo sizing"), modeBox, previewButton, ButtonAt(body, "Restore Defaults", 0, 0, 155, RestoreDefaults)), false);
            ViewLayout.Add(content, ViewLayout.Text("Missing-photo reporting scans every row. Preview checks photos; Word checks letter layout."), false);
            repairBox = new CheckBox { Text = "Repair common encoding errors in thank-you text", AutoSize = true, Checked = true };
            separateBox = new CheckBox { Text = "Put letters with missing photos in a separate DOCX", AutoSize = true };
            ViewLayout.Add(content, ViewLayout.Flow(repairBox, separateBox), false);
            tips.SetToolTip(repairBox, "Fix recognizable garbled punctuation/accents without rewriting the message. Exact name and scholarship wording stay unchanged. TextRepairReport.csv records repairs and issues.");
            tips.SetToolTip(separateBox, "Merged_Letters.docx contains photos. Merged_Letters_Missing.docx contains missing/unreadable photos. An empty group produces no file. Selected rows still apply.");
            pdfBox = new CheckBox { Text = "Also create a PDF when every generated card passes the Word layout check", AutoSize = true };
            ViewLayout.Add(content, pdfBox, false);
            var generate = ButtonAt(body, "Generate selected rows", 0, 0, 225, () => Run(true));
            generate.BackColor = Ink;
            generate.ForeColor = Color.White;
            ViewLayout.Add(content, ViewLayout.Flow(ButtonAt(body, "Check data and photos", 0, 0, 210, () => Run(false)), generate, ButtonAt(body, "Latest output", 0, 0, 130, () => OpenLatest("")), ButtonAt(body, "Open reports", 0, 0, 125, ReportMenu)), false);
            status = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, WordWrap = true, Height = 110, MinimumSize = new Size(0, 90), Text = "Add files, load columns, confirm mappings and preview a photo." };
            ViewLayout.Add(content, status, true);
            saveTimer.Interval = 500;
            saveTimer.Tick += (s, e) =>
            {
                saveTimer.Stop();
                Save();
            };
            csvBox.SelectedIndexChanged += (s, e) => FileChanged(csvBox);
            templateBox.SelectedIndexChanged += (s, e) => FileChanged(templateBox);
            foreach (var combo in fields.Values)
            {
                combo.SelectedIndexChanged += (s, e) =>
                {
                    tips.SetToolTip((Control)s, Convert.ToString(((ComboBox)s).SelectedItem));
                    QueueSave();
                };
                combo.DropDown += (s, e) => tips.SetToolTip((Control)s, Convert.ToString(((ComboBox)s).SelectedItem));
            }

            startBox.ValueChanged += (s, e) =>
            {
                UpdateRange();
                QueueSave();
            };
            countBox.ValueChanged += (s, e) =>
            {
                UpdateRange();
                QueueSave();
            };
            allBox.CheckedChanged += (s, e) =>
            {
                startBox.Enabled = countBox.Enabled = !allBox.Checked;
                UpdateRange();
                QueueSave();
            };
            modeBox.SelectedIndexChanged += (s, e) => QueueSave();
            pdfBox.CheckedChanged += (s, e) => QueueSave();
            repairBox.CheckedChanged += (s, e) => QueueSave();
            separateBox.CheckedChanged += (s, e) => QueueSave();
            tips.SetToolTip(modeBox, "Fit keeps the entire photo. Fill crops inside the opening. Overlay fills a larger adjustable rectangle over curved corners.");
            FormClosing += (s, e) =>
            {
                if (activePreview != null)
                    activePreview.Close();
                if (busy)
                    e.Cancel = true;
                else
                {
                    saveTimer.Stop();
                    Save();
                }
            };
            try
            {
                preferences = Settings.Read(Path.Combine(Root, "LetterMergerSettings.json"));
            }
            catch (Exception e)
            {
                status.Text = "Saved options could not be loaded. Defaults are in use. " + e.Message;
            }

            loading = true;
            try
            {
                RefreshFiles();
                RestoreSelection(csvBox, extraCsv, preferences.CsvPath, ".csv");
                RestoreSelection(templateBox, extraTemplates, preferences.TemplatePath, ".docx");
                startBox.Value = preferences.StartRow;
                countBox.Value = preferences.RowCount;
                allBox.Checked = preferences.AllRows;
                modeBox.SelectedItem = preferences.PhotoMode;
                pdfBox.Checked = preferences.CreatePdf;
                repairBox.Checked = preferences.RepairText;
                separateBox.Checked = preferences.SeparateMissingPhotos;
                if (csvBox.SelectedItem != null && templateBox.SelectedItem != null)
                {
                    LoadColumns(true);
                    foreach (string key in fields.Keys)
                        if (fields[key].Items.Contains(preferences.Columns[key]))
                            fields[key].SelectedItem = preferences.Columns[key];
                }
            }
            catch (Exception e)
            {
                status.Text = "Saved selections need review. " + e.GetBaseException().Message;
            }
            finally
            {
                loading = false;
                ready = true;
            }

            UpdateRange();
        }

        void Notice(string message)
        {
            MessageBox.Show(message, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void Progress(string message)
        {
            status.Text = message;
            Application.DoEvents();
        }

        void FileChanged(ComboBox combo)
        {
            if (loading)
                return;
            loaded = null;
            foreach (var c in fields.Values)
                c.Items.Clear();
            status.Text = "File selected. Click Load columns and confirm mappings.";
            tips.SetToolTip(combo, SelectedPath(combo));
            UpdateRange();
            QueueSave();
        }

        static string SelectedPath(ComboBox combo)
        {
            var f = combo.SelectedItem as FileInfo;
            return f == null ? "" : f.FullName;
        }

        void Sync(ComboBox combo, string folder, string extension, List<string> extras)
        {
            string old = SelectedPath(combo);
            var paths = Directory.EnumerateFiles(folder).Where(p => Path.GetExtension(p).Equals(extension, StringComparison.OrdinalIgnoreCase)).Concat(extras.Where(File.Exists)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(p => Path.GetFileName(p), StringComparer.OrdinalIgnoreCase).ToArray();
            combo.Items.Clear();
            foreach (string path in paths)
                combo.Items.Add(new FileInfo(path));
            if (old.Length > 0)
            {
                var f = combo.Items.Cast<FileInfo>().FirstOrDefault(i => i.FullName.Equals(old, StringComparison.OrdinalIgnoreCase));
                if (f != null)
                    combo.SelectedItem = f;
            }
            else if (combo.Items.Count > 0)
                combo.SelectedIndex = 0;
        }

        void RefreshFiles()
        {
            bool prior = loading;
            loading = true;
            try
            {
                Sync(csvBox, folders["Data"], ".csv", extraCsv);
                Sync(templateBox, folders["Templates"], ".docx", extraTemplates);
            }
            finally
            {
                loading = prior;
            }
        }

        void RestoreSelection(ComboBox combo, List<string> extras, string stored, string extension)
        {
            if (stored.Length == 0)
                return;
            string path = Settings.RestorePath(Root, stored, extension);
            if (path.Length == 0)
            {
                combo.SelectedIndex = -1;
                status.Text = "A saved file is missing. Choose it and load columns.";
                return;
            }

            if (!extras.Contains(path))
                extras.Add(path);
            var item = combo.Items.Cast<FileInfo>().FirstOrDefault(i => i.FullName.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (item == null)
            {
                item = new FileInfo(path);
                combo.Items.Add(item);
            }

            combo.SelectedItem = item;
        }

        void Browse(bool template)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = template ? "Word template (*.docx)|*.docx" : "CSV UTF-8 (*.csv)|*.csv",
                CheckFileExists = true,
                Multiselect = false,
                InitialDirectory = template ? folders["Templates"] : folders["Data"]
            }

            )
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var extras = template ? extraTemplates : extraCsv;
                    var combo = template ? templateBox : csvBox;
                    if (!extras.Contains(dialog.FileName))
                        extras.Add(dialog.FileName);
                    RefreshFiles();
                    combo.SelectedItem = combo.Items.Cast<FileInfo>().First(i => i.FullName.Equals(dialog.FileName, StringComparison.OrdinalIgnoreCase));
                    FileChanged(combo);
                }
        }

        static string Suggest(string role, List<CsvColumn> columns)
        {
            foreach (var c in columns)
            {
                if (c.Status == "REPEATED_HEADER")
                    continue;
                string value = Core.Canonical(c.Label);
                if (role == "Name" && new[]
                {
                    "name",
                    "studentname",
                    "fullname"
                }.Contains(value))
                    return c.Label;
                if (role == "Scholarship" && new[]
                {
                    "portfolioname",
                    "scholarship",
                    "scholarshipname"
                }.Contains(value))
                    return c.Label;
                if (role == "Message" && (value.StartsWith("pleasedraftathankyouletter") || new[]
                {
                    "thankyoutext",
                    "thankyouletter",
                    "message"
                }.Contains(value)))
                    return c.Label;
                if (role == "Photo" && (value.StartsWith("uploadaclearappropriatephotoofyourself") || new[]
                {
                    "photopath",
                    "photo",
                    "headshot"
                }.Contains(value)))
                    return c.Label;
            }

            return "";
        }

        void LoadColumns(bool quiet)
        {
            if (csvBox.SelectedItem == null || templateBox.SelectedItem == null)
                throw new InvalidDataException("Choose the CSV and the original A7 template first.");
            Core.ReadTemplate(SelectedPath(templateBox));
            var data = Core.ReadCsv(SelectedPath(csvBox));
            bool prior = loading;
            loading = true;
            try
            {
                foreach (string key in fields.Keys)
                {
                    var combo = fields[key];
                    string old = Convert.ToString(combo.SelectedItem);
                    combo.Items.Clear();
                    combo.Items.AddRange(data.Columns.Select(c => (object)c.Label).ToArray());
                    string selected = combo.Items.Contains(old) ? old : Suggest(key, data.Columns);
                    if (combo.Items.Contains(selected))
                        combo.SelectedItem = selected;
                }

                loaded = data;
            }
            finally
            {
                loading = prior;
            }

            UpdateRange();
            status.Text = "Loaded " + data.Rows.Count + " nonblank data rows. Confirm all four columns.";
            if (!quiet && data.Columns.Any(c => c.Status == "BLANK_HEADER" || c.Status == "REPEATED_HEADER"))
                Notice("Blank/repeated headings have unique position labels. All columns are retained. Confirm each mapping; CSVColumns.csv records original headings.");
            QueueSave();
        }

        Dictionary<string, int> Map(CsvData data)
        {
            var map = new Dictionary<string, int>();
            foreach (string key in fields.Keys)
            {
                string label = Convert.ToString(fields[key].SelectedItem);
                int index = data.Columns.FindIndex(c => c.Label == label);
                if (index < 0)
                    throw new InvalidDataException("Load columns and confirm all four mappings.");
                map[key] = index;
            }

            if (map.Values.Distinct().Count() != 4)
                throw new InvalidDataException("Choose four distinct columns.");
            return map;
        }

        void UpdateRange()
        {
            if (loaded == null)
            {
                rangeLabel.Text = "Load columns to see the range. Header and blank records do not count.";
                return;
            }

            try
            {
                var selection = Core.Select(loaded.Rows.Count, (int)startBox.Value, (int)countBox.Value, allBox.Checked);
                rangeLabel.Text = "Selected data rows " + selection.Start + "-" + selection.End + " of " + loaded.Rows.Count + " (" + selection.Count + " records).";
            }
            catch (Exception e)
            {
                rangeLabel.Text = e.Message;
            }
        }

        void QueueSave()
        {
            if (!ready || loading || busy)
                return;
            saveTimer.Stop();
            saveTimer.Start();
        }

        void Save()
        {
            if (!ready || loading || IsDisposed)
                return;
            try
            {
                string csv = SelectedPath(csvBox), template = SelectedPath(templateBox);
                if (csv.Length > 0)
                    preferences.CsvPath = Settings.StorePath(Root, csv);
                if (template.Length > 0)
                    preferences.TemplatePath = Settings.StorePath(Root, template);
                foreach (string key in fields.Keys)
                    if (fields[key].Items.Count > 0)
                        preferences.Columns[key] = Convert.ToString(fields[key].SelectedItem);
                preferences.StartRow = (int)startBox.Value;
                preferences.RowCount = (int)countBox.Value;
                preferences.AllRows = allBox.Checked;
                preferences.PhotoMode = Convert.ToString(modeBox.SelectedItem);
                preferences.CreatePdf = pdfBox.Checked;
                preferences.RepairText = repairBox.Checked;
                preferences.SeparateMissingPhotos = separateBox.Checked;
                Settings.Write(Path.Combine(Root, "LetterMergerSettings.json"), preferences);
            }
            catch (Exception e)
            {
                status.Text = "Options could not be saved. You can keep working. " + e.Message;
                App.LogException(e);
            }
        }

        public void RestoreDefaults()
        {
            if (busy)
                return;
            saveTimer.Stop();
            loading = true;
            try
            {
                preferences = new Preferences();
                extraCsv.Clear();
                extraTemplates.Clear();
                loaded = null;
                csvBox.SelectedIndex = -1;
                templateBox.SelectedIndex = -1;
                csvBox.Items.Clear();
                templateBox.Items.Clear();
                foreach (var combo in fields.Values)
                    combo.Items.Clear();
                startBox.Value = 1;
                countBox.Value = 5;
                allBox.Checked = false;
                modeBox.SelectedIndex = 0;
                pdfBox.Checked = false;
                repairBox.Checked = true;
                separateBox.Checked = false;
                RefreshFiles();
                UpdateRange();
                status.Text = "Defaults restored. Start 1, count 5, Fit, centered overlay 2.10 x 2.60 inches, PDF off, text repair on, photo separation off. Load columns again. Input/output files are preserved.";
            }
            finally
            {
                loading = false;
            }

            Save();
        }

        public bool DefaultsCorrect()
        {
            return startBox.Value == 1 && countBox.Value == 5 && !allBox.Checked && !pdfBox.Checked && repairBox.Checked && !separateBox.Checked && modeBox.SelectedIndex == 0 && preferences.Overlay.Width == 2.10 && preferences.Overlay.Height == 2.60 && preferences.Overlay.X == 0 && preferences.Overlay.Y == 0;
        }

        void SetBusy(bool value)
        {
            busy = value;
            EnableInputs(body, !value);
            startBox.Enabled = countBox.Enabled = !value && !allBox.Checked;
            UseWaitCursor = value && activePreview == null;
        }

        static void EnableInputs(Control parent, bool enabled)
        {
            foreach (Control control in parent.Controls)
            {
                if (control is Button || control is ComboBox || control is NumericUpDown || control is CheckBox)
                    control.Enabled = enabled;
                else
                    EnableInputs(control, enabled);
            }
        }

        void OpenPath(string path)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT)
                throw new PlatformNotSupportedException("Opening outputs in their default apps requires Windows.");
            Core.RejectLink(path);
            Process.Start(new ProcessStartInfo { FileName = Path.GetFullPath(path), UseShellExecute = true });
        }

        string Latest()
        {
            Core.RejectLink(folders["Output"]);
            var batch = Directory.EnumerateDirectories(folders["Output"]).Select(p => new DirectoryInfo(p)).Where(d => Regex.IsMatch(d.Name, "^Batch_[0-9]{8}_[0-9]{6}_[0-9a-fA-F]{6}$") && (d.Attributes & FileAttributes.ReparsePoint) == 0).OrderByDescending(d => d.LastWriteTimeUtc).FirstOrDefault();
            if (batch == null)
                throw new FileNotFoundException("No output batch exists yet.");
            return batch.FullName;
        }

        static readonly string[] Reports =
        {
            "Merged_Letters.docx",
            "Merged_Letters.pdf",
            "Merged_Letters_Missing.docx",
            "Merged_Letters_Missing.pdf",
            "TextRepairReport.csv",
            "DocumentValidation_Merged_Letters.txt",
            "DocumentValidation_Merged_Letters_Missing.txt",
            "MissingPhotos_AllRows.csv",
            "MergeReport.csv",
            "LayoutReport.csv",
            "RunSummary.txt",
            "CSVColumns.csv",
            "DocumentValidation.txt"
        };
        void OpenLatest(string name)
        {
            if (name.Length > 0 && !Reports.Contains(name))
                throw new InvalidDataException("Unsupported report.");
            string batch = Latest(), path = name.Length == 0 ? batch : Path.Combine(batch, name);
            if (name.Length > 0 && !File.Exists(path))
                throw new FileNotFoundException("This file was not produced in the latest batch: " + name);
            OpenPath(path);
        }

        void ReportMenu()
        {
            var menu = new ContextMenuStrip();
            foreach (string name in Reports)
            {
                string selected = name;
                menu.Items.Add(name, null, (s, e) =>
                {
                    try
                    {
                        OpenLatest(selected);
                    }
                    catch (Exception ex)
                    {
                        Notice(ex.Message);
                    }
                });
            }

            menu.Show(Cursor.Position);
        }

        void Preview()
        {
            if (activePreview != null)
            {
                if (activePreview.WindowState == FormWindowState.Minimized)
                    activePreview.WindowState = FormWindowState.Normal;
                activePreview.Activate();
                return;
            }
            if (busy)
                return;
            saveTimer.Stop();
            Save();
            var data = Core.ReadCsv(SelectedPath(csvBox));
            var t = Core.ReadTemplate(SelectedPath(templateBox));
            var map = Map(data);
            var selection = Core.Select(data.Rows.Count, (int)startBox.Value, (int)countBox.Value, allBox.Checked);
            var dialog = new PhotoPreview(t, data, map, selection, folders["Photos"], preferences.Overlay, Convert.ToString(modeBox.SelectedItem), choice =>
            {
                preferences.Overlay = choice.Item1;
                modeBox.SelectedItem = choice.Item2;
                status.Text = "Photo placement applied to all photos in the next selected batch.";
            });
            activePreview = dialog;
            dialog.FormClosed += (s, e) =>
            {
                activePreview = null;
                SetBusy(false);
                Save();
            };
            SetBusy(true);
            try
            {
                dialog.Show();
            }
            catch
            {
                activePreview = null;
                dialog.Dispose();
                SetBusy(false);
                throw;
            }
        }

        void Run(bool generate)
        {
            if (busy)
                return;
            saveTimer.Stop();
            Save();
            SetBusy(true);
            string batch = null, error = "", validationWarning = "";
            int generated = 0;
            CsvData data = null;
            Selection selection = null;
            var prepared = new List<Student>();
            var selected = new List<Student>();
            var missing = new List<string[]>();
            var layout = new List<LayoutRow>();
            var failures = new List<string>();
            var merge = new MergeBatchResult();
            string mode = Convert.ToString(modeBox.SelectedItem);
            try
            {
                data = Core.ReadCsv(SelectedPath(csvBox));
                var t = Core.ReadTemplate(SelectedPath(templateBox));
                Core.Box(t, mode, preferences.Overlay);
                var map = Map(data);
                selection = Core.Select(data.Rows.Count, (int)startBox.Value, (int)countBox.Value, allBox.Checked);
                batch = Path.Combine(folders["Output"], "Batch_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6));
                Core.RejectLink(folders["Output"]);
                Core.RejectLink(batch);
                Directory.CreateDirectory(batch);
                var owners = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var row in data.Rows)
                {
                    try
                    {
                        string path = Core.PhotoPath(folders["Photos"], row.Values[map["Photo"]]), owner = Core.PhotoOwner(row.Values[map["Name"]]);
                        if (owner.Length == 0)
                            continue;
                        if (!owners.ContainsKey(path))
                            owners[path] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        owners[path].Add(owner);
                    }
                    catch
                    {
                    }
                }

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    var row = data.Rows[i];
                    var s = new Student
                    {
                        Record = row.Record,
                        DataRow = i + 1,
                        Name = row.Values[map["Name"]],
                        Scholarship = row.Values[map["Scholarship"]],
                        Message = row.Values[map["Message"]],
                        PhotoPath = row.Values[map["Photo"]],
                        Selected = i + 1 >= selection.Start && i + 1 <= selection.End
                    };
                    Progress("Checking photos for all " + data.Rows.Count + " rows (row " + (i + 1) + "). Selected rows " + selection.Start + "-" + selection.End + " for " + (generate ? "merging" : "data checks") + ".");
                    try
                    {
                        string key = Core.PhotoPath(folders["Photos"], s.PhotoPath);
                        if (owners.ContainsKey(key) && owners[key].Count > 1)
                            throw new InvalidDataException("PHOTO_CONFLICT: this path belongs to different student names.");
                        s.Photo = Core.ResolvePhoto(folders["Photos"], s.PhotoPath);
                        s.PhotoDetails = "Photo found and readable.";
                    }
                    catch (Exception e)
                    {
                        s.Photo = "";
                        s.PhotoStatus = "PHOTO_UNAVAILABLE";
                        s.PhotoDetails = e.GetBaseException().Message;
                        missing.Add(new[] { s.Record.ToString(), s.DataRow.ToString(), s.Name, s.Scholarship, s.PhotoPath, s.Selected.ToString(), s.PhotoDetails });
                    }

                    if (!s.Selected)
                        continue;
                    selected.Add(s);
                    TextRepair.Prepare(s, repairBox.Checked);
                    var problems = new List<string>();
                    if (String.IsNullOrWhiteSpace(s.Name))
                        problems.Add("Student name is blank.");
                    if (String.IsNullOrWhiteSpace(s.Scholarship))
                        problems.Add("Scholarship is blank.");
                    if (String.IsNullOrWhiteSpace(s.Message))
                        problems.Add("Thank-you response is blank.");
                    foreach (string text in new[]
                    {
                        s.Name,
                        s.Scholarship,
                        s.Message
                    }

                    )
                        try
                        {
                            System.Xml.XmlConvert.VerifyXmlChars(text);
                        }
                        catch
                        {
                            problems.Add("Invalid XML character in supplied text.");
                        }

                    if (problems.Count > 0)
                    {
                        s.PhotoDetails = "SKIPPED_DATA: " + String.Join(" ", problems) + " " + s.PhotoDetails;
                        continue;
                    }

                    prepared.Add(s);
                }

                if (generate && prepared.Count > 0)
                {
                    merge = BatchMerge.Build(t.Source, prepared, batch, mode, preferences.Overlay, separateBox.Checked, Progress);
                    generated = merge.Outputs.Sum(o => o.Cards.Count);
                    failures.AddRange(merge.Errors);
                    foreach (var student in prepared.Where(s => s.PhotoStatus == "PHOTO_INSERT_ERROR"))
                        if (!missing.Any(r => r[0] == student.Record.ToString()))
                            missing.Add(new[] { student.Record.ToString(), student.DataRow.ToString(), student.Name, student.Scholarship, student.PhotoPath, "True", student.PhotoDetails });
                    foreach (var output in merge.Outputs)
                        layout.AddRange(WordLayout.Check(output.Path, output.Cards, output.Fold, pdfBox.Checked, batch, Progress));
                }

                if (generate && prepared.Count == 0)
                    validationWarning = "No selected rows had all required text fields. No letters were generated.";
            }
            catch (Exception e)
            {
                error = e.GetBaseException().Message;
                App.LogException(e);
            }
            finally
            {
                if (batch != null)
                {
                    Action<string, Action> write = (name, action) =>
                    {
                        try
                        {
                            action();
                        }
                        catch (Exception e)
                        {
                            failures.Add(name + ": " + e.Message);
                        }
                    };
                    write("MissingPhotos_AllRows.csv", () => Core.Report(Path.Combine(batch, "MissingPhotos_AllRows.csv"), new[] { "CSVRecord", "DataRow", "StudentName", "Scholarship", "PhotoPath", "SelectedInThisRun", "Reason" }, missing));
                    write("MergeReport.csv", () => Core.Report(Path.Combine(batch, "MergeReport.csv"), new[] { "CSVRecord", "DataRow", "StudentName", "Scholarship", "Status", "PhotoStatus", "PhotoMode", "Details", "File", "Card", "TextRepairStatus" }, selected.Select(s => new[] { s.Record.ToString(), s.DataRow.ToString(), s.Name, s.Scholarship, !prepared.Contains(s) ? "SKIPPED_DATA" : !generate ? "DATA_CHECKED" : String.IsNullOrEmpty(s.OutputFile) ? "GENERATION_FAILED" : s.PhotoAdded ? "GENERATED" : "GENERATED_NO_PHOTO", s.PhotoStatus, mode, s.PhotoDetails + " " + String.Join(" ", s.TextIssues) + (error.Length > 0 ? " RUN_ERROR: " + error : ""), s.OutputFile, String.IsNullOrEmpty(s.OutputFile) ? "" : s.Card.ToString(), s.TextIssues.Count > 0 ? "NEEDS_REVIEW" : s.TextChanges.Count == 0 ? "UNCHANGED" : s.EncodingRepairEnabled ? "REPAIRED" : "REPAIR_DISABLED" })));
                    write("TextRepairReport.csv", () => Core.Report(Path.Combine(batch, "TextRepairReport.csv"), TextRepair.Headers, TextRepair.ReportRows(selected)));
                    write("LayoutReport.csv", () => Core.Report(Path.Combine(batch, "LayoutReport.csv"), WordLayout.Headers, layout.Select(r => r.Values())));
                    if (data != null)
                        write("CSVColumns.csv", () => Core.Report(Path.Combine(batch, "CSVColumns.csv"), new[] { "CSVColumn", "OriginalHeader", "DropdownLabel", "HeaderStatus", "SelectedRole" }, data.Columns.Select((c, i) => new[] { c.Position.ToString(), c.Original, c.Label, c.Status, String.Join(";", fields.Where(f => Convert.ToString(f.Value.SelectedItem) == c.Label).Select(f => f.Key)) })));
                    string summary = "Selected " + (selection == null ? 0 : selection.Count) + " of " + (data == null ? 0 : data.Rows.Count) + " CSV rows. Eligible: " + prepared.Count + ". Generated cards: " + generated + ". Photo issues across all rows: " + missing.Count + ".\r\n" + (selection == null ? "" : "Selected data rows: " + selection.Start + "-" + selection.End + "; header and blank records excluded.\r\n") + "Photo sizing: " + mode + ". Original frame retained.\r\n";
                    summary += "Text encoding repair: " + (repairBox.Checked ? "on" : "off") + ". Selected responses repaired: " + selected.Count(s => s.EncodingRepairEnabled && s.TextChanges.Count > 0) + ". Text issues to review: " + selected.Count(s => s.TextIssues.Count > 0) + ".\r\n";
                    summary += "Separate missing-photo letters: " + (separateBox.Checked ? "on" : "off") + ".\r\n";
                    foreach (var output in merge.Outputs)
                        summary += Path.GetFileName(output.Path) + ": " + output.Cards.Count + " cards.\r\n";
                    if (mode == Core.Modes[2])
                        summary += "Overlay: " + preferences.Overlay.Width + " x " + preferences.Overlay.Height + " inches; X " + preferences.Overlay.X + ", Y " + preferences.Overlay.Y + ". Centered crop; covers curved strokes.\r\n";
                    summary += String.Join("   ", layout.GroupBy(r => r.Status).Select(g => g.Key + ": " + g.Count())) + "\r\n" + validationWarning;
                    if (error.Length > 0)
                        summary += "\r\nRUN_ERROR: " + error + ". Scanning or generation may be incomplete.";
                    if (failures.Count > 0)
                        summary += "\r\nBuild/report issues: " + String.Join(" | ", failures);
                    write("RunSummary.txt", () => File.WriteAllText(Path.Combine(batch, "RunSummary.txt"), summary, new UTF8Encoding(true)));
                    status.Text = summary + "\r\nOutput: " + batch;
                }
                else
                    status.Text = "Run could not start: " + error;
                SetBusy(false);
                Save();
                if (batch != null && Environment.OSVersion.Platform == PlatformID.Win32NT)
                {
                    try
                    {
                        OpenPath(batch);
                    }
                    catch (Exception e)
                    {
                        Notice(status.Text + "\r\nFolder could not be opened: " + e.Message);
                    }
                }
                else if (error.Length > 0)
                    Notice(status.Text);
            }
        }

        public void Render(string path)
        {
            Show();
            Application.DoEvents();
            using (var bitmap = new Bitmap(Width, Height))
            {
                DrawToBitmap(bitmap, new Rectangle(0, 0, Width, Height));
                bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }

            Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                saveTimer.Dispose();
                tips.Dispose();
            }

            base.Dispose(disposing);
        }
    }

    public sealed class PhotoPreview : Form
    {
        readonly CardTemplate template;
        readonly CsvData data;
        readonly Dictionary<string, int> map;
        readonly Selection selection;
        readonly string photos;
        readonly Action<Tuple<Overlay, string>> apply;
        readonly NumericUpDown row;
        readonly ComboBox mode;
        readonly ZoomPhotoView picture;
        readonly TrackBar zoomSlider;
        readonly Label zoomLabel;
        bool syncingZoom = false;
        readonly Label title, details;
        readonly Dictionary<string, NumericUpDown> controls = new Dictionary<string, NumericUpDown>();
        readonly Button previous, next;
        bool rendering = false;
        static Label LabelAt(Control parent, string text, int x, int y, int w, int h)
        {
            var c = new Label
            {
                Text = text,
                ForeColor = MainForm.Ink
            };
            c.SetBounds(x, y, w, h);
            parent.Controls.Add(c);
            return c;
        }

        static Button ButtonAt(Control parent, string text, int x, int y, int w, Action action)
        {
            var c = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                ForeColor = MainForm.Ink,
                BackColor = Color.White
            };
            c.SetBounds(x, y, w, 33);
            c.Click += (s, e) =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.GetBaseException().Message, "Photo preview");
                }
            };
            parent.Controls.Add(c);
            return c;
        }

        public PhotoPreview(CardTemplate t, CsvData d, Dictionary<string, int> mapping, Selection range, string folder, Overlay overlay, string photoMode, Action<Tuple<Overlay, string>> action)
        {
            template = t;
            data = d;
            map = mapping;
            selection = range;
            photos = folder;
            apply = action;
            Text = "Photo / crop preview";
            Font = new Font("Segoe UI", 10);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = ColorTranslator.FromHtml("#F8FAFC");
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96, 96);
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = MinimizeBox = true;
            ShowInTaskbar = true;
            KeyPreview = true;
            ClientSize = new Size(950, Math.Min(790, Screen.PrimaryScreen.WorkingArea.Height - 80));
            MinimumSize = new Size(790, 510);
            var shell = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(14), Margin = Padding.Empty };
            shell.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            shell.RowStyles.Add(new RowStyle(SizeType.Absolute, 320));
            shell.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(shell);
            var options = ViewLayout.Rows();
            var optionsHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Margin = Padding.Empty };
            optionsHost.Controls.Add(options);
            shell.Controls.Add(optionsHost, 0, 0);
            title = new Label { Text = "", Height = 52, AutoEllipsis = true, BackColor = MainForm.Ink, ForeColor = Color.White, Padding = new Padding(8, 5, 8, 5) };
            ViewLayout.Add(options, title, false);
            row = new NumericUpDown { Minimum = range.Start, Maximum = range.End, Value = range.Start, Width = 85 };
            previous = ButtonAt(this, "Previous", 0, 0, 100, () => { if (row.Value > row.Minimum) row.Value--; });
            next = ButtonAt(this, "Next", 0, 0, 85, () => { if (row.Value < row.Maximum) row.Value++; });
            mode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, IntegralHeight = false, MaxDropDownItems = 3, Width = 220, DropDownWidth = 220 };
            mode.Items.AddRange(Core.Modes);
            mode.SelectedItem = photoMode;
            var applyButton = ButtonAt(this, "Apply to all", 0, 0, 115, () =>
            {
                var choice = Choice();
                Core.Box(template, Convert.ToString(mode.SelectedItem), choice);
                apply(Tuple.Create(choice, Convert.ToString(mode.SelectedItem)));
                details.Text = "Applied to every photo in the next selected merge. Source files are unchanged.";
            });
            ViewLayout.Add(options, ViewLayout.Flow(ViewLayout.Text("Data row"), row, previous, next, mode, applyButton), false);
            details = ViewLayout.Text("");
            details.MinimumSize = new Size(0, 38);
            ViewLayout.Add(options, details, false);
            var limits = Core.Box(t, Core.Modes[0], overlay);
            var overlayFields = new List<Control>();
            foreach (var pair in new[] { new[] { "Width", "Width (in)" }, new[] { "Height", "Height (in)" }, new[] { "X", "Right / left (in)" }, new[] { "Y", "Down / up (in)" } })
            {
                var c = new NumericUpDown { Width = 150, DecimalPlaces = 2, Increment = .01M, Minimum = pair[0] == "Width" || pair[0] == "Height" ? .25M : -.5M, Maximum = pair[0] == "Width" ? (decimal)(Math.Floor(limits.FrameWidth / 914400 * 100) / 100) : pair[0] == "Height" ? (decimal)(Math.Floor(limits.FrameHeight / 914400 * 100) / 100) : .5M };
                double value = pair[0] == "Width" ? overlay.Width : pair[0] == "Height" ? overlay.Height : pair[0] == "X" ? overlay.X : overlay.Y;
                c.Value = Math.Min(c.Maximum, Math.Max(c.Minimum, (decimal)value));
                controls[pair[0]] = c;
                var group = ViewLayout.Rows();
                group.AutoSize = false;
                group.Size = new Size(160, 65);
                group.MinimumSize = group.MaximumSize = new Size(160, 65);
                ViewLayout.Add(group, ViewLayout.Text(pair[1]), false);
                ViewLayout.Add(group, c, false);
                overlayFields.Add(group);
            }
            ViewLayout.Add(options, ViewLayout.Flow(overlayFields.ToArray()), false);
            var reset = ButtonAt(this, "Reset overlay", 0, 0, 145, () =>
            {
                rendering = true;
                try
                {
                    var o = new Overlay();
                    foreach (var pair in controls)
                    {
                        double v = pair.Key == "Width" ? o.Width : pair.Key == "Height" ? o.Height : 0;
                        pair.Value.Value = Math.Min(pair.Value.Maximum, Math.Max(pair.Value.Minimum, (decimal)v));
                    }
                }
                finally { rendering = false; }
                UpdatePreview();
            });
            ViewLayout.Add(options, ViewLayout.Flow(reset, ViewLayout.Text("Overlay covers curve strokes. Apply to all saves batch settings.")), false);
            picture = new ZoomPhotoView { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 8) };
            shell.Controls.Add(picture, 0, 1);
            zoomLabel = ViewLayout.Text("Fit");
            zoomLabel.AutoSize = false;
            zoomLabel.Width = 100;
            zoomLabel.Height = 28;
            zoomSlider = new TrackBar { Minimum = 10, Maximum = 800, Value = 100, SmallChange = 10, LargeChange = 50, TickStyle = TickStyle.None, Width = 180, Height = 30, AutoSize = false, AccessibleName = "Preview zoom percent" };
            zoomSlider.ValueChanged += (s, e) => { if (!syncingZoom) picture.SetZoom(zoomSlider.Value / 100.0); };
            picture.ViewChanged += (s, e) => SyncZoom();
            var zoomControls = ViewLayout.Flow(ButtonAt(this, "Zoom -", 0, 0, 80, () => picture.ZoomBy(1 / 1.2)), ButtonAt(this, "Zoom +", 0, 0, 80, () => picture.ZoomBy(1.2)), zoomSlider, zoomLabel, ButtonAt(this, "Fit", 0, 0, 65, () => picture.Fit()), ButtonAt(this, "100%", 0, 0, 75, () => picture.SetZoom(1)));
            var bottom = ViewLayout.Rows();
            ViewLayout.Add(bottom, zoomControls, false);
            ViewLayout.Add(bottom, ViewLayout.Text("Wheel to zoom; drag to pan. Zoom changes only this view. Word checks the complete letter after generation."), false);
            shell.Controls.Add(bottom, 0, 2);
            bool sizingOptions = false;
            Action sizeOptions = () =>
            {
                if (sizingOptions) return;
                sizingOptions = true;
                try
                {
                    int preferred = options.GetPreferredSize(new Size(Math.Max(1, optionsHost.ClientSize.Width - SystemInformation.VerticalScrollBarWidth), 0)).Height;
                    int available = Math.Max(100, shell.ClientSize.Height - shell.Padding.Vertical - bottom.GetPreferredSize(new Size(Math.Max(1, shell.ClientSize.Width - shell.Padding.Horizontal), 0)).Height - 160);
                    float height = Math.Min(preferred, available);
                    if (Math.Abs(shell.RowStyles[0].Height - height) > 1) shell.RowStyles[0].Height = height;
                }
                finally { sizingOptions = false; }
            };
            shell.SizeChanged += (s, e) => sizeOptions();
            options.SizeChanged += (s, e) => sizeOptions();
            Shown += (s, e) => sizeOptions();
            KeyDown += (s, e) =>
            {
                if (!e.Control) return;
                if (e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add) picture.ZoomBy(1.2);
                else if (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract) picture.ZoomBy(1 / 1.2);
                else if (e.KeyCode == Keys.D0 || e.KeyCode == Keys.NumPad0) picture.SetZoom(1);
                else if (e.KeyCode == Keys.F) picture.Fit();
                else return;
                e.SuppressKeyPress = true;
            };
            row.ValueChanged += (s, e) => UpdatePreview();
            mode.SelectedIndexChanged += (s, e) => UpdatePreview();
            foreach (var c in controls.Values)
                c.ValueChanged += (s, e) => UpdatePreview();
            UpdatePreview();
        }

        Overlay Choice()
        {
            return new Overlay
            {
                Width = (double)controls["Width"].Value,
                Height = (double)controls["Height"].Value,
                X = (double)controls["X"].Value,
                Y = (double)controls["Y"].Value
            };
        }

        void UpdatePreview()
        {
            if (rendering)
                return;
            rendering = true;
            try
            {
                foreach (var c in controls.Values)
                    c.Enabled = Convert.ToString(mode.SelectedItem) == Core.Modes[2];
                int index = (int)row.Value;
                var record = data.Rows[index - 1];
                string name = record.Values[map["Name"]].Trim(), scholar = record.Values[map["Scholarship"]].Trim();
                title.Text = name + "\r\n" + scholar;
                previous.Enabled = index > selection.Start;
                next.Enabled = index < selection.End;
                string path = Core.PhotoPath(photos, record.Values[map["Photo"]]);
                foreach (var other in data.Rows)
                {
                    try
                    {
                        string otherPath = Core.PhotoPath(photos, other.Values[map["Photo"]]);
                        if (otherPath.Equals(path, StringComparison.OrdinalIgnoreCase) && !Core.PhotoOwner(other.Values[map["Name"]]).Equals(Core.PhotoOwner(name), StringComparison.OrdinalIgnoreCase))
                            throw new InvalidDataException("PHOTO_CONFLICT: path belongs to different student names.");
                    }
                    catch (InvalidDataException e)
                    {
                        if (e.Message.StartsWith("PHOTO_CONFLICT"))
                            throw;
                    }
                }

                path = Core.ResolvePhoto(photos, record.Values[map["Photo"]]);
                var box = Core.Box(template, Convert.ToString(mode.SelectedItem), Choice());
                var photo = Core.ReadPhoto(path, box, Convert.ToString(mode.SelectedItem));
                picture.Image = Core.Preview(template, photo);
                details.Text = "Data row " + index + "; CSV record " + record.Record + ". " + mode.SelectedItem + ". Cropped width " + photo.Plan.HorizontalCrop + "%; height " + photo.Plan.VerticalCrop + "%.";
                if (photo.Frames > 1)
                    details.Text += " First frame/page only.";
            }
            catch (Exception e)
            {
                picture.Image = null;
                details.Text = "Preview issue: " + e.GetBaseException().Message;
            }
            finally
            {
                rendering = false;
            }
        }

        void SyncZoom()
        {
            syncingZoom = true;
            try
            {
                zoomSlider.Value = Math.Max(zoomSlider.Minimum, Math.Min(zoomSlider.Maximum, (int)Math.Round(picture.State.Zoom * 100)));
                zoomLabel.Text = (picture.State.FitMode ? "Fit " : "") + (picture.State.Zoom * 100).ToString("0", CultureInfo.InvariantCulture) + "%";
            }
            finally { syncingZoom = false; }
        }
    }
}
