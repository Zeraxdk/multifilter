using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MultiFilter
{
    internal sealed class FilterForm : Form
    {
        static string _lastText = "";
        static bool _lastContains;

        static readonly Color Green = Color.FromArgb(33, 115, 70);
        static readonly Color WarnFg = Color.FromArgb(138, 90, 0);
        static readonly Color ErrFg = Color.FromArgb(164, 38, 44);

        readonly Excel.Application _app;
        readonly TextBox _values;
        readonly Label _column, _count, _status;
        readonly CheckBox _contains;
        readonly Button _apply, _close;

        public FilterForm(Excel.Application app)
        {
            _app = app;

            Text = "Multi-filter";
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(380, 470);
            MinimumSize = new Size(320, 380);
            Icon = Icons.FormIcon();
            KeyPreview = true;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 7,
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // kolonne
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // label
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // tekstboks
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // antal + indeholder
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // status
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // knapper
            Controls.Add(layout);

            _column = new Label
            {
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Green,
                Margin = new Padding(0, 0, 0, 10),
                Text = "Kolonne: –",
            };
            layout.Controls.Add(_column);

            layout.Controls.Add(new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4),
                Text = "Indsæt værdier – én per linje (eller adskilt af ; eller tab):",
            });

            _values = new TextBox
            {
                Multiline = true,
                AcceptsReturn = true,
                AcceptsTab = false,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10f),
                Text = _lastText,
            };
            _values.TextChanged += (s, e) => UpdateCount();
            layout.Controls.Add(_values);

            var optionsRow = new TableLayoutPanel { ColumnCount = 2, Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 6, 0, 0) };
            optionsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            optionsRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _contains = new CheckBox { Text = "Indeholder (delvis match)", AutoSize = true, Checked = _lastContains, Margin = new Padding(0) };
            _count = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Right, Margin = new Padding(0, 2, 0, 0) };
            optionsRow.Controls.Add(_contains, 0, 0);
            optionsRow.Controls.Add(_count, 1, 0);
            layout.Controls.Add(optionsRow);

            _status = new Label { AutoSize = true, MaximumSize = new Size(356, 0), Margin = new Padding(0, 8, 0, 0), Visible = false };
            layout.Controls.Add(_status);

            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 10, 0, 0) };
            _close = new Button { Text = "Luk", AutoSize = true, MinimumSize = new Size(80, 28), DialogResult = DialogResult.Cancel };
            _apply = new Button
            {
                Text = "Filtrér",
                AutoSize = true,
                MinimumSize = new Size(100, 28),
                BackColor = Green,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(Font, FontStyle.Bold),
            };
            _apply.FlatAppearance.BorderColor = Green;
            _apply.Click += (s, e) => DoFilter();
            var clearCol = new Button { Text = "Ryd kolonne", AutoSize = true, MinimumSize = new Size(90, 28) };
            clearCol.Click += (s, e) => DoClearColumn();
            buttons.Controls.Add(_close);
            buttons.Controls.Add(_apply);
            buttons.Controls.Add(clearCol);
            layout.Controls.Add(buttons);

            CancelButton = _close;
            KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; DoFilter(); }
            };

            var header = FilterEngine.PeekHeader(_app);
            _column.Text = header == null ? "Kolonne: (klik på en celle i dine data)" : "Kolonne: " + header;
            UpdateCount();
            FormClosing += (s, e) => { _lastText = _values.Text; _lastContains = _contains.Checked; };
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _values.Focus();
            _values.SelectAll();
        }

        void UpdateCount()
        {
            int n = FilterEngine.ParseTerms(_values.Text).Count;
            _count.Text = n == 1 ? "1 værdi" : n + " værdier";
        }

        void ShowStatus(string text, Color color)
        {
            _status.Text = text;
            _status.ForeColor = color;
            _status.Visible = true;
        }

        void DoFilter()
        {
            var terms = FilterEngine.ParseTerms(_values.Text);
            try
            {
                Cursor = Cursors.WaitCursor;
                var r = FilterEngine.Apply(_app, terms, _contains.Checked);
                if (r.Missing.Count == 0)
                {
                    _lastText = _values.Text;
                    _lastContains = _contains.Checked;
                    DialogResult = DialogResult.OK;   // alt fundet -> luk
                    Close();
                    return;
                }
                var list = string.Join(", ", r.Missing.Take(30));
                if (r.Missing.Count > 30) list += $" … (+{r.Missing.Count - 30})";
                ShowStatus($"Filtreret på {r.TermCount - r.Missing.Count} af {r.TermCount} værdier i \"{r.Header}\".\nIkke fundet: {list}", WarnFg);
            }
            catch (FilterException ex) { ShowStatus(ex.Message, ErrFg); }
            catch (Exception ex) { ShowStatus("Fejl: " + ex.Message, ErrFg); }
            finally { Cursor = Cursors.Default; }
        }

        void DoClearColumn()
        {
            try
            {
                var h = FilterEngine.ClearColumn(_app);
                ShowStatus($"Filteret på \"{h}\" er fjernet.", Green);
            }
            catch (FilterException ex) { ShowStatus(ex.Message, ErrFg); }
            catch (Exception ex) { ShowStatus("Fejl: " + ex.Message, ErrFg); }
        }
    }
}
