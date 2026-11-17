using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace MultiFilter
{
    /// <summary>Et filterområde og hvilken kolonne (1-baseret) der filtreres i.</summary>
    internal sealed class FilterTarget
    {
        public Excel.Range Range;
        public int Field;
        public string Header;
    }

    internal sealed class FilterResult
    {
        public int TermCount;
        public List<string> Missing = new List<string>();
        public int MatchedValues;
        public string Header;
    }

    internal sealed class FilterException : Exception
    {
        public FilterException(string message) : base(message) { }
    }

    internal static class FilterEngine
    {
        // Linjeskift, tab og semikolon. Ikke komma: det er decimaltegn i Danmark.
        static readonly Regex Separators = new Regex(@"[\r\n\t;]+", RegexOptions.Compiled);

        public static List<string> ParseTerms(string text)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in Separators.Split(text ?? ""))
            {
                var t = raw.Replace(' ', ' ').Trim();
                if (t.Length > 0 && seen.Add(t)) result.Add(t);
            }
            return result;
        }

        /// <summary>Finder filterområdet for den aktive celle (tabel, eksisterende filter eller dataområdet rundt om cellen).</summary>
        public static FilterTarget GetTarget(Excel.Application app, bool createIfMissing)
        {
            if (!(app.ActiveSheet is Excel.Worksheet ws))
                throw new FilterException("Vælg en celle i et regneark.");
            Excel.Range cell = app.ActiveCell;
            if (cell == null)
                throw new FilterException("Vælg en celle i den kolonne du vil filtrere.");

            Excel.ListObject lo = null;
            try { lo = cell.ListObject; } catch { }

            Excel.Range rng;
            if (lo != null)
            {
                if (!lo.ShowAutoFilter)
                {
                    if (!createIfMissing) throw new FilterException("Tabellen har ikke noget filter.");
                    lo.ShowAutoFilter = true;
                }
                rng = lo.Range;
            }
            else if (ws.AutoFilterMode)
            {
                rng = ws.AutoFilter.Range;
            }
            else
            {
                if (!createIfMissing) throw new FilterException("Der er intet filter på arket.");
                rng = cell.CurrentRegion;
                if (rng.Rows.Count < 2)
                    throw new FilterException("Klik på en celle inde i dine data (med overskrifter i første række).");
            }

            int firstCol = rng.Column;
            int lastCol = firstCol + rng.Columns.Count - 1;
            if (cell.Column < firstCol || cell.Column > lastCol)
                throw new FilterException($"Den aktive celle er uden for filterområdet ({rng.Address[false, false]}).\nKlik på en celle i den kolonne du vil filtrere.");

            int field = cell.Column - firstCol + 1;
            string header = Convert.ToString(((Excel.Range)rng.Cells[1, field]).Text);
            return new FilterTarget { Range = rng, Field = field, Header = string.IsNullOrWhiteSpace(header) ? "(uden overskrift)" : header };
        }

        public static FilterResult Apply(Excel.Application app, IList<string> terms, bool containsMode)
        {
            if (terms.Count == 0) throw new FilterException("Indsæt nogle værdier først.");

            var target = GetTarget(app, true);
            var rng = target.Range;
            int rows = rng.Rows.Count;
            if (rows < 2) throw new FilterException("Der er ingen data under overskriften.");

            Excel.Range column = (Excel.Range)rng.Columns[target.Field];
            Excel.Range data = column.Offset[1, 0].Resize[rows - 1, 1];
            object raw = data.Value2;
            object[,] values = raw as object[,];
            if (values == null)
            {
                values = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
                values[1, 1] = raw;
            }

            var culture = CultureInfo.CurrentCulture;
            var lowerTerms = terms.Select(t => t.ToLower(culture)).ToArray();
            var termIndex = new Dictionary<string, int>();
            var numIndex = new Dictionary<double, int>();
            for (int j = 0; j < lowerTerms.Length; j++)
            {
                termIndex[lowerTerms[j]] = j;
                if (double.TryParse(terms[j], NumberStyles.Float, culture, out var d))
                    numIndex[Math.Round(d, 9)] = j;
            }

            var found = new bool[terms.Count];
            var checkedKeys = new Dictionary<string, int>();   // nøgle -> række (eller -1 = ikke match)
            int lo1 = values.GetLowerBound(0), hi1 = values.GetUpperBound(0), c1 = values.GetLowerBound(1);

            for (int i = lo1; i <= hi1; i++)
            {
                object v = values[i, c1];
                if (v == null || v is int) continue;   // tom celle eller fejlværdi (#N/A osv.)

                string s = v is double dv ? dv.ToString(culture) : Convert.ToString(v, culture);
                s = s.Replace(' ', ' ').Trim().ToLower(culture);
                if (s.Length == 0 || checkedKeys.ContainsKey(s)) continue;

                bool match = false;
                if (containsMode)
                {
                    for (int j = 0; j < lowerTerms.Length; j++)
                        if (s.Contains(lowerTerms[j])) { found[j] = true; match = true; }
                }
                else
                {
                    if (termIndex.TryGetValue(s, out int j)) { found[j] = true; match = true; }
                    if (v is double d && numIndex.TryGetValue(Math.Round(d, 9), out int k)) { found[k] = true; match = true; }
                }
                checkedKeys[s] = match ? i : -1;
            }

            // Excel filtrerer på den viste tekst, så hent .Text for én celle per matchende værdi.
            var hits = new List<object>();
            var hitSet = new HashSet<string>();
            foreach (var row in checkedKeys.Values.Where(r => r > 0))
            {
                var c = (Excel.Range)data.Cells[row - lo1 + 1, 1];
                string text = DisplayText(app, c);
                if (hitSet.Add(text)) hits.Add(text);
            }

            var result = new FilterResult { TermCount = terms.Count, MatchedValues = hits.Count, Header = target.Header };
            for (int j = 0; j < terms.Count; j++)
                if (!found[j]) result.Missing.Add(terms[j]);

            if (hits.Count == 0)
                throw new FilterException($"Ingen af de {terms.Count} værdier blev fundet i kolonnen \"{target.Header}\".");

            rng.AutoFilter(target.Field, hits.ToArray(), Excel.XlAutoFilterOperator.xlFilterValues);
            return result;
        }

        static string DisplayText(Excel.Application app, Excel.Range c)
        {
            string t = Convert.ToString(c.Text);
            if (t.Length > 0 && t.Trim('#').Length == 0)   // ### i for smal kolonne
            {
                try { t = app.WorksheetFunction.Text(c.Value2, Convert.ToString(c.NumberFormat)); } catch { }
            }
            return t;
        }

        public static string ClearColumn(Excel.Application app)
        {
            var target = GetTarget(app, false);
            target.Range.AutoFilter(target.Field);
            return target.Header;
        }

        public static void ClearAll(Excel.Application app)
        {
            if (!(app.ActiveSheet is Excel.Worksheet ws)) return;
            try { if (ws.FilterMode) ws.ShowAllData(); } catch { }
            foreach (Excel.ListObject lo in ws.ListObjects)
            {
                try { if (lo.ShowAutoFilter && lo.AutoFilter.FilterMode) lo.AutoFilter.ShowAllData(); } catch { }
            }
        }

        /// <summary>Overskriften på den kolonne der vil blive filtreret, eller null.</summary>
        public static string PeekHeader(Excel.Application app)
        {
            try { return GetTarget(app, true).Header; } catch { return null; }
        }
    }
}
