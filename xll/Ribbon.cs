using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace MultiFilter
{
    [ComVisible(true)]
    public class MultiFilterRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab idMso='TabData'>
        <group id='mfGroup' label='Multi-filter' insertAfterMso='GroupSortFilter'>
          <button id='mfOpen' label='Multi-filter' size='large' getImage='GetImage' onAction='OnOpen'
                  screentip='Multi-filter' supertip='Filtrér kolonnen med den aktive celle på en liste af værdier, som du indsætter.'/>
          <button id='mfClearCol' label='Ryd kolonne' getImage='GetImage' onAction='OnClearColumn'
                  supertip='Fjerner filteret på kolonnen med den aktive celle.'/>
          <button id='mfClearAll' label='Ryd alle' getImage='GetImage' onAction='OnClearAll'
                  supertip='Fjerner alle filtre på arket (filterpilene bliver).'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso='ContextMenuCell'>
      <button id='mfCtxCell' label='Multi-filter...' getImage='GetImage' onAction='OnOpen' insertBeforeMso='Cut'/>
      <menuSeparator id='mfCtxCellSep' insertBeforeMso='Cut'/>
    </contextMenu>
    <contextMenu idMso='ContextMenuListRange'>
      <button id='mfCtxList' label='Multi-filter...' getImage='GetImage' onAction='OnOpen' insertBeforeMso='Cut'/>
      <menuSeparator id='mfCtxListSep' insertBeforeMso='Cut'/>
    </contextMenu>
  </contextMenus>
</customUI>";

        static Excel.Application App => (Excel.Application)ExcelDnaUtil.Application;

        public Bitmap GetImage(IRibbonControl control) =>
            control.Id == "mfOpen" ? Icons.Funnel(32, false)
            : control.Id.StartsWith("mfCtx") ? Icons.Funnel(16, false)
            : Icons.Funnel(16, true);

        public void OnOpen(IRibbonControl control) => OpenDialog();

        internal static void OpenDialog()
        {
            try
            {
                using (var f = new FilterForm(App))
                    f.ShowDialog(new ExcelWindow(App.Hwnd));
            }
            catch (Exception ex) { Error(ex); }
        }

        public void OnClearColumn(IRibbonControl control)
        {
            try { FilterEngine.ClearColumn(App); }
            catch (Exception ex) { Error(ex); }
        }

        public void OnClearAll(IRibbonControl control)
        {
            try { FilterEngine.ClearAll(App); }
            catch (Exception ex) { Error(ex); }
        }

        static void Error(Exception ex) =>
            MessageBox.Show(new ExcelWindow(App.Hwnd), ex is FilterException ? ex.Message : "Fejl: " + ex.Message,
                            "Multi-filter", MessageBoxButtons.OK, ex is FilterException ? MessageBoxIcon.Information : MessageBoxIcon.Error);
    }

    public static class Commands
    {
        /// <summary>Ctrl+Shift+M åbner Multi-filter.</summary>
        [ExcelCommand(Name = "MultiFilter.Open", ShortCut = "^+M")]
        public static void Open() => MultiFilterRibbon.OpenDialog();
    }

    internal sealed class ExcelWindow : IWin32Window
    {
        public ExcelWindow(int hwnd) { Handle = new IntPtr(hwnd); }
        public IntPtr Handle { get; }
    }

    internal static class Icons
    {
        static readonly Color Green = Color.FromArgb(33, 115, 70);

        /// <summary>Tragt-ikon; med rødt kryds til "ryd"-knapperne.</summary>
        public static Bitmap Funnel(int size, bool cross)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                float k = size / 32f;
                var pts = new[]
                {
                    new PointF(3 * k, 5 * k), new PointF(29 * k, 5 * k), new PointF(19 * k, 16 * k),
                    new PointF(19 * k, 27 * k), new PointF(13 * k, 30 * k), new PointF(13 * k, 16 * k),
                };
                using (var b = new SolidBrush(Green)) g.FillPolygon(b, pts);
                if (!cross)
                {
                    using (var p = new Pen(Color.White, Math.Max(1f, 1.6f * k)))
                    {
                        g.DrawLine(p, 9 * k, 8 * k, 23 * k, 8 * k);
                        g.DrawLine(p, 11 * k, 11 * k, 21 * k, 11 * k);
                    }
                }
                else
                {
                    using (var p = new Pen(Color.FromArgb(200, 40, 40), Math.Max(2f, 4f * k)) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                    {
                        g.DrawLine(p, 19 * k, 18 * k, 30 * k, 29 * k);
                        g.DrawLine(p, 30 * k, 18 * k, 19 * k, 29 * k);
                    }
                }
            }
            return bmp;
        }

        public static Icon FormIcon()
        {
            using (var bmp = Funnel(32, false))
                return Icon.FromHandle(bmp.GetHicon());
        }
    }
}
