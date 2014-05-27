using System;
using Microsoft.Office.Interop.Excel;

namespace Excel.TestDoubles
{
    public class WindowTestDouble : Window
    {
        public object Activate()
        {
            throw new NotImplementedException();
        }

        public object ActivateNext()
        {
            throw new NotImplementedException();
        }

        public object ActivatePrevious()
        {
            throw new NotImplementedException();
        }

        public bool Close(object SaveChanges, object Filename, object RouteWorkbook)
        {
            throw new NotImplementedException();
        }

        public object LargeScroll(object Down, object Up, object ToRight, object ToLeft)
        {
            throw new NotImplementedException();
        }

        public Window NewWindow()
        {
            throw new NotImplementedException();
        }

        public object _PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile,
            object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }

        public object PrintPreview(object EnableChanges)
        {
            throw new NotImplementedException();
        }

        public object ScrollWorkbookTabs(object Sheets, object Position)
        {
            throw new NotImplementedException();
        }

        public object SmallScroll(object Down, object Up, object ToRight, object ToLeft)
        {
            throw new NotImplementedException();
        }

        public int PointsToScreenPixelsX(int Points)
        {
            throw new NotImplementedException();
        }

        public int PointsToScreenPixelsY(int Points)
        {
            throw new NotImplementedException();
        }

        public object RangeFromPoint(int x, int y)
        {
            throw new NotImplementedException();
        }

        public void ScrollIntoView(int Left, int Top, int Width, int Height, object Start)
        {
            throw new NotImplementedException();
        }

        public object PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile,
            object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }

        public Application Application { get; private set; }
        public XlCreator Creator { get; private set; }
        public object Parent { get; private set; }
        public Range ActiveCell { get; private set; }
        public Chart ActiveChart { get; private set; }
        public Pane ActivePane { get; private set; }
        public object ActiveSheet { get; private set; }
        public object Caption { get; set; }
        public bool DisplayFormulas { get; set; }
        public bool DisplayGridlines { get; set; }
        public bool DisplayHeadings { get; set; }
        public bool DisplayHorizontalScrollBar { get; set; }
        public bool DisplayOutline { get; set; }
        public bool _DisplayRightToLeft { get; set; }
        public bool DisplayVerticalScrollBar { get; set; }
        public bool DisplayWorkbookTabs { get; set; }
        public bool DisplayZeros { get; set; }
        public bool EnableResize { get; set; }
        public bool FreezePanes { get; set; }
        public int GridlineColor { get; set; }
        public XlColorIndex GridlineColorIndex { get; set; }
        public double Height { get; set; }
        public int Index { get; private set; }
        public double Left { get; set; }
        public string OnWindow { get; set; }
        public Panes Panes { get; private set; }
        public Range RangeSelection { get; private set; }
        public int ScrollColumn { get; set; }
        public int ScrollRow { get; set; }
        public Sheets SelectedSheets { get; private set; }
        public object Selection { get; private set; }
        public bool Split { get; set; }
        public int SplitColumn { get; set; }
        public double SplitHorizontal { get; set; }
        public int SplitRow { get; set; }
        public double SplitVertical { get; set; }
        public double TabRatio { get; set; }
        public double Top { get; set; }
        public XlWindowType Type { get; private set; }
        public double UsableHeight { get; private set; }
        public double UsableWidth { get; private set; }
        public bool Visible { get; set; }
        public Range VisibleRange { get; private set; }
        public double Width { get; set; }
        public int WindowNumber { get; private set; }
        public XlWindowState WindowState { get; set; }
        public object Zoom { get; set; }
        public XlWindowView View { get; set; }
        public bool DisplayRightToLeft { get; set; }
        public SheetViews SheetViews { get; private set; }
        public object ActiveSheetView { get; private set; }
        public bool DisplayRuler { get; set; }
        public bool AutoFilterDateGrouping { get; set; }
        public bool DisplayWhitespace { get; set; }
        public int Hwnd { get; private set; }
    }
}