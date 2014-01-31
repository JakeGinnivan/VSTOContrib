using Microsoft.Office.Interop.Word;

namespace VSTOContrib.Word.Tests
{
    public class TestWindow : Window
    {
        public void Activate()
        {
            throw new System.NotImplementedException();
        }

        public void Close(ref object SaveChanges, ref object RouteDocument)
        {
            throw new System.NotImplementedException();
        }

        public void LargeScroll(ref object Down, ref object Up, ref object ToRight, ref object ToLeft)
        {
            throw new System.NotImplementedException();
        }

        public void SmallScroll(ref object Down, ref object Up, ref object ToRight, ref object ToLeft)
        {
            throw new System.NotImplementedException();
        }

        public Window NewWindow()
        {
            throw new System.NotImplementedException();
        }

        public void PrintOutOld(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object ActivePrinterMacGX, ref object ManualDuplexPrint)
        {
            throw new System.NotImplementedException();
        }

        public void PageScroll(ref object Down, ref object Up)
        {
            throw new System.NotImplementedException();
        }

        public void SetFocus()
        {
            throw new System.NotImplementedException();
        }

        public object RangeFromPoint(int x, int y)
        {
            throw new System.NotImplementedException();
        }

        public void ScrollIntoView(object obj, ref object Start)
        {
            throw new System.NotImplementedException();
        }

        public void GetPoint(out int ScreenPixelsLeft, out int ScreenPixelsTop, out int ScreenPixelsWidth, out int ScreenPixelsHeight,
            object obj)
        {
            throw new System.NotImplementedException();
        }

        public void PrintOut2000(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object ActivePrinterMacGX, ref object ManualDuplexPrint, ref object PrintZoomColumn,
            ref object PrintZoomRow, ref object PrintZoomPaperWidth, ref object PrintZoomPaperHeight)
        {
            throw new System.NotImplementedException();
        }

        public void PrintOut(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object ActivePrinterMacGX, ref object ManualDuplexPrint, ref object PrintZoomColumn,
            ref object PrintZoomRow, ref object PrintZoomPaperWidth, ref object PrintZoomPaperHeight)
        {
            throw new System.NotImplementedException();
        }

        public void ToggleShowAllReviewers()
        {
            throw new System.NotImplementedException();
        }

        public void ToggleRibbon()
        {
            throw new System.NotImplementedException();
        }

        public Application Application { get; private set; }
        public int Creator { get; private set; }
        public object Parent { get; private set; }
        public Pane ActivePane { get; private set; }
        public Document Document { get; private set; }
        public Panes Panes { get; private set; }
        public Selection Selection { get; private set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Split { get; set; }
        public int SplitVertical { get; set; }
        public string Caption { get; set; }
        public WdWindowState WindowState { get; set; }
        public bool DisplayRulers { get; set; }
        public bool DisplayVerticalRuler { get; set; }
        public View View { get; private set; }
        public WdWindowType Type { get; private set; }
        public Window Next { get; private set; }
        public Window Previous { get; private set; }
        public int WindowNumber { get; private set; }
        public bool DisplayVerticalScrollBar { get; set; }
        public bool DisplayHorizontalScrollBar { get; set; }
        public float StyleAreaWidth { get; set; }
        public bool DisplayScreenTips { get; set; }
        public int HorizontalPercentScrolled { get; set; }
        public int VerticalPercentScrolled { get; set; }
        public bool DocumentMap { get; set; }
        public bool Active { get; private set; }
        public int DocumentMapPercentWidth { get; set; }
        public int Index { get; private set; }
        public WdIMEMode IMEMode { get; set; }
        public int UsableWidth { get; private set; }
        public int UsableHeight { get; private set; }
        public bool EnvelopeVisible { get; set; }
        public bool DisplayRightRuler { get; set; }
        public bool DisplayLeftScrollBar { get; set; }
        public bool Visible { get; set; }
        public bool Thumbnails { get; set; }
        public WdShowSourceDocuments ShowSourceDocuments { get; set; }
        public int Hwnd { get; private set; }
    }
}