using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace Word.TestDoubles
{
    public class WindowsTestDouble : Windows
    {
        readonly List<WindowTestDouble> windows = new List<WindowTestDouble>();

        public WindowsTestDouble(ApplicationTestDouble applicationTestDouble)
        {
            Application = applicationTestDouble;
        }

        IEnumerator Windows.GetEnumerator()
        {
            return windows.GetEnumerator();
        }

        public Window Add(ref object window)
        {
            var windowTestDouble = (WindowTestDouble) window;
            windows.Add(windowTestDouble);
            return windowTestDouble;
        }

        public void Arrange(ref object arrangeStyle)
        {
            throw new System.NotImplementedException();
        }

        public bool CompareSideBySideWith(ref object document)
        {
            throw new System.NotImplementedException();
        }

        public bool BreakSideBySide()
        {
            throw new System.NotImplementedException();
        }

        public void ResetPositionsSideBySide()
        {
            throw new System.NotImplementedException();
        }

        public int Count { get { return windows.Count; } }

        public Application Application { get; private set; }
        public int Creator { get; private set; }
        public object Parent { get; private set; }

        public Window get_Item(ref object index)
        {
            return windows[((int) index) - 1];
        }

        public bool SyncScrollingSideBySide { get; set; }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return windows.GetEnumerator();
        }
    }
}