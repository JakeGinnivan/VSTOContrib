using System;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace Excel.TestDoubles
{
    public class WindowsTestDouble : Windows
    {
        readonly List<WindowTestDouble> windows = new List<WindowTestDouble>();

        public WindowsTestDouble(ApplicationTestDouble applicationTestDouble)
        {
            Application = applicationTestDouble;
        }

        public object Arrange(XlArrangeStyle arrangeStyle, object activeWorkbook, object syncHorizontal,
            object syncVertical)
        {
            throw new NotImplementedException();
        }

        IEnumerator Windows.GetEnumerator()
        {
            return windows.GetEnumerator();
        }

        public bool CompareSideBySideWith(object windowName)
        {
            throw new NotImplementedException();
        }

        public bool BreakSideBySide()
        {
            throw new NotImplementedException();
        }

        public void ResetPositionsSideBySide()
        {
            throw new NotImplementedException();
        }

        public Application Application { get; private set; }
        public XlCreator Creator { get; private set; }
        public object Parent { get; private set; }
        public int Count { get { return windows.Count; } }

        Window Windows.get_Item(object index)
        {
            return windows[((int)index) - 1];
        }

        Window Windows.this[object index]
        {
            get { return windows[((int)index) - 1]; }
        }

        public bool SyncScrollingSideBySide { get; set; }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return windows.GetEnumerator();
        }
    }
}