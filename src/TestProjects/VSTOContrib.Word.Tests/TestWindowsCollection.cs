using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace VSTOContrib.Word.Tests
{
    public class TestWindowsCollection : List<Window>, Windows
    {
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Window Add(ref object Window)
        {
            var window = (Window) Window;
            Add(window);
            return window;
        }

        public void Arrange(ref object ArrangeStyle)
        {
            throw new System.NotImplementedException();
        }

        public bool CompareSideBySideWith(ref object Document)
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

        public Application Application { get; private set; }
        public int Creator { get; private set; }
        public object Parent { get; private set; }
        public Window get_Item(ref object Index)
        {
            throw new System.NotImplementedException();
        }

        public bool SyncScrollingSideBySide { get; set; }

        IEnumerator Windows.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}