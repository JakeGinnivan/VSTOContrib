using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using VSTOContrib.Core.Tests.RibbonFactory.Internal;

namespace VSTOContrib.Core.Tests.RibbonFactory.TestAddin
{
    public class CustomTaskPaneCollectionDouble : CustomTaskPaneCollection
    {
        readonly List<CustomTaskPane> customTaskPanes = new List<CustomTaskPane>();

        public IEnumerator<CustomTaskPane> GetEnumerator()
        {
            return customTaskPanes.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Dispose()
        {
            foreach (var customTaskPane in customTaskPanes)
            {
                customTaskPane.Dispose();
            }
            customTaskPanes.Clear();
        }

        public void BeginInit()
        {
        }

        public void EndInit()
        {
        }

        public CustomTaskPane Add(UserControl control, string title)
        {
            var customTaskPaneDouble = new CustomTaskPaneDouble(title);
            customTaskPanes.Add(customTaskPaneDouble);
            return customTaskPaneDouble;
        }

        public CustomTaskPane Add(UserControl control, string title, object window)
        {
            var customTaskPaneDouble = new CustomTaskPaneDouble(title, window);
            customTaskPanes.Add(customTaskPaneDouble);
            return customTaskPaneDouble;
        }

        public bool Remove(CustomTaskPane customTaskPane)
        {
            return customTaskPanes.Remove(customTaskPane);
        }

        public void RemoveAt(int index)
        {
            throw new NotImplementedException();
        }

        public int Count { get { return customTaskPanes.Count; }}

        public CustomTaskPane this[int index]
        {
            get { return customTaskPanes[index]; }
        }
    }
}