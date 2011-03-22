using System.Collections.Generic;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Contrib.RibbonFactory.Interfaces.Internal;

namespace Office.Contrib.RibbonFactory.Internal
{
    internal class CustomTaskPaneRegister : ICustomTaskPaneRegister
    {
        private CustomTaskPaneCollection _customTaskPaneCollection;
        private readonly Dictionary<object, Queue<CustomTaskPane>> _taskPanesToCleanup;

        public CustomTaskPaneRegister()
        {
            _taskPanesToCleanup = new Dictionary<object, Queue<CustomTaskPane>>();
        }

        public void Initialise(CustomTaskPaneCollection customTaskPaneCollection)
        {
            _customTaskPaneCollection = customTaskPaneCollection;
        }

        public void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object context)
        {
            var registersCustomTaskPanes = ribbonViewModel as IRegisterCustomTaskPane;
            if (registersCustomTaskPanes != null)
            {
                registersCustomTaskPanes.RegisterTaskPanes(
                    (control, title) =>
                    {
                        var taskPane = _customTaskPaneCollection.Add(control, title, context);
                        if (!_taskPanesToCleanup.ContainsKey(context))
                            _taskPanesToCleanup.Add(context, new Queue<CustomTaskPane>());

                        _taskPanesToCleanup[context].Enqueue(taskPane);
                        return taskPane;
                    });
            }
        }

        public void Cleanup(object context)
        {
            if (!_taskPanesToCleanup.ContainsKey(context)) return;

            while (_taskPanesToCleanup[context].Count > 0)
            {
                _taskPanesToCleanup[context].Dequeue().Dispose();
            }
        }
    }
}
