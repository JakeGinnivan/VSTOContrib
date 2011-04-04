using System.Collections.Generic;
using Microsoft.Office.Tools;
using Office.Contrib.RibbonFactory.Interfaces;
using Office.Contrib.RibbonFactory.Interfaces.Internal;

namespace Office.Contrib.RibbonFactory.Internal
{
    internal class CustomTaskPaneRegister : ICustomTaskPaneRegister
    {
        private CustomTaskPaneCollection _customTaskPaneCollection;
        private readonly Dictionary<IRibbonViewModel, TaskPaneRegistrationInfo> _registrationInfo;
        private readonly Dictionary<IRibbonViewModel, OneToManyCustomTaskPaneAdapter> _ribbonTaskPanes;

        public CustomTaskPaneRegister()
        {
            _registrationInfo = new Dictionary<IRibbonViewModel, TaskPaneRegistrationInfo>();
            _ribbonTaskPanes = new Dictionary<IRibbonViewModel, OneToManyCustomTaskPaneAdapter>();
        }

        public void Initialise(CustomTaskPaneCollection customTaskPaneCollection)
        {
            _customTaskPaneCollection = customTaskPaneCollection;
        }

        public void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view)
        {
            var registersCustomTaskPanes = ribbonViewModel as IRegisterCustomTaskPane;
            if (registersCustomTaskPanes == null) return;

            if (!_registrationInfo.ContainsKey(ribbonViewModel))
            {
                registersCustomTaskPanes.RegisterTaskPanes(
                    (control, title) =>
                        {
                            var taskPaneRegistrationInfo = new TaskPaneRegistrationInfo(control, title);
                            _registrationInfo.Add(ribbonViewModel, taskPaneRegistrationInfo);
                            var taskPane = Register(view, taskPaneRegistrationInfo);
                            var taskPaneAdapter = new OneToManyCustomTaskPaneAdapter(taskPane);
                            _ribbonTaskPanes.Add(ribbonViewModel, taskPaneAdapter);
                            return taskPaneAdapter;
                        });
            }
            else
            {
                var adapter = _ribbonTaskPanes[ribbonViewModel];
                if (!adapter.ViewRegistered(view))
                {
                    adapter.Add(Register(view, _registrationInfo[ribbonViewModel]));
                }
            }
        }

        private CustomTaskPane Register(object view, TaskPaneRegistrationInfo taskPaneRegistrationInfo)
        {
            var taskPane = _customTaskPaneCollection.Add(taskPaneRegistrationInfo.Control, taskPaneRegistrationInfo.Title, view);

            return taskPane;
        }

        public void Cleanup(object view)
        {
            foreach (var adapter in _ribbonTaskPanes)
            {
                adapter.Value.CleanupView(view);
            }
        }
    }
}
