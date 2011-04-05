using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Interfaces.Internal;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class CustomTaskPaneRegister : ICustomTaskPaneRegister
    {
        private CustomTaskPaneCollection _customTaskPaneCollection;
        private readonly Dictionary<IRibbonViewModel, List<TaskPaneRegistrationInfo>> _registrationInfo;
        private readonly Dictionary<IRibbonViewModel, List<OneToManyCustomTaskPaneAdapter>> _ribbonTaskPanes;

        public CustomTaskPaneRegister()
        {
            _registrationInfo = new Dictionary<IRibbonViewModel, List<TaskPaneRegistrationInfo>>();
            _ribbonTaskPanes = new Dictionary<IRibbonViewModel, List<OneToManyCustomTaskPaneAdapter>>();
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
                    (controlFactory, title) =>
                        {
                            var taskPaneRegistrationInfo = new TaskPaneRegistrationInfo(controlFactory, title);
                            if (!_registrationInfo.ContainsKey(ribbonViewModel))
                                _registrationInfo.Add(ribbonViewModel, new List<TaskPaneRegistrationInfo>());
                            _registrationInfo[ribbonViewModel].Add(taskPaneRegistrationInfo);

                            var taskPane = Register(view, taskPaneRegistrationInfo);
                            var taskPaneAdapter = new OneToManyCustomTaskPaneAdapter(taskPane);

                            if (!_ribbonTaskPanes.ContainsKey(ribbonViewModel))
                                _ribbonTaskPanes.Add(ribbonViewModel, new List<OneToManyCustomTaskPaneAdapter>());

                            _ribbonTaskPanes[ribbonViewModel].Add(taskPaneAdapter);
                            return taskPaneAdapter;
                        });
            }
            else
            {
                var adapters = _ribbonTaskPanes[ribbonViewModel];
                foreach (var taskPaneAdapter in adapters)
                {
                    if (!taskPaneAdapter.ViewRegistered(view))
                    {
                        foreach (var taskPaneRegistrationInfo in _registrationInfo[ribbonViewModel])
                        {
                            taskPaneAdapter.Add(Register(view, taskPaneRegistrationInfo));
                        }
                    }
                    else
                        taskPaneAdapter.Refresh(view);
                }
            }
        }

        private CustomTaskPane Register(object view, TaskPaneRegistrationInfo taskPaneRegistrationInfo)
        {
            var taskPane = _customTaskPaneCollection.Add(taskPaneRegistrationInfo.ControlFactory(), taskPaneRegistrationInfo.Title, view);

            return taskPane;
        }

        public void Cleanup(object view)
        {
            foreach (var adapter in _ribbonTaskPanes.Values.SelectMany(v=>v))
            {
                adapter.CleanupView(view);
            }
        }
    }
}
