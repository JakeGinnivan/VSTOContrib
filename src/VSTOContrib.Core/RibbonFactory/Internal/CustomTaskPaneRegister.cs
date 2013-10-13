using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Interfaces.Internal;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class CustomTaskPaneRegister : ICustomTaskPaneRegister
    {
        readonly Func<CustomTaskPaneCollection> customTaskPaneCollection;
        private readonly Dictionary<IRibbonViewModel, List<TaskPaneRegistrationInfo>> registrationInfo;
        private readonly Dictionary<IRibbonViewModel, List<OneToManyCustomTaskPaneAdapter>> ribbonTaskPanes;

        public CustomTaskPaneRegister(Func<object> customTaskPaneCollection)
        {
            this.customTaskPaneCollection = ()=>(CustomTaskPaneCollection)customTaskPaneCollection();
            registrationInfo = new Dictionary<IRibbonViewModel, List<TaskPaneRegistrationInfo>>();
            ribbonTaskPanes = new Dictionary<IRibbonViewModel, List<OneToManyCustomTaskPaneAdapter>>();
        }

        public void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view, object viewContext)
        {
            var registersCustomTaskPanes = ribbonViewModel as IRegisterCustomTaskPane;
            if (registersCustomTaskPanes == null) return;

            if (!registrationInfo.ContainsKey(ribbonViewModel))
            {
                registersCustomTaskPanes.RegisterTaskPanes(
                    (controlFactory, title, initiallyVisible) =>
                        {
                            var taskPaneRegistrationInfo = new TaskPaneRegistrationInfo(controlFactory, title);
                            if (!registrationInfo.ContainsKey(ribbonViewModel))
                                registrationInfo.Add(ribbonViewModel, new List<TaskPaneRegistrationInfo>());
                            registrationInfo[ribbonViewModel].Add(taskPaneRegistrationInfo);

                            var taskPane = Register(view, taskPaneRegistrationInfo);
                            var taskPaneAdapter = new OneToManyCustomTaskPaneAdapter(taskPane, viewContext)
                            {
                                Visible = initiallyVisible
                            };

                            if (!ribbonTaskPanes.ContainsKey(ribbonViewModel))
                                ribbonTaskPanes.Add(ribbonViewModel, new List<OneToManyCustomTaskPaneAdapter>());

                            ribbonTaskPanes[ribbonViewModel].Add(taskPaneAdapter);
                            return taskPaneAdapter;
                        });
            }
            else
            {
                var adapters = ribbonTaskPanes[ribbonViewModel];
                foreach (var taskPaneAdapter in adapters)
                {
                    if (!taskPaneAdapter.ViewRegistered(view))
                    {
                        foreach (var taskPaneRegistrationInfo in registrationInfo[ribbonViewModel])
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
            var taskPane = customTaskPaneCollection().Add(taskPaneRegistrationInfo.ControlFactory(), taskPaneRegistrationInfo.Title, view);

            return taskPane;
        }

        public void Cleanup(object view)
        {
            foreach (var adapter in ribbonTaskPanes.Values.SelectMany(v=>v))
            {
                adapter.CleanupView(view);
            }
        }

        public void ChangeVisibilityForContext(object context, bool visible)
        {
            foreach (var adapter in ribbonTaskPanes.Values.SelectMany(v => v).Where(v=>v.ViewContext == context))
            {
                //TODO Remember what it was 
                adapter.Visible = visible;
            }
        }
    }
}
