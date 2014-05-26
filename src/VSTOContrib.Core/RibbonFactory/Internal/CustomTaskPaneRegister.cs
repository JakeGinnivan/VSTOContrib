using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    internal class CustomTaskPaneRegister : ICustomTaskPaneRegister
    {
        Lazy<CustomTaskPaneCollection> customTaskPaneCollection;
        readonly Dictionary<IRibbonViewModel, List<TaskPaneRegistrationInfo>> registrationInfo;
        readonly Dictionary<IRibbonViewModel, List<OneToManyCustomTaskPaneAdapter>> ribbonTaskPanes;
        readonly Dictionary<object, List<IRibbonViewModel>> windowToTaskPaneLookup;

        public CustomTaskPaneRegister(AddInBase addinBase)
        {
            customTaskPaneCollection = new Lazy<CustomTaskPaneCollection>(() =>
            {
                var field = addinBase.GetType().GetField("CustomTaskPanes", BindingFlags.Instance | BindingFlags.NonPublic);
                return (CustomTaskPaneCollection)field.GetValue(addinBase);
            });
            registrationInfo = new Dictionary<IRibbonViewModel, List<TaskPaneRegistrationInfo>>();
            ribbonTaskPanes = new Dictionary<IRibbonViewModel, List<OneToManyCustomTaskPaneAdapter>>();
            windowToTaskPaneLookup = new Dictionary<object, List<IRibbonViewModel>>();
        }

        public void RegisterCustomTaskPanes(IRibbonViewModel ribbonViewModel, object view, object viewContext)
        {
            var registersCustomTaskPanes = ribbonViewModel as IRegisterCustomTaskPane;
            if (registersCustomTaskPanes == null) return;

            if (!registrationInfo.ContainsKey(ribbonViewModel))
            {
                registersCustomTaskPanes.RegisterTaskPanes((controlFactory, title, initiallyVisible) =>
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

                    if (!windowToTaskPaneLookup.ContainsKey(view))
                        windowToTaskPaneLookup.Add(view, new List<IRibbonViewModel>());

                    ribbonTaskPanes[ribbonViewModel].Add(taskPaneAdapter);
                    windowToTaskPaneLookup[view].Add(ribbonViewModel);
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

            foreach (var oneToManyCustomTaskPaneAdapter in windowToTaskPaneLookup[view]
                .Except(new[] { ribbonViewModel })
                .SelectMany(viewModelToHide => ribbonTaskPanes[viewModelToHide]))
            {
                oneToManyCustomTaskPaneAdapter.HideIfVisible();
            }

            foreach (var toRestore in ribbonTaskPanes[ribbonViewModel])
            {
                toRestore.RestoreIfNeeded();
            }
        }

        private CustomTaskPane Register(object view, TaskPaneRegistrationInfo taskPaneRegistrationInfo)
        {
            var taskPane = customTaskPaneCollection.Value.Add(taskPaneRegistrationInfo.ControlFactory(), taskPaneRegistrationInfo.Title, view);

            return taskPane;
        }

        public void Cleanup(object view)
        {
            foreach (var adapter in ribbonTaskPanes.Values.SelectMany(v => v))
            {
                adapter.CleanupView(view);
            }
        }

        public void CleanupViewModel(IRibbonViewModel viewModelInstance)
        {
            var adaptersForViewModel = ribbonTaskPanes[viewModelInstance];
            ribbonTaskPanes.Remove(viewModelInstance);
            foreach (var oneToManyCustomTaskPaneAdapter in adaptersForViewModel)
            {
                oneToManyCustomTaskPaneAdapter.Dispose();
            }
        }

        public void Dispose()
        {
            var taskPanes = ribbonTaskPanes.ToArray();
            ribbonTaskPanes.Clear();
            foreach (var ribbonTaskPane in taskPanes)
            {
                var oneToManyCustomTaskPaneAdapters = ribbonTaskPane.Value.ToArray();
                ribbonTaskPane.Value.Clear();
                foreach (var oneToManyCustomTaskPaneAdapter in oneToManyCustomTaskPaneAdapters)
                {
                    oneToManyCustomTaskPaneAdapter.Dispose();
                }
            }

            customTaskPaneCollection.Value.Dispose();
            customTaskPaneCollection = null;
        }
    }
}
