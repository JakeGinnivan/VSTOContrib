using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.Office.Tools;
using VSTOContrib.Core.Domain;
using VSTOContrib.Core.RibbonFactory.Interfaces;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class CustomTaskPaneRegister
    {
        readonly Dictionary<OfficeContextDomain, List<TaskPaneRegistration>> registrations;
        Lazy<CustomTaskPaneCollection> customTaskPaneCollection;

        public CustomTaskPaneRegister(AddInBase addinBase, OfficeApplicationDomain domain)
        {
            customTaskPaneCollection = new Lazy<CustomTaskPaneCollection>(() =>
            {
                var field = addinBase.GetType().GetField("CustomTaskPanes", BindingFlags.Instance | BindingFlags.NonPublic);
                return (CustomTaskPaneCollection)field.GetValue(addinBase);
            });

            domain.NewContext += DomainOnNewContext;
            domain.ContextClosed += DomainOnContextClosed;
            registrations = new Dictionary<OfficeContextDomain, List<TaskPaneRegistration>>();
        }

        void DomainOnNewContext(OfficeContextDomain officeContextDomain)
        {
            var registerCustomTaskPane = officeContextDomain.ViewModel as IRegisterCustomTaskPane;
            if (registerCustomTaskPane == null) return;
            officeContextDomain.NewView += OfficeContextDomainOnNewView;
            officeContextDomain.ViewClosed += OfficeContextDomainOnViewClosed;
            var contextRegistrations = new List<TaskPaneRegistration>();
            registrations.Add(officeContextDomain, contextRegistrations);
            registerCustomTaskPane.RegisterTaskPanes((controlFactory, title, initiallyVisible) =>
            {
                var registrationInfo = new TaskPaneRegistrationInfo(controlFactory, title);
                var registration = new TaskPaneRegistration(registrationInfo, new OneToManyCustomTaskPaneAdapter(title));
                contextRegistrations.Add(registration);
                return registration.Adapter;
            });

            foreach (var view in officeContextDomain.Views)
            {
                OfficeContextDomainOnNewView(view);
            }
        }

        void DomainOnContextClosed(OfficeContextDomain officeContextDomain)
        {
            officeContextDomain.NewView -= OfficeContextDomainOnNewView;
            officeContextDomain.ViewClosed -= OfficeContextDomainOnViewClosed;
        }

        void OfficeContextDomainOnNewView(OfficeViewDomain officeViewDomain)
        {
            foreach (var taskPaneRegistration in TaskPaneRegistrationsForView(officeViewDomain))
            {
                var control = taskPaneRegistration.RegistrationInfo.ControlFactory();
                var customTaskPane = customTaskPaneCollection.Value.Add(control, taskPaneRegistration.RegistrationInfo.Title);
                taskPaneRegistration.Adapter.Add(officeViewDomain.Window, customTaskPane);
            }
        }

        void OfficeContextDomainOnViewClosed(OfficeViewDomain officeViewDomain)
        {
            foreach (var taskPaneRegistration in TaskPaneRegistrationsForView(officeViewDomain))
            {
                taskPaneRegistration.Adapter.CleanupView(officeViewDomain.Window);
            }
        }

        IEnumerable<TaskPaneRegistration> TaskPaneRegistrationsForView(OfficeViewDomain officeViewDomain)
        {
            return officeViewDomain.Contexts
                .Select(context => registrations[context])
                .SelectMany(registration => registration);
        }

        public void Dispose()
        {
            foreach (var registration in registrations.Values.SelectMany(_ => _))
            {
                registration.Adapter.Dispose();
            }
            registrations.Clear();
            customTaskPaneCollection.Value.Dispose();
            customTaskPaneCollection = null;
        }
    }
}
