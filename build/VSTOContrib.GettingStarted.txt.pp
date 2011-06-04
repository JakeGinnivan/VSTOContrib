**Note - VSTO Contrib is still under active development, and as such the bootstrapping method signatures may change.
Changes will be documented, and simple to fix. Just be aware that your solution may not compile after an update at this stage.**

------------------------------------------------------------------------------------------
                             To use the Ribbon Factory
------------------------------------------------------------------------------------------

 1. Expand VSTO generated code (or delete the region). 
 2. Ignore the comment about not modifying the contents of the internal startup method
 3. Edit the InternalStartup method to look like this (or copy and overwrite)

        private void InternalStartup()
        {
            _core = new AddinBootstrapper();
            {{Application}}RibbonFactory.SetApplication(Application);
            RibbonFactory.Current.InitialiseFactory(
                t => (IRibbonViewModel)_core.Resolve(t),
                CustomTaskPanes);

            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }

 4. Override the CreateRibbonExtensibilityObject method to specify the VSTO Contrib Ribbon Factory

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new {{Application}}RibbonFactory(typeof(AddinBootstrapper).Assembly);
        }

 5. Modify the ThisAddinShutdown method and add the following two lines

            _core.Dispose();
            System.Windows.Application.Current.Shutdown();

 6. Move the AddinBootstrapper.cs to a class library, this class library will hold all your application logic, and will be testable!

------------------------------------------------------------------------------------------
                                     To Create a ViewModel
------------------------------------------------------------------------------------------

*Coming soon - a VSTOContrib.RibbonFactorySample NuGet package, which will do this for you

 1. In your Class library, or Core project (not the add-in project), create a new class, you could call it DocumentViewModel for a word add-in for instance.

 2. Here is an example ViewModel for a Word add-in, the sample project will add the samples based on the Microsoft.Office.Interop.* reference

using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.RibbonFactory.Internal;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;

namespace RazorDocs.Core
{
    [WordRibbonViewModel]
    public class RazorDocs : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
    {
        private bool _panelShown;
        private Document _document;
        private ICustomTaskPaneWrapper _myAddinTaskPane;

        public void Initialised(object context)
        {
            _document = (Document) context;
        }

        public void CurrentViewChanged(object currentView)
        {
        }

        public IRibbonUI RibbonUi { get; set; }

        public bool PanelShown
        {
            get { return _panelShown; }
            set
            {
                if (_panelShown == value) return;
                _panelShown = value;
                _myAddinTaskPane.Visible = value;
                RaisePropertyChanged(()=>PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            _myAddinTaskPane = register(
                ()=>new WpfPanelHost
                {
                    Child = new MyAddinPanel //This is a WPF User control
                    {
                        DataContext = new MyAddinPanelViewModel(this) //Viewmodel for the user control
                    }
                }, "MyAddin Awesome Taskpane");
            _myAddinTaskPane.Visible = true;
            PanelShown = true;
            _myAddinTaskPane.VisibleChanged += TaskPaneVisibleChanged;
            TaskPaneVisibleChanged(this, EventArgs.Empty);
        }

        public void Cleanup()
        {
            _myAddinTaskPane.VisibleChanged -= TaskPaneVisibleChanged;
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            _panelShown = _myAddinTaskPane.Visible;
            RaisePropertyChanged(()=>PanelShown);
        }
    }
}

------------------------------------------------------------------------------------------
							And the Ribbon XML
------------------------------------------------------------------------------------------

*Notice that the onAction and getPressed are the same, VSTO Contrib will bind the callbacks to the PanelShown property on your ViewModel!*

<?xml version="1.0" encoding="UTF-8"?>
<customUI onLoad="Ribbon_Load" xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="myAddinGroup" label="Sample Addin Group">
          <toggleButton id="showMyAddinPaneButton" onAction="PanelShown" getPressed="PanelShown" label="Show Panel" showImage="false" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>

------------------------------------------------------------------------------------------

Grab the VSTO Contrib source from codeplex for some sample projects...

------------------------------------------------------------------------------------------

For Autofac integration use this bootstrapper instead of the default one

using System;
using Autofac;

namespace $rootnamespace$Core
{
    public class AddinBootstrapper : IDisposable
    {
        private readonly IContainer _container;

        public AddinBootstrapper()
        {
            var containerBuilder = new ContainerBuilder();

            RegisterComponents(containerBuilder);

            _container = containerBuilder.Build();
        }

        private static void RegisterComponents(ContainerBuilder containerBuilder)
        {
            
        }

        public object Resolve(Type type)
        {
            return _container.Resolve(type);
        }

        public T Resolve<T>()
        {
            return _container.Resolve<T>();
        }

        public void Dispose()
        {
            _container.Dispose();
        }
    }
}
