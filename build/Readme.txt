------------------------------------------------------------------------------------------

Feedback is very welcome!

Start a discussion or raise an issue at http://vstocontrib.codeplex.com/ or https://github.com/JakeGinnivan/VSTOContrib

------------------------------------------------------------------------------------------
                                      Introduction
------------------------------------------------------------------------------------------
VSTO Contrib lets you easily unit test, use IoC/DI and develop in a MVVM style within Office Add-ins. 

It supports Outlook, Word, Excel and PowerPoint 2010+, and supports .net 4.0.


------------------------------------------------------------------------------------------
                                      Terminology
------------------------------------------------------------------------------------------


|  Term    |                             Meaning                                         |
|----------------------------------------------------------------------------------------|
|   View   |   The Window/Inspector                                                      |
| Context  |   The document, spreadsheet, slideshow or Item (MailItem, ContactItem etc)  |


------------------------------------------------------------------------------------------
                                    Breaking Changes
------------------------------------------------------------------------------------------
Since v0.10.x there have been numerous changes to the internals of VSTO Contrib, this
is to allow full support for Office 2013 and simplify the getting started experience!

 - RaisePropertyChanged has been removed, use OnPropertyChanged instead
 - NotifyPropertyChanged base class is available to use for your WPF view models
 - Bootstrapping is far simpler, see the getting started for the new bootstrapping code
   - No longer need to modify internal startup

------------------------------------------------------------------------------------------
                                      Whats New
------------------------------------------------------------------------------------------

 - A viewmodel is created when there is no context (for example, word is open with no documents)
 - The VSTO Factory is now available in viewmodels, so you can get VSTO objects, for example in word:
    - vstoDocument = ((ApplicationFactory)VstoFactory).GetVstoObject(document);
      vstoDocument.SelectionChange += VstoDocumentOnSelectionChange;

------------------------------------------------------------------------------------------
                                   Getting Started
------------------------------------------------------------------------------------------
If you want to manually or you are putting VSTO Contrib into an existing project 
follow these instructions. If you have just installed VSTO Contrib into a new project
install the QuickStart projects

1. Create an empty class library (Maybe MyAddin.Core ?)
2. Override the CreateRibbonExtensibilityObject method to specify the VSTO Contrib Ribbon Factory
   There are a number of arguments required. This is because much of the VSTO addin is code generated, so you have to pass it in

	protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
	{
		//Required for WPF support
        if (System.Windows.Application.Current == null)
            new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };

		var assemblyContainingViewModels = typeof (ThisAddIn).Assembly; // This should be the assembly containing all your VSTOContrib viewmodels
		return new VSTOContrib.{{Application}}.RibbonFactory.{{Application}}RibbonFactory(new VSTOContrib.Core.DefaultViewModelFactory(), new Lazy<Microsoft.Office.Tools.CustomTaskPaneCollection>(() => CustomTaskPanes), Globals.Factory, assemblyContainingViewModels);
	}

3. Modify the `ThisAddin_Startup` method and add the following line

	VSTOContrib.Core.RibbonFactory.RibbonFactory.Current.SetApplication(Application, this);

4. Modify the `ThisAddin_Shutdown` method and add the following two lines

	System.Windows.Application.Current.Shutdown();


Now VSTO Contrib is ready to go! Next step is to create a ViewModel.

ViewModel's in VSTO Contrib are the way you interact with the application you are hosted in. VSTO Contrib will create an instance of a view model PER CONTEXT.

Contexts in VSTO Contrib vary per application, in Outlook context are appointments, mail items, and explorers (main windows). In word each document is a context,
PowerPoint is a presentation, and excel is a spreadsheet. VSTO Contrib will handle window management, ribbon events wire up and
registering custom task panes in each window etc.

Best of all, view models are testable! Follow the instructions to create your first view model.

-----------------------------------------------------------------------------------------
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

namespace MyAddin.Core
{
    [WordRibbonViewModel]
    public class DocumentViewModel : OfficeViewModelBase, IRibbonViewModel, IRegisterCustomTaskPane
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
                OnPropertyChanged(()=>PanelShown);
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
            OnPropertyChanged(()=>PanelShown);
        }
    }
}

------------------------------------------------------------------------------------------
							And the Ribbon XML
------------------------------------------------------------------------------------------

*Notice that the onAction and getPressed are the same, VSTO Contrib will bind the callbacks to the PanelShown property on your ViewModel!*

The RibbonXml file must be named the same as the viewmodel and be in the same folder, for example FooViewModel.cs can be named:
Foo.xml
FooView.xml
FooViewModel.xml

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
