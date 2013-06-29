------------------------------------------------------------------------------------------
                                      Introduction
------------------------------------------------------------------------------------------

VSTOContrib.Autofac allows you to use Autofac to resolve your view models



------------------------------------------------------------------------------------------
                                    Usage Instructions
------------------------------------------------------------------------------------------

1. Open your ThisAddIn.cs file
2. Go to the CreateRibbonExtensibilityObject() method
3. Replace `new <OFFICEAPP>RibbonFactory(new DefaultViewModelFactory(), ...)` with
           `new <OFFICEAPP>RibbonFactory(new AutofacViewModelFactory(container), ...)` or
           `new <OFFICEAPP>RibbonFactory(new AutofacViewModelFactory(new AddinAutofacModule()), ...)`
4. In your container registration register your ribbon view models with
           `builder.RegisterRibbonViewModels(typeof(AddinModule).Assembly);` (exists in the VSTOContrib.Autofac namespace)
   

VSTOContrib will now create a lifetime scope for each viewmodel it resolves, and dispose the lifetime scope when the viewmodel is cleaned up