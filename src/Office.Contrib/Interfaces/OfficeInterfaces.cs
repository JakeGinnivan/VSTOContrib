using System;

namespace Office.Contrib.Interfaces
{
	/// <summary>
	/// Wrapper interface for IAccessible which adds IDispose to the interface
	/// </summary>
	public interface IIAccessible : Microsoft.Office.Core.IAccessible, IDisposable { }

	/// <summary>
	/// Wrapper interface for _IMsoDispObj which adds IDispose to the interface
	/// </summary>
	public interface I_IMsoDispObj : Microsoft.Office.Core._IMsoDispObj, IDisposable { }

	/// <summary>
	/// Wrapper interface for _IMsoOleAccDispObj which adds IDispose to the interface
	/// </summary>
	public interface I_IMsoOleAccDispObj : Microsoft.Office.Core._IMsoOleAccDispObj, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBars which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBars : Microsoft.Office.Core._CommandBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBar which adds IDispose to the interface
	/// </summary>
	public interface ICommandBar : Microsoft.Office.Core.CommandBar, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarControls which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarControls : Microsoft.Office.Core.CommandBarControls, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarControl which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarControl : Microsoft.Office.Core.CommandBarControl, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarButton which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButton : Microsoft.Office.Core._CommandBarButton, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarPopup which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarPopup : Microsoft.Office.Core.CommandBarPopup, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarComboBox which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBox : Microsoft.Office.Core._CommandBarComboBox, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarActiveX which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarActiveX : Microsoft.Office.Core._CommandBarActiveX, IDisposable { }

	/// <summary>
	/// Wrapper interface for Adjustments which adds IDispose to the interface
	/// </summary>
	public interface IAdjustments : Microsoft.Office.Core.Adjustments, IDisposable { }

	/// <summary>
	/// Wrapper interface for CalloutFormat which adds IDispose to the interface
	/// </summary>
	public interface ICalloutFormat : Microsoft.Office.Core.CalloutFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IColorFormat : Microsoft.Office.Core.ColorFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IConnectorFormat : Microsoft.Office.Core.ConnectorFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for FillFormat which adds IDispose to the interface
	/// </summary>
	public interface IFillFormat : Microsoft.Office.Core.FillFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IFreeformBuilder : Microsoft.Office.Core.FreeformBuilder, IDisposable { }

	/// <summary>
	/// Wrapper interface for GroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IGroupShapes : Microsoft.Office.Core.GroupShapes, IDisposable { }

	/// <summary>
	/// Wrapper interface for LineFormat which adds IDispose to the interface
	/// </summary>
	public interface ILineFormat : Microsoft.Office.Core.LineFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShapeNode which adds IDispose to the interface
	/// </summary>
	public interface IShapeNode : Microsoft.Office.Core.ShapeNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShapeNodes which adds IDispose to the interface
	/// </summary>
	public interface IShapeNodes : Microsoft.Office.Core.ShapeNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for PictureFormat which adds IDispose to the interface
	/// </summary>
	public interface IPictureFormat : Microsoft.Office.Core.PictureFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShadowFormat which adds IDispose to the interface
	/// </summary>
	public interface IShadowFormat : Microsoft.Office.Core.ShadowFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for Script which adds IDispose to the interface
	/// </summary>
	public interface IScript : Microsoft.Office.Core.Script, IDisposable { }

	/// <summary>
	/// Wrapper interface for Scripts which adds IDispose to the interface
	/// </summary>
	public interface IScripts : Microsoft.Office.Core.Scripts, IDisposable { }

	/// <summary>
	/// Wrapper interface for Shape which adds IDispose to the interface
	/// </summary>
	public interface IShape : Microsoft.Office.Core.Shape, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IShapeRange : Microsoft.Office.Core.ShapeRange, IDisposable { }

	/// <summary>
	/// Wrapper interface for Shapes which adds IDispose to the interface
	/// </summary>
	public interface IShapes : Microsoft.Office.Core.Shapes, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
	/// </summary>
	public interface ITextEffectFormat : Microsoft.Office.Core.TextEffectFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextFrame which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame : Microsoft.Office.Core.TextFrame, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
	/// </summary>
	public interface IThreeDFormat : Microsoft.Office.Core.ThreeDFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDispCagNotifySink which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDispCagNotifySink : Microsoft.Office.Core.IMsoDispCagNotifySink, IDisposable { }

	/// <summary>
	/// Wrapper interface for Balloon which adds IDispose to the interface
	/// </summary>
	public interface IBalloon : Microsoft.Office.Core.Balloon, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonCheckboxes which adds IDispose to the interface
	/// </summary>
	public interface IBalloonCheckboxes : Microsoft.Office.Core.BalloonCheckboxes, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonCheckbox which adds IDispose to the interface
	/// </summary>
	public interface IBalloonCheckbox : Microsoft.Office.Core.BalloonCheckbox, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonLabels which adds IDispose to the interface
	/// </summary>
	public interface IBalloonLabels : Microsoft.Office.Core.BalloonLabels, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonLabel which adds IDispose to the interface
	/// </summary>
	public interface IBalloonLabel : Microsoft.Office.Core.BalloonLabel, IDisposable { }

	/// <summary>
	/// Wrapper interface for AnswerWizardFiles which adds IDispose to the interface
	/// </summary>
	public interface IAnswerWizardFiles : Microsoft.Office.Core.AnswerWizardFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for AnswerWizard which adds IDispose to the interface
	/// </summary>
	public interface IAnswerWizard : Microsoft.Office.Core.AnswerWizard, IDisposable { }

	/// <summary>
	/// Wrapper interface for Assistant which adds IDispose to the interface
	/// </summary>
	public interface IAssistant : Microsoft.Office.Core.Assistant, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentProperty which adds IDispose to the interface
	/// </summary>
	public interface IDocumentProperty : Microsoft.Office.Core.DocumentProperty, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentProperties which adds IDispose to the interface
	/// </summary>
	public interface IDocumentProperties : Microsoft.Office.Core.DocumentProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for IFoundFiles which adds IDispose to the interface
	/// </summary>
	public interface IIFoundFiles : Microsoft.Office.Core.IFoundFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for IFind which adds IDispose to the interface
	/// </summary>
	public interface IIFind : Microsoft.Office.Core.IFind, IDisposable { }

	/// <summary>
	/// Wrapper interface for FoundFiles which adds IDispose to the interface
	/// </summary>
	public interface IFoundFiles : Microsoft.Office.Core.FoundFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for PropertyTest which adds IDispose to the interface
	/// </summary>
	public interface IPropertyTest : Microsoft.Office.Core.PropertyTest, IDisposable { }

	/// <summary>
	/// Wrapper interface for PropertyTests which adds IDispose to the interface
	/// </summary>
	public interface IPropertyTests : Microsoft.Office.Core.PropertyTests, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileSearch which adds IDispose to the interface
	/// </summary>
	public interface IFileSearch : Microsoft.Office.Core.FileSearch, IDisposable { }

	/// <summary>
	/// Wrapper interface for COMAddIn which adds IDispose to the interface
	/// </summary>
	public interface ICOMAddIn : Microsoft.Office.Core.COMAddIn, IDisposable { }

	/// <summary>
	/// Wrapper interface for COMAddIns which adds IDispose to the interface
	/// </summary>
	public interface ICOMAddIns : Microsoft.Office.Core.COMAddIns, IDisposable { }

	/// <summary>
	/// Wrapper interface for LanguageSettings which adds IDispose to the interface
	/// </summary>
	public interface ILanguageSettings : Microsoft.Office.Core.LanguageSettings, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICommandBarsEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarsEvents : Microsoft.Office.Core.ICommandBarsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarsEvents : Microsoft.Office.Core._CommandBarsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarsEvents_Event : Microsoft.Office.Core._CommandBarsEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBars which adds IDispose to the interface
	/// </summary>
	public interface ICommandBars : Microsoft.Office.Core.CommandBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICommandBarComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarComboBoxEvents : Microsoft.Office.Core.ICommandBarComboBoxEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBoxEvents : Microsoft.Office.Core._CommandBarComboBoxEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarComboBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBoxEvents_Event : Microsoft.Office.Core._CommandBarComboBoxEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarComboBox which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarComboBox : Microsoft.Office.Core.CommandBarComboBox, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICommandBarButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarButtonEvents : Microsoft.Office.Core.ICommandBarButtonEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButtonEvents : Microsoft.Office.Core._CommandBarButtonEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarButtonEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButtonEvents_Event : Microsoft.Office.Core._CommandBarButtonEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarButton which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarButton : Microsoft.Office.Core.CommandBarButton, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebPageFont which adds IDispose to the interface
	/// </summary>
	public interface IWebPageFont : Microsoft.Office.Core.WebPageFont, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebPageFonts which adds IDispose to the interface
	/// </summary>
	public interface IWebPageFonts : Microsoft.Office.Core.WebPageFonts, IDisposable { }

	/// <summary>
	/// Wrapper interface for HTMLProjectItem which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProjectItem : Microsoft.Office.Core.HTMLProjectItem, IDisposable { }

	/// <summary>
	/// Wrapper interface for HTMLProjectItems which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProjectItems : Microsoft.Office.Core.HTMLProjectItems, IDisposable { }

	/// <summary>
	/// Wrapper interface for HTMLProject which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProject : Microsoft.Office.Core.HTMLProject, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions : Microsoft.Office.Core.MsoDebugOptions, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialogSelectedItems which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogSelectedItems : Microsoft.Office.Core.FileDialogSelectedItems, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialogFilter which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogFilter : Microsoft.Office.Core.FileDialogFilter, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialogFilters which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogFilters : Microsoft.Office.Core.FileDialogFilters, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialog which adds IDispose to the interface
	/// </summary>
	public interface IFileDialog : Microsoft.Office.Core.FileDialog, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureSet which adds IDispose to the interface
	/// </summary>
	public interface ISignatureSet : Microsoft.Office.Core.SignatureSet, IDisposable { }

	/// <summary>
	/// Wrapper interface for Signature which adds IDispose to the interface
	/// </summary>
	public interface ISignature : Microsoft.Office.Core.Signature, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVB which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVB : Microsoft.Office.Core.IMsoEnvelopeVB, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVBEvents which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVBEvents : Microsoft.Office.Core.IMsoEnvelopeVBEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVBEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVBEvents_Event : Microsoft.Office.Core.IMsoEnvelopeVBEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoEnvelope which adds IDispose to the interface
	/// </summary>
	public interface IMsoEnvelope : Microsoft.Office.Core.MsoEnvelope, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileTypes which adds IDispose to the interface
	/// </summary>
	public interface IFileTypes : Microsoft.Office.Core.FileTypes, IDisposable { }

	/// <summary>
	/// Wrapper interface for SearchFolders which adds IDispose to the interface
	/// </summary>
	public interface ISearchFolders : Microsoft.Office.Core.SearchFolders, IDisposable { }

	/// <summary>
	/// Wrapper interface for ScopeFolders which adds IDispose to the interface
	/// </summary>
	public interface IScopeFolders : Microsoft.Office.Core.ScopeFolders, IDisposable { }

	/// <summary>
	/// Wrapper interface for ScopeFolder which adds IDispose to the interface
	/// </summary>
	public interface IScopeFolder : Microsoft.Office.Core.ScopeFolder, IDisposable { }

	/// <summary>
	/// Wrapper interface for SearchScope which adds IDispose to the interface
	/// </summary>
	public interface ISearchScope : Microsoft.Office.Core.SearchScope, IDisposable { }

	/// <summary>
	/// Wrapper interface for SearchScopes which adds IDispose to the interface
	/// </summary>
	public interface ISearchScopes : Microsoft.Office.Core.SearchScopes, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDiagram which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDiagram : Microsoft.Office.Core.IMsoDiagram, IDisposable { }

	/// <summary>
	/// Wrapper interface for DiagramNodes which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodes : Microsoft.Office.Core.DiagramNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodeChildren : Microsoft.Office.Core.DiagramNodeChildren, IDisposable { }

	/// <summary>
	/// Wrapper interface for DiagramNode which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNode : Microsoft.Office.Core.DiagramNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for CanvasShapes which adds IDispose to the interface
	/// </summary>
	public interface ICanvasShapes : Microsoft.Office.Core.CanvasShapes, IDisposable { }

	/// <summary>
	/// Wrapper interface for OfficeDataSourceObject which adds IDispose to the interface
	/// </summary>
	public interface IOfficeDataSourceObject : Microsoft.Office.Core.OfficeDataSourceObject, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOColumn which adds IDispose to the interface
	/// </summary>
	public interface IODSOColumn : Microsoft.Office.Core.ODSOColumn, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOColumns which adds IDispose to the interface
	/// </summary>
	public interface IODSOColumns : Microsoft.Office.Core.ODSOColumns, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOFilter which adds IDispose to the interface
	/// </summary>
	public interface IODSOFilter : Microsoft.Office.Core.ODSOFilter, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOFilters which adds IDispose to the interface
	/// </summary>
	public interface IODSOFilters : Microsoft.Office.Core.ODSOFilters, IDisposable { }

	/// <summary>
	/// Wrapper interface for NewFile which adds IDispose to the interface
	/// </summary>
	public interface INewFile : Microsoft.Office.Core.NewFile, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponent which adds IDispose to the interface
	/// </summary>
	public interface IWebComponent : Microsoft.Office.Core.WebComponent, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponentWindowExternal which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentWindowExternal : Microsoft.Office.Core.WebComponentWindowExternal, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponentFormat which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentFormat : Microsoft.Office.Core.WebComponentFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ILicWizExternal which adds IDispose to the interface
	/// </summary>
	public interface IILicWizExternal : Microsoft.Office.Core.ILicWizExternal, IDisposable { }

	/// <summary>
	/// Wrapper interface for ILicValidator which adds IDispose to the interface
	/// </summary>
	public interface IILicValidator : Microsoft.Office.Core.ILicValidator, IDisposable { }

	/// <summary>
	/// Wrapper interface for ILicAgent which adds IDispose to the interface
	/// </summary>
	public interface IILicAgent : Microsoft.Office.Core.ILicAgent, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEServicesDialog which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEServicesDialog : Microsoft.Office.Core.IMsoEServicesDialog, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponentProperties which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentProperties : Microsoft.Office.Core.WebComponentProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartDocument which adds IDispose to the interface
	/// </summary>
	public interface ISmartDocument : Microsoft.Office.Core.SmartDocument, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceMember which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceMember : Microsoft.Office.Core.SharedWorkspaceMember, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceMembers which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceMembers : Microsoft.Office.Core.SharedWorkspaceMembers, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceTask which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceTask : Microsoft.Office.Core.SharedWorkspaceTask, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceTasks which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceTasks : Microsoft.Office.Core.SharedWorkspaceTasks, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFile which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFile : Microsoft.Office.Core.SharedWorkspaceFile, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFiles which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFiles : Microsoft.Office.Core.SharedWorkspaceFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFolder which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFolder : Microsoft.Office.Core.SharedWorkspaceFolder, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFolders which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFolders : Microsoft.Office.Core.SharedWorkspaceFolders, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceLink which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceLink : Microsoft.Office.Core.SharedWorkspaceLink, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceLinks which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceLinks : Microsoft.Office.Core.SharedWorkspaceLinks, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspace which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspace : Microsoft.Office.Core.SharedWorkspace, IDisposable { }

	/// <summary>
	/// Wrapper interface for Sync which adds IDispose to the interface
	/// </summary>
	public interface ISync : Microsoft.Office.Core.Sync, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentLibraryVersion which adds IDispose to the interface
	/// </summary>
	public interface IDocumentLibraryVersion : Microsoft.Office.Core.DocumentLibraryVersion, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentLibraryVersions which adds IDispose to the interface
	/// </summary>
	public interface IDocumentLibraryVersions : Microsoft.Office.Core.DocumentLibraryVersions, IDisposable { }

	/// <summary>
	/// Wrapper interface for UserPermission which adds IDispose to the interface
	/// </summary>
	public interface IUserPermission : Microsoft.Office.Core.UserPermission, IDisposable { }

	/// <summary>
	/// Wrapper interface for Permission which adds IDispose to the interface
	/// </summary>
	public interface IPermission : Microsoft.Office.Core.Permission, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTRunResult which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTRunResult : Microsoft.Office.Core.MsoDebugOptions_UTRunResult, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UT which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UT : Microsoft.Office.Core.MsoDebugOptions_UT, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTs which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTs : Microsoft.Office.Core.MsoDebugOptions_UTs, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTManager which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTManager : Microsoft.Office.Core.MsoDebugOptions_UTManager, IDisposable { }

	/// <summary>
	/// Wrapper interface for MetaProperty which adds IDispose to the interface
	/// </summary>
	public interface IMetaProperty : Microsoft.Office.Core.MetaProperty, IDisposable { }

	/// <summary>
	/// Wrapper interface for MetaProperties which adds IDispose to the interface
	/// </summary>
	public interface IMetaProperties : Microsoft.Office.Core.MetaProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for PolicyItem which adds IDispose to the interface
	/// </summary>
	public interface IPolicyItem : Microsoft.Office.Core.PolicyItem, IDisposable { }

	/// <summary>
	/// Wrapper interface for ServerPolicy which adds IDispose to the interface
	/// </summary>
	public interface IServerPolicy : Microsoft.Office.Core.ServerPolicy, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentInspector which adds IDispose to the interface
	/// </summary>
	public interface IDocumentInspector : Microsoft.Office.Core.DocumentInspector, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentInspectors which adds IDispose to the interface
	/// </summary>
	public interface IDocumentInspectors : Microsoft.Office.Core.DocumentInspectors, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTask which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTask : Microsoft.Office.Core.WorkflowTask, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTasks which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTasks : Microsoft.Office.Core.WorkflowTasks, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTemplate which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTemplate : Microsoft.Office.Core.WorkflowTemplate, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTemplates which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTemplates : Microsoft.Office.Core.WorkflowTemplates, IDisposable { }

	/// <summary>
	/// Wrapper interface for IDocumentInspector which adds IDispose to the interface
	/// </summary>
	public interface IIDocumentInspector : Microsoft.Office.Core.IDocumentInspector, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureSetup which adds IDispose to the interface
	/// </summary>
	public interface ISignatureSetup : Microsoft.Office.Core.SignatureSetup, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureInfo which adds IDispose to the interface
	/// </summary>
	public interface ISignatureInfo : Microsoft.Office.Core.SignatureInfo, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureProvider which adds IDispose to the interface
	/// </summary>
	public interface ISignatureProvider : Microsoft.Office.Core.SignatureProvider, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLPrefixMapping which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPrefixMapping : Microsoft.Office.Core.CustomXMLPrefixMapping, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLPrefixMappings which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPrefixMappings : Microsoft.Office.Core.CustomXMLPrefixMappings, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLSchema which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLSchema : Microsoft.Office.Core.CustomXMLSchema, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLSchemaCollection which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLSchemaCollection : Microsoft.Office.Core._CustomXMLSchemaCollection, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLSchemaCollection which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLSchemaCollection : Microsoft.Office.Core.CustomXMLSchemaCollection, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLNodes which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLNodes : Microsoft.Office.Core.CustomXMLNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLNode which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLNode : Microsoft.Office.Core.CustomXMLNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLValidationError which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLValidationError : Microsoft.Office.Core.CustomXMLValidationError, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLValidationErrors which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLValidationErrors : Microsoft.Office.Core.CustomXMLValidationErrors, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPart which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPart : Microsoft.Office.Core._CustomXMLPart, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICustomXMLPartEvents which adds IDispose to the interface
	/// </summary>
	public interface IICustomXMLPartEvents : Microsoft.Office.Core.ICustomXMLPartEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartEvents : Microsoft.Office.Core._CustomXMLPartEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartEvents_Event : Microsoft.Office.Core._CustomXMLPartEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLPart which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPart : Microsoft.Office.Core.CustomXMLPart, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLParts which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLParts : Microsoft.Office.Core._CustomXMLParts, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICustomXMLPartsEvents which adds IDispose to the interface
	/// </summary>
	public interface IICustomXMLPartsEvents : Microsoft.Office.Core.ICustomXMLPartsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartsEvents : Microsoft.Office.Core._CustomXMLPartsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartsEvents_Event : Microsoft.Office.Core._CustomXMLPartsEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLParts which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLParts : Microsoft.Office.Core.CustomXMLParts, IDisposable { }

	/// <summary>
	/// Wrapper interface for GradientStop which adds IDispose to the interface
	/// </summary>
	public interface IGradientStop : Microsoft.Office.Core.GradientStop, IDisposable { }

	/// <summary>
	/// Wrapper interface for GradientStops which adds IDispose to the interface
	/// </summary>
	public interface IGradientStops : Microsoft.Office.Core.GradientStops, IDisposable { }

	/// <summary>
	/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
	/// </summary>
	public interface ISoftEdgeFormat : Microsoft.Office.Core.SoftEdgeFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for GlowFormat which adds IDispose to the interface
	/// </summary>
	public interface IGlowFormat : Microsoft.Office.Core.GlowFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
	/// </summary>
	public interface IReflectionFormat : Microsoft.Office.Core.ReflectionFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ParagraphFormat2 which adds IDispose to the interface
	/// </summary>
	public interface IParagraphFormat2 : Microsoft.Office.Core.ParagraphFormat2, IDisposable { }

	/// <summary>
	/// Wrapper interface for Font2 which adds IDispose to the interface
	/// </summary>
	public interface IFont2 : Microsoft.Office.Core.Font2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextColumn2 which adds IDispose to the interface
	/// </summary>
	public interface ITextColumn2 : Microsoft.Office.Core.TextColumn2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextRange2 which adds IDispose to the interface
	/// </summary>
	public interface ITextRange2 : Microsoft.Office.Core.TextRange2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextFrame2 which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame2 : Microsoft.Office.Core.TextFrame2, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeColor which adds IDispose to the interface
	/// </summary>
	public interface IThemeColor : Microsoft.Office.Core.ThemeColor, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeColorScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeColorScheme : Microsoft.Office.Core.ThemeColorScheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeFont which adds IDispose to the interface
	/// </summary>
	public interface IThemeFont : Microsoft.Office.Core.ThemeFont, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeFonts which adds IDispose to the interface
	/// </summary>
	public interface IThemeFonts : Microsoft.Office.Core.ThemeFonts, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeFontScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeFontScheme : Microsoft.Office.Core.ThemeFontScheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeEffectScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeEffectScheme : Microsoft.Office.Core.ThemeEffectScheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for OfficeTheme which adds IDispose to the interface
	/// </summary>
	public interface IOfficeTheme : Microsoft.Office.Core.OfficeTheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomTaskPane which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPane : Microsoft.Office.Core._CustomTaskPane, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomTaskPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface ICustomTaskPaneEvents : Microsoft.Office.Core.CustomTaskPaneEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomTaskPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPaneEvents : Microsoft.Office.Core._CustomTaskPaneEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomTaskPaneEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPaneEvents_Event : Microsoft.Office.Core._CustomTaskPaneEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomTaskPane which adds IDispose to the interface
	/// </summary>
	public interface ICustomTaskPane : Microsoft.Office.Core.CustomTaskPane, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICTPFactory which adds IDispose to the interface
	/// </summary>
	public interface IICTPFactory : Microsoft.Office.Core.ICTPFactory, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICustomTaskPaneConsumer which adds IDispose to the interface
	/// </summary>
	public interface IICustomTaskPaneConsumer : Microsoft.Office.Core.ICustomTaskPaneConsumer, IDisposable { }

	/// <summary>
	/// Wrapper interface for IRibbonUI which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonUI : Microsoft.Office.Core.IRibbonUI, IDisposable { }

	/// <summary>
	/// Wrapper interface for IRibbonControl which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonControl : Microsoft.Office.Core.IRibbonControl, IDisposable { }

	/// <summary>
	/// Wrapper interface for IRibbonExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonExtensibility : Microsoft.Office.Core.IRibbonExtensibility, IDisposable { }

	/// <summary>
	/// Wrapper interface for IAssistance which adds IDispose to the interface
	/// </summary>
	public interface IIAssistance : Microsoft.Office.Core.IAssistance, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartData which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartData : Microsoft.Office.Core.IMsoChartData, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChart which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChart : Microsoft.Office.Core.IMsoChart, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoCorners which adds IDispose to the interface
	/// </summary>
	public interface IIMsoCorners : Microsoft.Office.Core.IMsoCorners, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoLegend which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLegend : Microsoft.Office.Core.IMsoLegend, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoBorder which adds IDispose to the interface
	/// </summary>
	public interface IIMsoBorder : Microsoft.Office.Core.IMsoBorder, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoWalls which adds IDispose to the interface
	/// </summary>
	public interface IIMsoWalls : Microsoft.Office.Core.IMsoWalls, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoFloor which adds IDispose to the interface
	/// </summary>
	public interface IIMsoFloor : Microsoft.Office.Core.IMsoFloor, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoPlotArea which adds IDispose to the interface
	/// </summary>
	public interface IIMsoPlotArea : Microsoft.Office.Core.IMsoPlotArea, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartArea which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartArea : Microsoft.Office.Core.IMsoChartArea, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoSeriesLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoSeriesLines : Microsoft.Office.Core.IMsoSeriesLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoLeaderLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLeaderLines : Microsoft.Office.Core.IMsoLeaderLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for GridLines which adds IDispose to the interface
	/// </summary>
	public interface IGridLines : Microsoft.Office.Core.GridLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoUpBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoUpBars : Microsoft.Office.Core.IMsoUpBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDownBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDownBars : Microsoft.Office.Core.IMsoDownBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoInterior which adds IDispose to the interface
	/// </summary>
	public interface IIMsoInterior : Microsoft.Office.Core.IMsoInterior, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFillFormat : Microsoft.Office.Core.ChartFillFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for LegendEntries which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntries : Microsoft.Office.Core.LegendEntries, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartFont which adds IDispose to the interface
	/// </summary>
	public interface IChartFont : Microsoft.Office.Core.ChartFont, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartColorFormat : Microsoft.Office.Core.ChartColorFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for LegendEntry which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntry : Microsoft.Office.Core.LegendEntry, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoLegendKey which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLegendKey : Microsoft.Office.Core.IMsoLegendKey, IDisposable { }

	/// <summary>
	/// Wrapper interface for SeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface ISeriesCollection : Microsoft.Office.Core.SeriesCollection, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoSeries which adds IDispose to the interface
	/// </summary>
	public interface IIMsoSeries : Microsoft.Office.Core.IMsoSeries, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoErrorBars : Microsoft.Office.Core.IMsoErrorBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoTrendline which adds IDispose to the interface
	/// </summary>
	public interface IIMsoTrendline : Microsoft.Office.Core.IMsoTrendline, IDisposable { }

	/// <summary>
	/// Wrapper interface for Trendlines which adds IDispose to the interface
	/// </summary>
	public interface ITrendlines : Microsoft.Office.Core.Trendlines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDataLabels which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataLabels : Microsoft.Office.Core.IMsoDataLabels, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDataLabel which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataLabel : Microsoft.Office.Core.IMsoDataLabel, IDisposable { }

	/// <summary>
	/// Wrapper interface for Points which adds IDispose to the interface
	/// </summary>
	public interface IPoints : Microsoft.Office.Core.Points, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartPoint which adds IDispose to the interface
	/// </summary>
	public interface IChartPoint : Microsoft.Office.Core.ChartPoint, IDisposable { }

	/// <summary>
	/// Wrapper interface for Axes which adds IDispose to the interface
	/// </summary>
	public interface IAxes : Microsoft.Office.Core.Axes, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoAxis which adds IDispose to the interface
	/// </summary>
	public interface IIMsoAxis : Microsoft.Office.Core.IMsoAxis, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDataTable which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataTable : Microsoft.Office.Core.IMsoDataTable, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartTitle : Microsoft.Office.Core.IMsoChartTitle, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoAxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IIMsoAxisTitle : Microsoft.Office.Core.IMsoAxisTitle, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDisplayUnitLabel : Microsoft.Office.Core.IMsoDisplayUnitLabel, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoTickLabels which adds IDispose to the interface
	/// </summary>
	public interface IIMsoTickLabels : Microsoft.Office.Core.IMsoTickLabels, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoHyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IIMsoHyperlinks : Microsoft.Office.Core.IMsoHyperlinks, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDropLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDropLines : Microsoft.Office.Core.IMsoDropLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoHiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoHiLoLines : Microsoft.Office.Core.IMsoHiLoLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartGroup : Microsoft.Office.Core.IMsoChartGroup, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IChartGroups : Microsoft.Office.Core.ChartGroups, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoCharacters which adds IDispose to the interface
	/// </summary>
	public interface IIMsoCharacters : Microsoft.Office.Core.IMsoCharacters, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartFormat : Microsoft.Office.Core.IMsoChartFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for BulletFormat2 which adds IDispose to the interface
	/// </summary>
	public interface IBulletFormat2 : Microsoft.Office.Core.BulletFormat2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TabStops2 which adds IDispose to the interface
	/// </summary>
	public interface ITabStops2 : Microsoft.Office.Core.TabStops2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TabStop2 which adds IDispose to the interface
	/// </summary>
	public interface ITabStop2 : Microsoft.Office.Core.TabStop2, IDisposable { }

	/// <summary>
	/// Wrapper interface for Ruler2 which adds IDispose to the interface
	/// </summary>
	public interface IRuler2 : Microsoft.Office.Core.Ruler2, IDisposable { }

	/// <summary>
	/// Wrapper interface for RulerLevels2 which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevels2 : Microsoft.Office.Core.RulerLevels2, IDisposable { }

	/// <summary>
	/// Wrapper interface for RulerLevel2 which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevel2 : Microsoft.Office.Core.RulerLevel2, IDisposable { }

	/// <summary>
	/// Wrapper interface for EncryptionProvider which adds IDispose to the interface
	/// </summary>
	public interface IEncryptionProvider : Microsoft.Office.Core.EncryptionProvider, IDisposable { }

	/// <summary>
	/// Wrapper interface for IBlogExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIBlogExtensibility : Microsoft.Office.Core.IBlogExtensibility, IDisposable { }

	/// <summary>
	/// Wrapper interface for IBlogPictureExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIBlogPictureExtensibility : Microsoft.Office.Core.IBlogPictureExtensibility, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverterPreferences which adds IDispose to the interface
	/// </summary>
	public interface IIConverterPreferences : Microsoft.Office.Core.IConverterPreferences, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverterApplicationPreferences which adds IDispose to the interface
	/// </summary>
	public interface IIConverterApplicationPreferences : Microsoft.Office.Core.IConverterApplicationPreferences, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverterUICallback which adds IDispose to the interface
	/// </summary>
	public interface IIConverterUICallback : Microsoft.Office.Core.IConverterUICallback, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverter which adds IDispose to the interface
	/// </summary>
	public interface IIConverter : Microsoft.Office.Core.IConverter, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArt which adds IDispose to the interface
	/// </summary>
	public interface ISmartArt : Microsoft.Office.Core.SmartArt, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtNodes which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtNodes : Microsoft.Office.Core.SmartArtNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtNode which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtNode : Microsoft.Office.Core.SmartArtNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtLayouts which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtLayouts : Microsoft.Office.Core.SmartArtLayouts, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtLayout which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtLayout : Microsoft.Office.Core.SmartArtLayout, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtQuickStyles which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtQuickStyles : Microsoft.Office.Core.SmartArtQuickStyles, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtQuickStyle which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtQuickStyle : Microsoft.Office.Core.SmartArtQuickStyle, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtColors which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtColors : Microsoft.Office.Core.SmartArtColors, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtColor which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtColor : Microsoft.Office.Core.SmartArtColor, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerField which adds IDispose to the interface
	/// </summary>
	public interface IPickerField : Microsoft.Office.Core.PickerField, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerFields which adds IDispose to the interface
	/// </summary>
	public interface IPickerFields : Microsoft.Office.Core.PickerFields, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerProperty which adds IDispose to the interface
	/// </summary>
	public interface IPickerProperty : Microsoft.Office.Core.PickerProperty, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerProperties which adds IDispose to the interface
	/// </summary>
	public interface IPickerProperties : Microsoft.Office.Core.PickerProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerResult which adds IDispose to the interface
	/// </summary>
	public interface IPickerResult : Microsoft.Office.Core.PickerResult, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerResults which adds IDispose to the interface
	/// </summary>
	public interface IPickerResults : Microsoft.Office.Core.PickerResults, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerDialog which adds IDispose to the interface
	/// </summary>
	public interface IPickerDialog : Microsoft.Office.Core.PickerDialog, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoContactCard which adds IDispose to the interface
	/// </summary>
	public interface IIMsoContactCard : Microsoft.Office.Core.IMsoContactCard, IDisposable { }

	/// <summary>
	/// Wrapper interface for EffectParameter which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameter : Microsoft.Office.Core.EffectParameter, IDisposable { }

	/// <summary>
	/// Wrapper interface for EffectParameters which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameters : Microsoft.Office.Core.EffectParameters, IDisposable { }

	/// <summary>
	/// Wrapper interface for PictureEffect which adds IDispose to the interface
	/// </summary>
	public interface IPictureEffect : Microsoft.Office.Core.PictureEffect, IDisposable { }

	/// <summary>
	/// Wrapper interface for PictureEffects which adds IDispose to the interface
	/// </summary>
	public interface IPictureEffects : Microsoft.Office.Core.PictureEffects, IDisposable { }

	/// <summary>
	/// Wrapper interface for Crop which adds IDispose to the interface
	/// </summary>
	public interface ICrop : Microsoft.Office.Core.Crop, IDisposable { }

	/// <summary>
	/// Wrapper interface for ContactCard which adds IDispose to the interface
	/// </summary>
	public interface IContactCard : Microsoft.Office.Core.ContactCard, IDisposable { }

	}