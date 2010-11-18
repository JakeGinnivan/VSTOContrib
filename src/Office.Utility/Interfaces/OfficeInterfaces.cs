using System;
using Microsoft.Office.Core;

namespace Office.Utility
{
	/// <summary>
	/// Wrapper interface for IAccessible which adds IDispose to the interface
	/// </summary>
	public interface IIAccessible : IAccessible, IDisposable { }

	/// <summary>
	/// Wrapper interface for _IMsoDispObj which adds IDispose to the interface
	/// </summary>
	public interface I_IMsoDispObj : _IMsoDispObj, IDisposable { }

	/// <summary>
	/// Wrapper interface for _IMsoOleAccDispObj which adds IDispose to the interface
	/// </summary>
	public interface I_IMsoOleAccDispObj : _IMsoOleAccDispObj, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBars which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBars : _CommandBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBar which adds IDispose to the interface
	/// </summary>
	public interface ICommandBar : CommandBar, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarControls which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarControls : CommandBarControls, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarControl which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarControl : CommandBarControl, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarButton which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButton : _CommandBarButton, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarPopup which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarPopup : CommandBarPopup, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarComboBox which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBox : _CommandBarComboBox, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarActiveX which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarActiveX : _CommandBarActiveX, IDisposable { }

	/// <summary>
	/// Wrapper interface for Adjustments which adds IDispose to the interface
	/// </summary>
	public interface IAdjustments : Adjustments, IDisposable { }

	/// <summary>
	/// Wrapper interface for CalloutFormat which adds IDispose to the interface
	/// </summary>
	public interface ICalloutFormat : CalloutFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IColorFormat : ColorFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IConnectorFormat : ConnectorFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for FillFormat which adds IDispose to the interface
	/// </summary>
	public interface IFillFormat : FillFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IFreeformBuilder : FreeformBuilder, IDisposable { }

	/// <summary>
	/// Wrapper interface for GroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IGroupShapes : GroupShapes, IDisposable { }

	/// <summary>
	/// Wrapper interface for LineFormat which adds IDispose to the interface
	/// </summary>
	public interface ILineFormat : LineFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShapeNode which adds IDispose to the interface
	/// </summary>
	public interface IShapeNode : ShapeNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShapeNodes which adds IDispose to the interface
	/// </summary>
	public interface IShapeNodes : ShapeNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for PictureFormat which adds IDispose to the interface
	/// </summary>
	public interface IPictureFormat : PictureFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShadowFormat which adds IDispose to the interface
	/// </summary>
	public interface IShadowFormat : ShadowFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for Script which adds IDispose to the interface
	/// </summary>
	public interface IScript : Script, IDisposable { }

	/// <summary>
	/// Wrapper interface for Scripts which adds IDispose to the interface
	/// </summary>
	public interface IScripts : Scripts, IDisposable { }

	/// <summary>
	/// Wrapper interface for Shape which adds IDispose to the interface
	/// </summary>
	public interface IShape : Shape, IDisposable { }

	/// <summary>
	/// Wrapper interface for ShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IShapeRange : ShapeRange, IDisposable { }

	/// <summary>
	/// Wrapper interface for Shapes which adds IDispose to the interface
	/// </summary>
	public interface IShapes : Shapes, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
	/// </summary>
	public interface ITextEffectFormat : TextEffectFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextFrame which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame : TextFrame, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
	/// </summary>
	public interface IThreeDFormat : ThreeDFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDispCagNotifySink which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDispCagNotifySink : IMsoDispCagNotifySink, IDisposable { }

	/// <summary>
	/// Wrapper interface for Balloon which adds IDispose to the interface
	/// </summary>
	public interface IBalloon : Balloon, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonCheckboxes which adds IDispose to the interface
	/// </summary>
	public interface IBalloonCheckboxes : BalloonCheckboxes, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonCheckbox which adds IDispose to the interface
	/// </summary>
	public interface IBalloonCheckbox : BalloonCheckbox, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonLabels which adds IDispose to the interface
	/// </summary>
	public interface IBalloonLabels : BalloonLabels, IDisposable { }

	/// <summary>
	/// Wrapper interface for BalloonLabel which adds IDispose to the interface
	/// </summary>
	public interface IBalloonLabel : BalloonLabel, IDisposable { }

	/// <summary>
	/// Wrapper interface for AnswerWizardFiles which adds IDispose to the interface
	/// </summary>
	public interface IAnswerWizardFiles : AnswerWizardFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for AnswerWizard which adds IDispose to the interface
	/// </summary>
	public interface IAnswerWizard : AnswerWizard, IDisposable { }

	/// <summary>
	/// Wrapper interface for Assistant which adds IDispose to the interface
	/// </summary>
	public interface IAssistant : Assistant, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentProperty which adds IDispose to the interface
	/// </summary>
	public interface IDocumentProperty : DocumentProperty, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentProperties which adds IDispose to the interface
	/// </summary>
	public interface IDocumentProperties : DocumentProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for IFoundFiles which adds IDispose to the interface
	/// </summary>
	public interface IIFoundFiles : IFoundFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for IFind which adds IDispose to the interface
	/// </summary>
	public interface IIFind : IFind, IDisposable { }

	/// <summary>
	/// Wrapper interface for FoundFiles which adds IDispose to the interface
	/// </summary>
	public interface IFoundFiles : FoundFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for PropertyTest which adds IDispose to the interface
	/// </summary>
	public interface IPropertyTest : PropertyTest, IDisposable { }

	/// <summary>
	/// Wrapper interface for PropertyTests which adds IDispose to the interface
	/// </summary>
	public interface IPropertyTests : PropertyTests, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileSearch which adds IDispose to the interface
	/// </summary>
	public interface IFileSearch : FileSearch, IDisposable { }

	/// <summary>
	/// Wrapper interface for COMAddIn which adds IDispose to the interface
	/// </summary>
	public interface ICOMAddIn : COMAddIn, IDisposable { }

	/// <summary>
	/// Wrapper interface for COMAddIns which adds IDispose to the interface
	/// </summary>
	public interface ICOMAddIns : COMAddIns, IDisposable { }

	/// <summary>
	/// Wrapper interface for LanguageSettings which adds IDispose to the interface
	/// </summary>
	public interface ILanguageSettings : LanguageSettings, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICommandBarsEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarsEvents : ICommandBarsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarsEvents : _CommandBarsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarsEvents_Event : _CommandBarsEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBars which adds IDispose to the interface
	/// </summary>
	public interface ICommandBars : CommandBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICommandBarComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarComboBoxEvents : ICommandBarComboBoxEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarComboBoxEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBoxEvents : _CommandBarComboBoxEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarComboBoxEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarComboBoxEvents_Event : _CommandBarComboBoxEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarComboBox which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarComboBox : CommandBarComboBox, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICommandBarButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface IICommandBarButtonEvents : ICommandBarButtonEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarButtonEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButtonEvents : _CommandBarButtonEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CommandBarButtonEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CommandBarButtonEvents_Event : _CommandBarButtonEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CommandBarButton which adds IDispose to the interface
	/// </summary>
	public interface ICommandBarButton : CommandBarButton, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebPageFont which adds IDispose to the interface
	/// </summary>
	public interface IWebPageFont : WebPageFont, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebPageFonts which adds IDispose to the interface
	/// </summary>
	public interface IWebPageFonts : WebPageFonts, IDisposable { }

	/// <summary>
	/// Wrapper interface for HTMLProjectItem which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProjectItem : HTMLProjectItem, IDisposable { }

	/// <summary>
	/// Wrapper interface for HTMLProjectItems which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProjectItems : HTMLProjectItems, IDisposable { }

	/// <summary>
	/// Wrapper interface for HTMLProject which adds IDispose to the interface
	/// </summary>
	public interface IHTMLProject : HTMLProject, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions : MsoDebugOptions, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialogSelectedItems which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogSelectedItems : FileDialogSelectedItems, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialogFilter which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogFilter : FileDialogFilter, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialogFilters which adds IDispose to the interface
	/// </summary>
	public interface IFileDialogFilters : FileDialogFilters, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileDialog which adds IDispose to the interface
	/// </summary>
	public interface IFileDialog : FileDialog, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureSet which adds IDispose to the interface
	/// </summary>
	public interface ISignatureSet : SignatureSet, IDisposable { }

	/// <summary>
	/// Wrapper interface for Signature which adds IDispose to the interface
	/// </summary>
	public interface ISignature : Signature, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVB which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVB : IMsoEnvelopeVB, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVBEvents which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVBEvents : IMsoEnvelopeVBEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEnvelopeVBEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEnvelopeVBEvents_Event : IMsoEnvelopeVBEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoEnvelope which adds IDispose to the interface
	/// </summary>
	public interface IMsoEnvelope : MsoEnvelope, IDisposable { }

	/// <summary>
	/// Wrapper interface for FileTypes which adds IDispose to the interface
	/// </summary>
	public interface IFileTypes : FileTypes, IDisposable { }

	/// <summary>
	/// Wrapper interface for SearchFolders which adds IDispose to the interface
	/// </summary>
	public interface ISearchFolders : SearchFolders, IDisposable { }

	/// <summary>
	/// Wrapper interface for ScopeFolders which adds IDispose to the interface
	/// </summary>
	public interface IScopeFolders : ScopeFolders, IDisposable { }

	/// <summary>
	/// Wrapper interface for ScopeFolder which adds IDispose to the interface
	/// </summary>
	public interface IScopeFolder : ScopeFolder, IDisposable { }

	/// <summary>
	/// Wrapper interface for SearchScope which adds IDispose to the interface
	/// </summary>
	public interface ISearchScope : SearchScope, IDisposable { }

	/// <summary>
	/// Wrapper interface for SearchScopes which adds IDispose to the interface
	/// </summary>
	public interface ISearchScopes : SearchScopes, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDiagram which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDiagram : IMsoDiagram, IDisposable { }

	/// <summary>
	/// Wrapper interface for DiagramNodes which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodes : DiagramNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodeChildren : DiagramNodeChildren, IDisposable { }

	/// <summary>
	/// Wrapper interface for DiagramNode which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNode : DiagramNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for CanvasShapes which adds IDispose to the interface
	/// </summary>
	public interface ICanvasShapes : CanvasShapes, IDisposable { }

	/// <summary>
	/// Wrapper interface for OfficeDataSourceObject which adds IDispose to the interface
	/// </summary>
	public interface IOfficeDataSourceObject : OfficeDataSourceObject, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOColumn which adds IDispose to the interface
	/// </summary>
	public interface IODSOColumn : ODSOColumn, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOColumns which adds IDispose to the interface
	/// </summary>
	public interface IODSOColumns : ODSOColumns, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOFilter which adds IDispose to the interface
	/// </summary>
	public interface IODSOFilter : ODSOFilter, IDisposable { }

	/// <summary>
	/// Wrapper interface for ODSOFilters which adds IDispose to the interface
	/// </summary>
	public interface IODSOFilters : ODSOFilters, IDisposable { }

	/// <summary>
	/// Wrapper interface for NewFile which adds IDispose to the interface
	/// </summary>
	public interface INewFile : NewFile, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponent which adds IDispose to the interface
	/// </summary>
	public interface IWebComponent : WebComponent, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponentWindowExternal which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentWindowExternal : WebComponentWindowExternal, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponentFormat which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentFormat : WebComponentFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ILicWizExternal which adds IDispose to the interface
	/// </summary>
	public interface IILicWizExternal : ILicWizExternal, IDisposable { }

	/// <summary>
	/// Wrapper interface for ILicValidator which adds IDispose to the interface
	/// </summary>
	public interface IILicValidator : ILicValidator, IDisposable { }

	/// <summary>
	/// Wrapper interface for ILicAgent which adds IDispose to the interface
	/// </summary>
	public interface IILicAgent : ILicAgent, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoEServicesDialog which adds IDispose to the interface
	/// </summary>
	public interface IIMsoEServicesDialog : IMsoEServicesDialog, IDisposable { }

	/// <summary>
	/// Wrapper interface for WebComponentProperties which adds IDispose to the interface
	/// </summary>
	public interface IWebComponentProperties : WebComponentProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartDocument which adds IDispose to the interface
	/// </summary>
	public interface ISmartDocument : SmartDocument, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceMember which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceMember : SharedWorkspaceMember, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceMembers which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceMembers : SharedWorkspaceMembers, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceTask which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceTask : SharedWorkspaceTask, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceTasks which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceTasks : SharedWorkspaceTasks, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFile which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFile : SharedWorkspaceFile, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFiles which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFiles : SharedWorkspaceFiles, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFolder which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFolder : SharedWorkspaceFolder, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceFolders which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceFolders : SharedWorkspaceFolders, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceLink which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceLink : SharedWorkspaceLink, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspaceLinks which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspaceLinks : SharedWorkspaceLinks, IDisposable { }

	/// <summary>
	/// Wrapper interface for SharedWorkspace which adds IDispose to the interface
	/// </summary>
	public interface ISharedWorkspace : SharedWorkspace, IDisposable { }

	/// <summary>
	/// Wrapper interface for Sync which adds IDispose to the interface
	/// </summary>
	public interface ISync : Sync, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentLibraryVersion which adds IDispose to the interface
	/// </summary>
	public interface IDocumentLibraryVersion : DocumentLibraryVersion, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentLibraryVersions which adds IDispose to the interface
	/// </summary>
	public interface IDocumentLibraryVersions : DocumentLibraryVersions, IDisposable { }

	/// <summary>
	/// Wrapper interface for UserPermission which adds IDispose to the interface
	/// </summary>
	public interface IUserPermission : UserPermission, IDisposable { }

	/// <summary>
	/// Wrapper interface for Permission which adds IDispose to the interface
	/// </summary>
	public interface IPermission : Permission, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTRunResult which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTRunResult : MsoDebugOptions_UTRunResult, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UT which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UT : MsoDebugOptions_UT, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTs which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTs : MsoDebugOptions_UTs, IDisposable { }

	/// <summary>
	/// Wrapper interface for MsoDebugOptions_UTManager which adds IDispose to the interface
	/// </summary>
	public interface IMsoDebugOptions_UTManager : MsoDebugOptions_UTManager, IDisposable { }

	/// <summary>
	/// Wrapper interface for MetaProperty which adds IDispose to the interface
	/// </summary>
	public interface IMetaProperty : MetaProperty, IDisposable { }

	/// <summary>
	/// Wrapper interface for MetaProperties which adds IDispose to the interface
	/// </summary>
	public interface IMetaProperties : MetaProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for PolicyItem which adds IDispose to the interface
	/// </summary>
	public interface IPolicyItem : PolicyItem, IDisposable { }

	/// <summary>
	/// Wrapper interface for ServerPolicy which adds IDispose to the interface
	/// </summary>
	public interface IServerPolicy : ServerPolicy, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentInspector which adds IDispose to the interface
	/// </summary>
	public interface IDocumentInspector : DocumentInspector, IDisposable { }

	/// <summary>
	/// Wrapper interface for DocumentInspectors which adds IDispose to the interface
	/// </summary>
	public interface IDocumentInspectors : DocumentInspectors, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTask which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTask : WorkflowTask, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTasks which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTasks : WorkflowTasks, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTemplate which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTemplate : WorkflowTemplate, IDisposable { }

	/// <summary>
	/// Wrapper interface for WorkflowTemplates which adds IDispose to the interface
	/// </summary>
	public interface IWorkflowTemplates : WorkflowTemplates, IDisposable { }

	/// <summary>
	/// Wrapper interface for IDocumentInspector which adds IDispose to the interface
	/// </summary>
	public interface IIDocumentInspector : IDocumentInspector, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureSetup which adds IDispose to the interface
	/// </summary>
	public interface ISignatureSetup : SignatureSetup, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureInfo which adds IDispose to the interface
	/// </summary>
	public interface ISignatureInfo : SignatureInfo, IDisposable { }

	/// <summary>
	/// Wrapper interface for SignatureProvider which adds IDispose to the interface
	/// </summary>
	public interface ISignatureProvider : SignatureProvider, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLPrefixMapping which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPrefixMapping : CustomXMLPrefixMapping, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLPrefixMappings which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPrefixMappings : CustomXMLPrefixMappings, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLSchema which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLSchema : CustomXMLSchema, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLSchemaCollection which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLSchemaCollection : _CustomXMLSchemaCollection, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLSchemaCollection which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLSchemaCollection : CustomXMLSchemaCollection, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLNodes which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLNodes : CustomXMLNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLNode which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLNode : CustomXMLNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLValidationError which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLValidationError : CustomXMLValidationError, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLValidationErrors which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLValidationErrors : CustomXMLValidationErrors, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPart which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPart : _CustomXMLPart, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICustomXMLPartEvents which adds IDispose to the interface
	/// </summary>
	public interface IICustomXMLPartEvents : ICustomXMLPartEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartEvents : _CustomXMLPartEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartEvents_Event : _CustomXMLPartEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLPart which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLPart : CustomXMLPart, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLParts which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLParts : _CustomXMLParts, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICustomXMLPartsEvents which adds IDispose to the interface
	/// </summary>
	public interface IICustomXMLPartsEvents : ICustomXMLPartsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartsEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartsEvents : _CustomXMLPartsEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomXMLPartsEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomXMLPartsEvents_Event : _CustomXMLPartsEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomXMLParts which adds IDispose to the interface
	/// </summary>
	public interface ICustomXMLParts : CustomXMLParts, IDisposable { }

	/// <summary>
	/// Wrapper interface for GradientStop which adds IDispose to the interface
	/// </summary>
	public interface IGradientStop : GradientStop, IDisposable { }

	/// <summary>
	/// Wrapper interface for GradientStops which adds IDispose to the interface
	/// </summary>
	public interface IGradientStops : GradientStops, IDisposable { }

	/// <summary>
	/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
	/// </summary>
	public interface ISoftEdgeFormat : SoftEdgeFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for GlowFormat which adds IDispose to the interface
	/// </summary>
	public interface IGlowFormat : GlowFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
	/// </summary>
	public interface IReflectionFormat : ReflectionFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for ParagraphFormat2 which adds IDispose to the interface
	/// </summary>
	public interface IParagraphFormat2 : ParagraphFormat2, IDisposable { }

	/// <summary>
	/// Wrapper interface for Font2 which adds IDispose to the interface
	/// </summary>
	public interface IFont2 : Font2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextColumn2 which adds IDispose to the interface
	/// </summary>
	public interface ITextColumn2 : TextColumn2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextRange2 which adds IDispose to the interface
	/// </summary>
	public interface ITextRange2 : TextRange2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TextFrame2 which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame2 : TextFrame2, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeColor which adds IDispose to the interface
	/// </summary>
	public interface IThemeColor : ThemeColor, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeColorScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeColorScheme : ThemeColorScheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeFont which adds IDispose to the interface
	/// </summary>
	public interface IThemeFont : ThemeFont, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeFonts which adds IDispose to the interface
	/// </summary>
	public interface IThemeFonts : ThemeFonts, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeFontScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeFontScheme : ThemeFontScheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for ThemeEffectScheme which adds IDispose to the interface
	/// </summary>
	public interface IThemeEffectScheme : ThemeEffectScheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for OfficeTheme which adds IDispose to the interface
	/// </summary>
	public interface IOfficeTheme : OfficeTheme, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomTaskPane which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPane : _CustomTaskPane, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomTaskPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface ICustomTaskPaneEvents : CustomTaskPaneEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomTaskPaneEvents which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPaneEvents : _CustomTaskPaneEvents, IDisposable { }

	/// <summary>
	/// Wrapper interface for _CustomTaskPaneEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface I_CustomTaskPaneEvents_Event : _CustomTaskPaneEvents_Event, IDisposable { }

	/// <summary>
	/// Wrapper interface for CustomTaskPane which adds IDispose to the interface
	/// </summary>
	public interface ICustomTaskPane : CustomTaskPane, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICTPFactory which adds IDispose to the interface
	/// </summary>
	public interface IICTPFactory : ICTPFactory, IDisposable { }

	/// <summary>
	/// Wrapper interface for ICustomTaskPaneConsumer which adds IDispose to the interface
	/// </summary>
	public interface IICustomTaskPaneConsumer : ICustomTaskPaneConsumer, IDisposable { }

	/// <summary>
	/// Wrapper interface for IRibbonUI which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonUI : IRibbonUI, IDisposable { }

	/// <summary>
	/// Wrapper interface for IRibbonControl which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonControl : IRibbonControl, IDisposable { }

	/// <summary>
	/// Wrapper interface for IRibbonExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIRibbonExtensibility : IRibbonExtensibility, IDisposable { }

	/// <summary>
	/// Wrapper interface for IAssistance which adds IDispose to the interface
	/// </summary>
	public interface IIAssistance : IAssistance, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartData which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartData : IMsoChartData, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChart which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChart : IMsoChart, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoCorners which adds IDispose to the interface
	/// </summary>
	public interface IIMsoCorners : IMsoCorners, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoLegend which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLegend : IMsoLegend, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoBorder which adds IDispose to the interface
	/// </summary>
	public interface IIMsoBorder : IMsoBorder, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoWalls which adds IDispose to the interface
	/// </summary>
	public interface IIMsoWalls : IMsoWalls, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoFloor which adds IDispose to the interface
	/// </summary>
	public interface IIMsoFloor : IMsoFloor, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoPlotArea which adds IDispose to the interface
	/// </summary>
	public interface IIMsoPlotArea : IMsoPlotArea, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartArea which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartArea : IMsoChartArea, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoSeriesLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoSeriesLines : IMsoSeriesLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoLeaderLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLeaderLines : IMsoLeaderLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for GridLines which adds IDispose to the interface
	/// </summary>
	public interface IGridLines : GridLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoUpBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoUpBars : IMsoUpBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDownBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDownBars : IMsoDownBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoInterior which adds IDispose to the interface
	/// </summary>
	public interface IIMsoInterior : IMsoInterior, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFillFormat : ChartFillFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for LegendEntries which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntries : LegendEntries, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartFont which adds IDispose to the interface
	/// </summary>
	public interface IChartFont : ChartFont, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartColorFormat : ChartColorFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for LegendEntry which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntry : LegendEntry, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoLegendKey which adds IDispose to the interface
	/// </summary>
	public interface IIMsoLegendKey : IMsoLegendKey, IDisposable { }

	/// <summary>
	/// Wrapper interface for SeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface ISeriesCollection : SeriesCollection, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoSeries which adds IDispose to the interface
	/// </summary>
	public interface IIMsoSeries : IMsoSeries, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IIMsoErrorBars : IMsoErrorBars, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoTrendline which adds IDispose to the interface
	/// </summary>
	public interface IIMsoTrendline : IMsoTrendline, IDisposable { }

	/// <summary>
	/// Wrapper interface for Trendlines which adds IDispose to the interface
	/// </summary>
	public interface ITrendlines : Trendlines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDataLabels which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataLabels : IMsoDataLabels, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDataLabel which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataLabel : IMsoDataLabel, IDisposable { }

	/// <summary>
	/// Wrapper interface for Points which adds IDispose to the interface
	/// </summary>
	public interface IPoints : Points, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartPoint which adds IDispose to the interface
	/// </summary>
	public interface IChartPoint : ChartPoint, IDisposable { }

	/// <summary>
	/// Wrapper interface for Axes which adds IDispose to the interface
	/// </summary>
	public interface IAxes : Axes, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoAxis which adds IDispose to the interface
	/// </summary>
	public interface IIMsoAxis : IMsoAxis, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDataTable which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDataTable : IMsoDataTable, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartTitle : IMsoChartTitle, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoAxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IIMsoAxisTitle : IMsoAxisTitle, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDisplayUnitLabel : IMsoDisplayUnitLabel, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoTickLabels which adds IDispose to the interface
	/// </summary>
	public interface IIMsoTickLabels : IMsoTickLabels, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoHyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IIMsoHyperlinks : IMsoHyperlinks, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoDropLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoDropLines : IMsoDropLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoHiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IIMsoHiLoLines : IMsoHiLoLines, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartGroup : IMsoChartGroup, IDisposable { }

	/// <summary>
	/// Wrapper interface for ChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IChartGroups : ChartGroups, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoCharacters which adds IDispose to the interface
	/// </summary>
	public interface IIMsoCharacters : IMsoCharacters, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IIMsoChartFormat : IMsoChartFormat, IDisposable { }

	/// <summary>
	/// Wrapper interface for BulletFormat2 which adds IDispose to the interface
	/// </summary>
	public interface IBulletFormat2 : BulletFormat2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TabStops2 which adds IDispose to the interface
	/// </summary>
	public interface ITabStops2 : TabStops2, IDisposable { }

	/// <summary>
	/// Wrapper interface for TabStop2 which adds IDispose to the interface
	/// </summary>
	public interface ITabStop2 : TabStop2, IDisposable { }

	/// <summary>
	/// Wrapper interface for Ruler2 which adds IDispose to the interface
	/// </summary>
	public interface IRuler2 : Ruler2, IDisposable { }

	/// <summary>
	/// Wrapper interface for RulerLevels2 which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevels2 : RulerLevels2, IDisposable { }

	/// <summary>
	/// Wrapper interface for RulerLevel2 which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevel2 : RulerLevel2, IDisposable { }

	/// <summary>
	/// Wrapper interface for EncryptionProvider which adds IDispose to the interface
	/// </summary>
	public interface IEncryptionProvider : EncryptionProvider, IDisposable { }

	/// <summary>
	/// Wrapper interface for IBlogExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIBlogExtensibility : IBlogExtensibility, IDisposable { }

	/// <summary>
	/// Wrapper interface for IBlogPictureExtensibility which adds IDispose to the interface
	/// </summary>
	public interface IIBlogPictureExtensibility : IBlogPictureExtensibility, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverterPreferences which adds IDispose to the interface
	/// </summary>
	public interface IIConverterPreferences : IConverterPreferences, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverterApplicationPreferences which adds IDispose to the interface
	/// </summary>
	public interface IIConverterApplicationPreferences : IConverterApplicationPreferences, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverterUICallback which adds IDispose to the interface
	/// </summary>
	public interface IIConverterUICallback : IConverterUICallback, IDisposable { }

	/// <summary>
	/// Wrapper interface for IConverter which adds IDispose to the interface
	/// </summary>
	public interface IIConverter : IConverter, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArt which adds IDispose to the interface
	/// </summary>
	public interface ISmartArt : SmartArt, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtNodes which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtNodes : SmartArtNodes, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtNode which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtNode : SmartArtNode, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtLayouts which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtLayouts : SmartArtLayouts, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtLayout which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtLayout : SmartArtLayout, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtQuickStyles which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtQuickStyles : SmartArtQuickStyles, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtQuickStyle which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtQuickStyle : SmartArtQuickStyle, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtColors which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtColors : SmartArtColors, IDisposable { }

	/// <summary>
	/// Wrapper interface for SmartArtColor which adds IDispose to the interface
	/// </summary>
	public interface ISmartArtColor : SmartArtColor, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerField which adds IDispose to the interface
	/// </summary>
	public interface IPickerField : PickerField, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerFields which adds IDispose to the interface
	/// </summary>
	public interface IPickerFields : PickerFields, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerProperty which adds IDispose to the interface
	/// </summary>
	public interface IPickerProperty : PickerProperty, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerProperties which adds IDispose to the interface
	/// </summary>
	public interface IPickerProperties : PickerProperties, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerResult which adds IDispose to the interface
	/// </summary>
	public interface IPickerResult : PickerResult, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerResults which adds IDispose to the interface
	/// </summary>
	public interface IPickerResults : PickerResults, IDisposable { }

	/// <summary>
	/// Wrapper interface for PickerDialog which adds IDispose to the interface
	/// </summary>
	public interface IPickerDialog : PickerDialog, IDisposable { }

	/// <summary>
	/// Wrapper interface for IMsoContactCard which adds IDispose to the interface
	/// </summary>
	public interface IIMsoContactCard : IMsoContactCard, IDisposable { }

	/// <summary>
	/// Wrapper interface for EffectParameter which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameter : EffectParameter, IDisposable { }

	/// <summary>
	/// Wrapper interface for EffectParameters which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameters : EffectParameters, IDisposable { }

	/// <summary>
	/// Wrapper interface for PictureEffect which adds IDispose to the interface
	/// </summary>
	public interface IPictureEffect : PictureEffect, IDisposable { }

	/// <summary>
	/// Wrapper interface for PictureEffects which adds IDispose to the interface
	/// </summary>
	public interface IPictureEffects : PictureEffects, IDisposable { }

	/// <summary>
	/// Wrapper interface for Crop which adds IDispose to the interface
	/// </summary>
	public interface ICrop : Crop, IDisposable { }

	/// <summary>
	/// Wrapper interface for ContactCard which adds IDispose to the interface
	/// </summary>
	public interface IContactCard : ContactCard, IDisposable { }

	}