using Office.Utility.Proxies;

namespace Office.Utility.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeProxyExtensions
	{
		/// <summary>
		/// Wrapper Proxy for IAccessible COM interface which is disposible
		/// </summary>
		public static IAccessibleProxy ToProxy(this Microsoft.Office.Core.IAccessible resource)
		{
			return new IAccessibleProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _IMsoDispObj COM interface which is disposible
		/// </summary>
		public static _IMsoDispObjProxy ToProxy(this Microsoft.Office.Core._IMsoDispObj resource)
		{
			return new _IMsoDispObjProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _IMsoOleAccDispObj COM interface which is disposible
		/// </summary>
		public static _IMsoOleAccDispObjProxy ToProxy(this Microsoft.Office.Core._IMsoOleAccDispObj resource)
		{
			return new _IMsoOleAccDispObjProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBars COM interface which is disposible
		/// </summary>
		public static _CommandBarsProxy ToProxy(this Microsoft.Office.Core._CommandBars resource)
		{
			return new _CommandBarsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBar COM interface which is disposible
		/// </summary>
		public static CommandBarProxy ToProxy(this Microsoft.Office.Core.CommandBar resource)
		{
			return new CommandBarProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBarControls COM interface which is disposible
		/// </summary>
		public static CommandBarControlsProxy ToProxy(this Microsoft.Office.Core.CommandBarControls resource)
		{
			return new CommandBarControlsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBarControl COM interface which is disposible
		/// </summary>
		public static CommandBarControlProxy ToProxy(this Microsoft.Office.Core.CommandBarControl resource)
		{
			return new CommandBarControlProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarButton COM interface which is disposible
		/// </summary>
		public static _CommandBarButtonProxy ToProxy(this Microsoft.Office.Core._CommandBarButton resource)
		{
			return new _CommandBarButtonProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBarPopup COM interface which is disposible
		/// </summary>
		public static CommandBarPopupProxy ToProxy(this Microsoft.Office.Core.CommandBarPopup resource)
		{
			return new CommandBarPopupProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarComboBox COM interface which is disposible
		/// </summary>
		public static _CommandBarComboBoxProxy ToProxy(this Microsoft.Office.Core._CommandBarComboBox resource)
		{
			return new _CommandBarComboBoxProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarActiveX COM interface which is disposible
		/// </summary>
		public static _CommandBarActiveXProxy ToProxy(this Microsoft.Office.Core._CommandBarActiveX resource)
		{
			return new _CommandBarActiveXProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Adjustments COM interface which is disposible
		/// </summary>
		public static AdjustmentsProxy ToProxy(this Microsoft.Office.Core.Adjustments resource)
		{
			return new AdjustmentsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CalloutFormat COM interface which is disposible
		/// </summary>
		public static CalloutFormatProxy ToProxy(this Microsoft.Office.Core.CalloutFormat resource)
		{
			return new CalloutFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ColorFormat COM interface which is disposible
		/// </summary>
		public static ColorFormatProxy ToProxy(this Microsoft.Office.Core.ColorFormat resource)
		{
			return new ColorFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ConnectorFormat COM interface which is disposible
		/// </summary>
		public static ConnectorFormatProxy ToProxy(this Microsoft.Office.Core.ConnectorFormat resource)
		{
			return new ConnectorFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FillFormat COM interface which is disposible
		/// </summary>
		public static FillFormatProxy ToProxy(this Microsoft.Office.Core.FillFormat resource)
		{
			return new FillFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FreeformBuilder COM interface which is disposible
		/// </summary>
		public static FreeformBuilderProxy ToProxy(this Microsoft.Office.Core.FreeformBuilder resource)
		{
			return new FreeformBuilderProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for GroupShapes COM interface which is disposible
		/// </summary>
		public static GroupShapesProxy ToProxy(this Microsoft.Office.Core.GroupShapes resource)
		{
			return new GroupShapesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for LineFormat COM interface which is disposible
		/// </summary>
		public static LineFormatProxy ToProxy(this Microsoft.Office.Core.LineFormat resource)
		{
			return new LineFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ShapeNode COM interface which is disposible
		/// </summary>
		public static ShapeNodeProxy ToProxy(this Microsoft.Office.Core.ShapeNode resource)
		{
			return new ShapeNodeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ShapeNodes COM interface which is disposible
		/// </summary>
		public static ShapeNodesProxy ToProxy(this Microsoft.Office.Core.ShapeNodes resource)
		{
			return new ShapeNodesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PictureFormat COM interface which is disposible
		/// </summary>
		public static PictureFormatProxy ToProxy(this Microsoft.Office.Core.PictureFormat resource)
		{
			return new PictureFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ShadowFormat COM interface which is disposible
		/// </summary>
		public static ShadowFormatProxy ToProxy(this Microsoft.Office.Core.ShadowFormat resource)
		{
			return new ShadowFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Script COM interface which is disposible
		/// </summary>
		public static ScriptProxy ToProxy(this Microsoft.Office.Core.Script resource)
		{
			return new ScriptProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Scripts COM interface which is disposible
		/// </summary>
		public static ScriptsProxy ToProxy(this Microsoft.Office.Core.Scripts resource)
		{
			return new ScriptsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Shape COM interface which is disposible
		/// </summary>
		public static ShapeProxy ToProxy(this Microsoft.Office.Core.Shape resource)
		{
			return new ShapeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ShapeRange COM interface which is disposible
		/// </summary>
		public static ShapeRangeProxy ToProxy(this Microsoft.Office.Core.ShapeRange resource)
		{
			return new ShapeRangeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Shapes COM interface which is disposible
		/// </summary>
		public static ShapesProxy ToProxy(this Microsoft.Office.Core.Shapes resource)
		{
			return new ShapesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TextEffectFormat COM interface which is disposible
		/// </summary>
		public static TextEffectFormatProxy ToProxy(this Microsoft.Office.Core.TextEffectFormat resource)
		{
			return new TextEffectFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TextFrame COM interface which is disposible
		/// </summary>
		public static TextFrameProxy ToProxy(this Microsoft.Office.Core.TextFrame resource)
		{
			return new TextFrameProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThreeDFormat COM interface which is disposible
		/// </summary>
		public static ThreeDFormatProxy ToProxy(this Microsoft.Office.Core.ThreeDFormat resource)
		{
			return new ThreeDFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDispCagNotifySink COM interface which is disposible
		/// </summary>
		public static IMsoDispCagNotifySinkProxy ToProxy(this Microsoft.Office.Core.IMsoDispCagNotifySink resource)
		{
			return new IMsoDispCagNotifySinkProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Balloon COM interface which is disposible
		/// </summary>
		public static BalloonProxy ToProxy(this Microsoft.Office.Core.Balloon resource)
		{
			return new BalloonProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for BalloonCheckboxes COM interface which is disposible
		/// </summary>
		public static BalloonCheckboxesProxy ToProxy(this Microsoft.Office.Core.BalloonCheckboxes resource)
		{
			return new BalloonCheckboxesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for BalloonCheckbox COM interface which is disposible
		/// </summary>
		public static BalloonCheckboxProxy ToProxy(this Microsoft.Office.Core.BalloonCheckbox resource)
		{
			return new BalloonCheckboxProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for BalloonLabels COM interface which is disposible
		/// </summary>
		public static BalloonLabelsProxy ToProxy(this Microsoft.Office.Core.BalloonLabels resource)
		{
			return new BalloonLabelsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for BalloonLabel COM interface which is disposible
		/// </summary>
		public static BalloonLabelProxy ToProxy(this Microsoft.Office.Core.BalloonLabel resource)
		{
			return new BalloonLabelProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for AnswerWizardFiles COM interface which is disposible
		/// </summary>
		public static AnswerWizardFilesProxy ToProxy(this Microsoft.Office.Core.AnswerWizardFiles resource)
		{
			return new AnswerWizardFilesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for AnswerWizard COM interface which is disposible
		/// </summary>
		public static AnswerWizardProxy ToProxy(this Microsoft.Office.Core.AnswerWizard resource)
		{
			return new AnswerWizardProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Assistant COM interface which is disposible
		/// </summary>
		public static AssistantProxy ToProxy(this Microsoft.Office.Core.Assistant resource)
		{
			return new AssistantProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DocumentProperty COM interface which is disposible
		/// </summary>
		public static DocumentPropertyProxy ToProxy(this Microsoft.Office.Core.DocumentProperty resource)
		{
			return new DocumentPropertyProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DocumentProperties COM interface which is disposible
		/// </summary>
		public static DocumentPropertiesProxy ToProxy(this Microsoft.Office.Core.DocumentProperties resource)
		{
			return new DocumentPropertiesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IFoundFiles COM interface which is disposible
		/// </summary>
		public static IFoundFilesProxy ToProxy(this Microsoft.Office.Core.IFoundFiles resource)
		{
			return new IFoundFilesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IFind COM interface which is disposible
		/// </summary>
		public static IFindProxy ToProxy(this Microsoft.Office.Core.IFind resource)
		{
			return new IFindProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FoundFiles COM interface which is disposible
		/// </summary>
		public static FoundFilesProxy ToProxy(this Microsoft.Office.Core.FoundFiles resource)
		{
			return new FoundFilesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PropertyTest COM interface which is disposible
		/// </summary>
		public static PropertyTestProxy ToProxy(this Microsoft.Office.Core.PropertyTest resource)
		{
			return new PropertyTestProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PropertyTests COM interface which is disposible
		/// </summary>
		public static PropertyTestsProxy ToProxy(this Microsoft.Office.Core.PropertyTests resource)
		{
			return new PropertyTestsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FileSearch COM interface which is disposible
		/// </summary>
		public static FileSearchProxy ToProxy(this Microsoft.Office.Core.FileSearch resource)
		{
			return new FileSearchProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for COMAddIn COM interface which is disposible
		/// </summary>
		public static COMAddInProxy ToProxy(this Microsoft.Office.Core.COMAddIn resource)
		{
			return new COMAddInProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for COMAddIns COM interface which is disposible
		/// </summary>
		public static COMAddInsProxy ToProxy(this Microsoft.Office.Core.COMAddIns resource)
		{
			return new COMAddInsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for LanguageSettings COM interface which is disposible
		/// </summary>
		public static LanguageSettingsProxy ToProxy(this Microsoft.Office.Core.LanguageSettings resource)
		{
			return new LanguageSettingsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICommandBarsEvents COM interface which is disposible
		/// </summary>
		public static ICommandBarsEventsProxy ToProxy(this Microsoft.Office.Core.ICommandBarsEvents resource)
		{
			return new ICommandBarsEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarsEvents COM interface which is disposible
		/// </summary>
		public static _CommandBarsEventsProxy ToProxy(this Microsoft.Office.Core._CommandBarsEvents resource)
		{
			return new _CommandBarsEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarsEvents_Event COM interface which is disposible
		/// </summary>
		public static _CommandBarsEvents_EventProxy ToProxy(this Microsoft.Office.Core._CommandBarsEvents_Event resource)
		{
			return new _CommandBarsEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBars COM interface which is disposible
		/// </summary>
		public static CommandBarsProxy ToProxy(this Microsoft.Office.Core.CommandBars resource)
		{
			return new CommandBarsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICommandBarComboBoxEvents COM interface which is disposible
		/// </summary>
		public static ICommandBarComboBoxEventsProxy ToProxy(this Microsoft.Office.Core.ICommandBarComboBoxEvents resource)
		{
			return new ICommandBarComboBoxEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarComboBoxEvents COM interface which is disposible
		/// </summary>
		public static _CommandBarComboBoxEventsProxy ToProxy(this Microsoft.Office.Core._CommandBarComboBoxEvents resource)
		{
			return new _CommandBarComboBoxEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarComboBoxEvents_Event COM interface which is disposible
		/// </summary>
		public static _CommandBarComboBoxEvents_EventProxy ToProxy(this Microsoft.Office.Core._CommandBarComboBoxEvents_Event resource)
		{
			return new _CommandBarComboBoxEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBarComboBox COM interface which is disposible
		/// </summary>
		public static CommandBarComboBoxProxy ToProxy(this Microsoft.Office.Core.CommandBarComboBox resource)
		{
			return new CommandBarComboBoxProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICommandBarButtonEvents COM interface which is disposible
		/// </summary>
		public static ICommandBarButtonEventsProxy ToProxy(this Microsoft.Office.Core.ICommandBarButtonEvents resource)
		{
			return new ICommandBarButtonEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarButtonEvents COM interface which is disposible
		/// </summary>
		public static _CommandBarButtonEventsProxy ToProxy(this Microsoft.Office.Core._CommandBarButtonEvents resource)
		{
			return new _CommandBarButtonEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CommandBarButtonEvents_Event COM interface which is disposible
		/// </summary>
		public static _CommandBarButtonEvents_EventProxy ToProxy(this Microsoft.Office.Core._CommandBarButtonEvents_Event resource)
		{
			return new _CommandBarButtonEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CommandBarButton COM interface which is disposible
		/// </summary>
		public static CommandBarButtonProxy ToProxy(this Microsoft.Office.Core.CommandBarButton resource)
		{
			return new CommandBarButtonProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WebPageFont COM interface which is disposible
		/// </summary>
		public static WebPageFontProxy ToProxy(this Microsoft.Office.Core.WebPageFont resource)
		{
			return new WebPageFontProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WebPageFonts COM interface which is disposible
		/// </summary>
		public static WebPageFontsProxy ToProxy(this Microsoft.Office.Core.WebPageFonts resource)
		{
			return new WebPageFontsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for HTMLProjectItem COM interface which is disposible
		/// </summary>
		public static HTMLProjectItemProxy ToProxy(this Microsoft.Office.Core.HTMLProjectItem resource)
		{
			return new HTMLProjectItemProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for HTMLProjectItems COM interface which is disposible
		/// </summary>
		public static HTMLProjectItemsProxy ToProxy(this Microsoft.Office.Core.HTMLProjectItems resource)
		{
			return new HTMLProjectItemsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for HTMLProject COM interface which is disposible
		/// </summary>
		public static HTMLProjectProxy ToProxy(this Microsoft.Office.Core.HTMLProject resource)
		{
			return new HTMLProjectProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MsoDebugOptions COM interface which is disposible
		/// </summary>
		public static MsoDebugOptionsProxy ToProxy(this Microsoft.Office.Core.MsoDebugOptions resource)
		{
			return new MsoDebugOptionsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FileDialogSelectedItems COM interface which is disposible
		/// </summary>
		public static FileDialogSelectedItemsProxy ToProxy(this Microsoft.Office.Core.FileDialogSelectedItems resource)
		{
			return new FileDialogSelectedItemsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FileDialogFilter COM interface which is disposible
		/// </summary>
		public static FileDialogFilterProxy ToProxy(this Microsoft.Office.Core.FileDialogFilter resource)
		{
			return new FileDialogFilterProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FileDialogFilters COM interface which is disposible
		/// </summary>
		public static FileDialogFiltersProxy ToProxy(this Microsoft.Office.Core.FileDialogFilters resource)
		{
			return new FileDialogFiltersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FileDialog COM interface which is disposible
		/// </summary>
		public static FileDialogProxy ToProxy(this Microsoft.Office.Core.FileDialog resource)
		{
			return new FileDialogProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SignatureSet COM interface which is disposible
		/// </summary>
		public static SignatureSetProxy ToProxy(this Microsoft.Office.Core.SignatureSet resource)
		{
			return new SignatureSetProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Signature COM interface which is disposible
		/// </summary>
		public static SignatureProxy ToProxy(this Microsoft.Office.Core.Signature resource)
		{
			return new SignatureProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoEnvelopeVB COM interface which is disposible
		/// </summary>
		public static IMsoEnvelopeVBProxy ToProxy(this Microsoft.Office.Core.IMsoEnvelopeVB resource)
		{
			return new IMsoEnvelopeVBProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoEnvelopeVBEvents COM interface which is disposible
		/// </summary>
		public static IMsoEnvelopeVBEventsProxy ToProxy(this Microsoft.Office.Core.IMsoEnvelopeVBEvents resource)
		{
			return new IMsoEnvelopeVBEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoEnvelopeVBEvents_Event COM interface which is disposible
		/// </summary>
		public static IMsoEnvelopeVBEvents_EventProxy ToProxy(this Microsoft.Office.Core.IMsoEnvelopeVBEvents_Event resource)
		{
			return new IMsoEnvelopeVBEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MsoEnvelope COM interface which is disposible
		/// </summary>
		public static MsoEnvelopeProxy ToProxy(this Microsoft.Office.Core.MsoEnvelope resource)
		{
			return new MsoEnvelopeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for FileTypes COM interface which is disposible
		/// </summary>
		public static FileTypesProxy ToProxy(this Microsoft.Office.Core.FileTypes resource)
		{
			return new FileTypesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SearchFolders COM interface which is disposible
		/// </summary>
		public static SearchFoldersProxy ToProxy(this Microsoft.Office.Core.SearchFolders resource)
		{
			return new SearchFoldersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ScopeFolders COM interface which is disposible
		/// </summary>
		public static ScopeFoldersProxy ToProxy(this Microsoft.Office.Core.ScopeFolders resource)
		{
			return new ScopeFoldersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ScopeFolder COM interface which is disposible
		/// </summary>
		public static ScopeFolderProxy ToProxy(this Microsoft.Office.Core.ScopeFolder resource)
		{
			return new ScopeFolderProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SearchScope COM interface which is disposible
		/// </summary>
		public static SearchScopeProxy ToProxy(this Microsoft.Office.Core.SearchScope resource)
		{
			return new SearchScopeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SearchScopes COM interface which is disposible
		/// </summary>
		public static SearchScopesProxy ToProxy(this Microsoft.Office.Core.SearchScopes resource)
		{
			return new SearchScopesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDiagram COM interface which is disposible
		/// </summary>
		public static IMsoDiagramProxy ToProxy(this Microsoft.Office.Core.IMsoDiagram resource)
		{
			return new IMsoDiagramProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DiagramNodes COM interface which is disposible
		/// </summary>
		public static DiagramNodesProxy ToProxy(this Microsoft.Office.Core.DiagramNodes resource)
		{
			return new DiagramNodesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DiagramNodeChildren COM interface which is disposible
		/// </summary>
		public static DiagramNodeChildrenProxy ToProxy(this Microsoft.Office.Core.DiagramNodeChildren resource)
		{
			return new DiagramNodeChildrenProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DiagramNode COM interface which is disposible
		/// </summary>
		public static DiagramNodeProxy ToProxy(this Microsoft.Office.Core.DiagramNode resource)
		{
			return new DiagramNodeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CanvasShapes COM interface which is disposible
		/// </summary>
		public static CanvasShapesProxy ToProxy(this Microsoft.Office.Core.CanvasShapes resource)
		{
			return new CanvasShapesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for OfficeDataSourceObject COM interface which is disposible
		/// </summary>
		public static OfficeDataSourceObjectProxy ToProxy(this Microsoft.Office.Core.OfficeDataSourceObject resource)
		{
			return new OfficeDataSourceObjectProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ODSOColumn COM interface which is disposible
		/// </summary>
		public static ODSOColumnProxy ToProxy(this Microsoft.Office.Core.ODSOColumn resource)
		{
			return new ODSOColumnProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ODSOColumns COM interface which is disposible
		/// </summary>
		public static ODSOColumnsProxy ToProxy(this Microsoft.Office.Core.ODSOColumns resource)
		{
			return new ODSOColumnsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ODSOFilter COM interface which is disposible
		/// </summary>
		public static ODSOFilterProxy ToProxy(this Microsoft.Office.Core.ODSOFilter resource)
		{
			return new ODSOFilterProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ODSOFilters COM interface which is disposible
		/// </summary>
		public static ODSOFiltersProxy ToProxy(this Microsoft.Office.Core.ODSOFilters resource)
		{
			return new ODSOFiltersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for NewFile COM interface which is disposible
		/// </summary>
		public static NewFileProxy ToProxy(this Microsoft.Office.Core.NewFile resource)
		{
			return new NewFileProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WebComponent COM interface which is disposible
		/// </summary>
		public static WebComponentProxy ToProxy(this Microsoft.Office.Core.WebComponent resource)
		{
			return new WebComponentProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WebComponentWindowExternal COM interface which is disposible
		/// </summary>
		public static WebComponentWindowExternalProxy ToProxy(this Microsoft.Office.Core.WebComponentWindowExternal resource)
		{
			return new WebComponentWindowExternalProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WebComponentFormat COM interface which is disposible
		/// </summary>
		public static WebComponentFormatProxy ToProxy(this Microsoft.Office.Core.WebComponentFormat resource)
		{
			return new WebComponentFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ILicWizExternal COM interface which is disposible
		/// </summary>
		public static ILicWizExternalProxy ToProxy(this Microsoft.Office.Core.ILicWizExternal resource)
		{
			return new ILicWizExternalProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ILicValidator COM interface which is disposible
		/// </summary>
		public static ILicValidatorProxy ToProxy(this Microsoft.Office.Core.ILicValidator resource)
		{
			return new ILicValidatorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ILicAgent COM interface which is disposible
		/// </summary>
		public static ILicAgentProxy ToProxy(this Microsoft.Office.Core.ILicAgent resource)
		{
			return new ILicAgentProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoEServicesDialog COM interface which is disposible
		/// </summary>
		public static IMsoEServicesDialogProxy ToProxy(this Microsoft.Office.Core.IMsoEServicesDialog resource)
		{
			return new IMsoEServicesDialogProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WebComponentProperties COM interface which is disposible
		/// </summary>
		public static WebComponentPropertiesProxy ToProxy(this Microsoft.Office.Core.WebComponentProperties resource)
		{
			return new WebComponentPropertiesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartDocument COM interface which is disposible
		/// </summary>
		public static SmartDocumentProxy ToProxy(this Microsoft.Office.Core.SmartDocument resource)
		{
			return new SmartDocumentProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceMember COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceMemberProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceMember resource)
		{
			return new SharedWorkspaceMemberProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceMembers COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceMembersProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceMembers resource)
		{
			return new SharedWorkspaceMembersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceTask COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceTaskProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceTask resource)
		{
			return new SharedWorkspaceTaskProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceTasks COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceTasksProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceTasks resource)
		{
			return new SharedWorkspaceTasksProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceFile COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceFileProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceFile resource)
		{
			return new SharedWorkspaceFileProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceFiles COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceFilesProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceFiles resource)
		{
			return new SharedWorkspaceFilesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceFolder COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceFolderProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceFolder resource)
		{
			return new SharedWorkspaceFolderProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceFolders COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceFoldersProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceFolders resource)
		{
			return new SharedWorkspaceFoldersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceLink COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceLinkProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceLink resource)
		{
			return new SharedWorkspaceLinkProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspaceLinks COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceLinksProxy ToProxy(this Microsoft.Office.Core.SharedWorkspaceLinks resource)
		{
			return new SharedWorkspaceLinksProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SharedWorkspace COM interface which is disposible
		/// </summary>
		public static SharedWorkspaceProxy ToProxy(this Microsoft.Office.Core.SharedWorkspace resource)
		{
			return new SharedWorkspaceProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Sync COM interface which is disposible
		/// </summary>
		public static SyncProxy ToProxy(this Microsoft.Office.Core.Sync resource)
		{
			return new SyncProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DocumentLibraryVersion COM interface which is disposible
		/// </summary>
		public static DocumentLibraryVersionProxy ToProxy(this Microsoft.Office.Core.DocumentLibraryVersion resource)
		{
			return new DocumentLibraryVersionProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DocumentLibraryVersions COM interface which is disposible
		/// </summary>
		public static DocumentLibraryVersionsProxy ToProxy(this Microsoft.Office.Core.DocumentLibraryVersions resource)
		{
			return new DocumentLibraryVersionsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for UserPermission COM interface which is disposible
		/// </summary>
		public static UserPermissionProxy ToProxy(this Microsoft.Office.Core.UserPermission resource)
		{
			return new UserPermissionProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Permission COM interface which is disposible
		/// </summary>
		public static PermissionProxy ToProxy(this Microsoft.Office.Core.Permission resource)
		{
			return new PermissionProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MsoDebugOptions_UTRunResult COM interface which is disposible
		/// </summary>
		public static MsoDebugOptions_UTRunResultProxy ToProxy(this Microsoft.Office.Core.MsoDebugOptions_UTRunResult resource)
		{
			return new MsoDebugOptions_UTRunResultProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MsoDebugOptions_UT COM interface which is disposible
		/// </summary>
		public static MsoDebugOptions_UTProxy ToProxy(this Microsoft.Office.Core.MsoDebugOptions_UT resource)
		{
			return new MsoDebugOptions_UTProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MsoDebugOptions_UTs COM interface which is disposible
		/// </summary>
		public static MsoDebugOptions_UTsProxy ToProxy(this Microsoft.Office.Core.MsoDebugOptions_UTs resource)
		{
			return new MsoDebugOptions_UTsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MsoDebugOptions_UTManager COM interface which is disposible
		/// </summary>
		public static MsoDebugOptions_UTManagerProxy ToProxy(this Microsoft.Office.Core.MsoDebugOptions_UTManager resource)
		{
			return new MsoDebugOptions_UTManagerProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MetaProperty COM interface which is disposible
		/// </summary>
		public static MetaPropertyProxy ToProxy(this Microsoft.Office.Core.MetaProperty resource)
		{
			return new MetaPropertyProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for MetaProperties COM interface which is disposible
		/// </summary>
		public static MetaPropertiesProxy ToProxy(this Microsoft.Office.Core.MetaProperties resource)
		{
			return new MetaPropertiesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PolicyItem COM interface which is disposible
		/// </summary>
		public static PolicyItemProxy ToProxy(this Microsoft.Office.Core.PolicyItem resource)
		{
			return new PolicyItemProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ServerPolicy COM interface which is disposible
		/// </summary>
		public static ServerPolicyProxy ToProxy(this Microsoft.Office.Core.ServerPolicy resource)
		{
			return new ServerPolicyProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DocumentInspector COM interface which is disposible
		/// </summary>
		public static DocumentInspectorProxy ToProxy(this Microsoft.Office.Core.DocumentInspector resource)
		{
			return new DocumentInspectorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for DocumentInspectors COM interface which is disposible
		/// </summary>
		public static DocumentInspectorsProxy ToProxy(this Microsoft.Office.Core.DocumentInspectors resource)
		{
			return new DocumentInspectorsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WorkflowTask COM interface which is disposible
		/// </summary>
		public static WorkflowTaskProxy ToProxy(this Microsoft.Office.Core.WorkflowTask resource)
		{
			return new WorkflowTaskProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WorkflowTasks COM interface which is disposible
		/// </summary>
		public static WorkflowTasksProxy ToProxy(this Microsoft.Office.Core.WorkflowTasks resource)
		{
			return new WorkflowTasksProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WorkflowTemplate COM interface which is disposible
		/// </summary>
		public static WorkflowTemplateProxy ToProxy(this Microsoft.Office.Core.WorkflowTemplate resource)
		{
			return new WorkflowTemplateProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for WorkflowTemplates COM interface which is disposible
		/// </summary>
		public static WorkflowTemplatesProxy ToProxy(this Microsoft.Office.Core.WorkflowTemplates resource)
		{
			return new WorkflowTemplatesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IDocumentInspector COM interface which is disposible
		/// </summary>
		public static IDocumentInspectorProxy ToProxy(this Microsoft.Office.Core.IDocumentInspector resource)
		{
			return new IDocumentInspectorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SignatureSetup COM interface which is disposible
		/// </summary>
		public static SignatureSetupProxy ToProxy(this Microsoft.Office.Core.SignatureSetup resource)
		{
			return new SignatureSetupProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SignatureInfo COM interface which is disposible
		/// </summary>
		public static SignatureInfoProxy ToProxy(this Microsoft.Office.Core.SignatureInfo resource)
		{
			return new SignatureInfoProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SignatureProvider COM interface which is disposible
		/// </summary>
		public static SignatureProviderProxy ToProxy(this Microsoft.Office.Core.SignatureProvider resource)
		{
			return new SignatureProviderProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLPrefixMapping COM interface which is disposible
		/// </summary>
		public static CustomXMLPrefixMappingProxy ToProxy(this Microsoft.Office.Core.CustomXMLPrefixMapping resource)
		{
			return new CustomXMLPrefixMappingProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLPrefixMappings COM interface which is disposible
		/// </summary>
		public static CustomXMLPrefixMappingsProxy ToProxy(this Microsoft.Office.Core.CustomXMLPrefixMappings resource)
		{
			return new CustomXMLPrefixMappingsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLSchema COM interface which is disposible
		/// </summary>
		public static CustomXMLSchemaProxy ToProxy(this Microsoft.Office.Core.CustomXMLSchema resource)
		{
			return new CustomXMLSchemaProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLSchemaCollection COM interface which is disposible
		/// </summary>
		public static _CustomXMLSchemaCollectionProxy ToProxy(this Microsoft.Office.Core._CustomXMLSchemaCollection resource)
		{
			return new _CustomXMLSchemaCollectionProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLSchemaCollection COM interface which is disposible
		/// </summary>
		public static CustomXMLSchemaCollectionProxy ToProxy(this Microsoft.Office.Core.CustomXMLSchemaCollection resource)
		{
			return new CustomXMLSchemaCollectionProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLNodes COM interface which is disposible
		/// </summary>
		public static CustomXMLNodesProxy ToProxy(this Microsoft.Office.Core.CustomXMLNodes resource)
		{
			return new CustomXMLNodesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLNode COM interface which is disposible
		/// </summary>
		public static CustomXMLNodeProxy ToProxy(this Microsoft.Office.Core.CustomXMLNode resource)
		{
			return new CustomXMLNodeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLValidationError COM interface which is disposible
		/// </summary>
		public static CustomXMLValidationErrorProxy ToProxy(this Microsoft.Office.Core.CustomXMLValidationError resource)
		{
			return new CustomXMLValidationErrorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLValidationErrors COM interface which is disposible
		/// </summary>
		public static CustomXMLValidationErrorsProxy ToProxy(this Microsoft.Office.Core.CustomXMLValidationErrors resource)
		{
			return new CustomXMLValidationErrorsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLPart COM interface which is disposible
		/// </summary>
		public static _CustomXMLPartProxy ToProxy(this Microsoft.Office.Core._CustomXMLPart resource)
		{
			return new _CustomXMLPartProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICustomXMLPartEvents COM interface which is disposible
		/// </summary>
		public static ICustomXMLPartEventsProxy ToProxy(this Microsoft.Office.Core.ICustomXMLPartEvents resource)
		{
			return new ICustomXMLPartEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLPartEvents COM interface which is disposible
		/// </summary>
		public static _CustomXMLPartEventsProxy ToProxy(this Microsoft.Office.Core._CustomXMLPartEvents resource)
		{
			return new _CustomXMLPartEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLPartEvents_Event COM interface which is disposible
		/// </summary>
		public static _CustomXMLPartEvents_EventProxy ToProxy(this Microsoft.Office.Core._CustomXMLPartEvents_Event resource)
		{
			return new _CustomXMLPartEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLPart COM interface which is disposible
		/// </summary>
		public static CustomXMLPartProxy ToProxy(this Microsoft.Office.Core.CustomXMLPart resource)
		{
			return new CustomXMLPartProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLParts COM interface which is disposible
		/// </summary>
		public static _CustomXMLPartsProxy ToProxy(this Microsoft.Office.Core._CustomXMLParts resource)
		{
			return new _CustomXMLPartsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICustomXMLPartsEvents COM interface which is disposible
		/// </summary>
		public static ICustomXMLPartsEventsProxy ToProxy(this Microsoft.Office.Core.ICustomXMLPartsEvents resource)
		{
			return new ICustomXMLPartsEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLPartsEvents COM interface which is disposible
		/// </summary>
		public static _CustomXMLPartsEventsProxy ToProxy(this Microsoft.Office.Core._CustomXMLPartsEvents resource)
		{
			return new _CustomXMLPartsEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomXMLPartsEvents_Event COM interface which is disposible
		/// </summary>
		public static _CustomXMLPartsEvents_EventProxy ToProxy(this Microsoft.Office.Core._CustomXMLPartsEvents_Event resource)
		{
			return new _CustomXMLPartsEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomXMLParts COM interface which is disposible
		/// </summary>
		public static CustomXMLPartsProxy ToProxy(this Microsoft.Office.Core.CustomXMLParts resource)
		{
			return new CustomXMLPartsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for GradientStop COM interface which is disposible
		/// </summary>
		public static GradientStopProxy ToProxy(this Microsoft.Office.Core.GradientStop resource)
		{
			return new GradientStopProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for GradientStops COM interface which is disposible
		/// </summary>
		public static GradientStopsProxy ToProxy(this Microsoft.Office.Core.GradientStops resource)
		{
			return new GradientStopsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SoftEdgeFormat COM interface which is disposible
		/// </summary>
		public static SoftEdgeFormatProxy ToProxy(this Microsoft.Office.Core.SoftEdgeFormat resource)
		{
			return new SoftEdgeFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for GlowFormat COM interface which is disposible
		/// </summary>
		public static GlowFormatProxy ToProxy(this Microsoft.Office.Core.GlowFormat resource)
		{
			return new GlowFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ReflectionFormat COM interface which is disposible
		/// </summary>
		public static ReflectionFormatProxy ToProxy(this Microsoft.Office.Core.ReflectionFormat resource)
		{
			return new ReflectionFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ParagraphFormat2 COM interface which is disposible
		/// </summary>
		public static ParagraphFormat2Proxy ToProxy(this Microsoft.Office.Core.ParagraphFormat2 resource)
		{
			return new ParagraphFormat2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Font2 COM interface which is disposible
		/// </summary>
		public static Font2Proxy ToProxy(this Microsoft.Office.Core.Font2 resource)
		{
			return new Font2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TextColumn2 COM interface which is disposible
		/// </summary>
		public static TextColumn2Proxy ToProxy(this Microsoft.Office.Core.TextColumn2 resource)
		{
			return new TextColumn2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TextRange2 COM interface which is disposible
		/// </summary>
		public static TextRange2Proxy ToProxy(this Microsoft.Office.Core.TextRange2 resource)
		{
			return new TextRange2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TextFrame2 COM interface which is disposible
		/// </summary>
		public static TextFrame2Proxy ToProxy(this Microsoft.Office.Core.TextFrame2 resource)
		{
			return new TextFrame2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThemeColor COM interface which is disposible
		/// </summary>
		public static ThemeColorProxy ToProxy(this Microsoft.Office.Core.ThemeColor resource)
		{
			return new ThemeColorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThemeColorScheme COM interface which is disposible
		/// </summary>
		public static ThemeColorSchemeProxy ToProxy(this Microsoft.Office.Core.ThemeColorScheme resource)
		{
			return new ThemeColorSchemeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThemeFont COM interface which is disposible
		/// </summary>
		public static ThemeFontProxy ToProxy(this Microsoft.Office.Core.ThemeFont resource)
		{
			return new ThemeFontProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThemeFonts COM interface which is disposible
		/// </summary>
		public static ThemeFontsProxy ToProxy(this Microsoft.Office.Core.ThemeFonts resource)
		{
			return new ThemeFontsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThemeFontScheme COM interface which is disposible
		/// </summary>
		public static ThemeFontSchemeProxy ToProxy(this Microsoft.Office.Core.ThemeFontScheme resource)
		{
			return new ThemeFontSchemeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ThemeEffectScheme COM interface which is disposible
		/// </summary>
		public static ThemeEffectSchemeProxy ToProxy(this Microsoft.Office.Core.ThemeEffectScheme resource)
		{
			return new ThemeEffectSchemeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for OfficeTheme COM interface which is disposible
		/// </summary>
		public static OfficeThemeProxy ToProxy(this Microsoft.Office.Core.OfficeTheme resource)
		{
			return new OfficeThemeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomTaskPane COM interface which is disposible
		/// </summary>
		public static _CustomTaskPaneProxy ToProxy(this Microsoft.Office.Core._CustomTaskPane resource)
		{
			return new _CustomTaskPaneProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomTaskPaneEvents COM interface which is disposible
		/// </summary>
		public static CustomTaskPaneEventsProxy ToProxy(this Microsoft.Office.Core.CustomTaskPaneEvents resource)
		{
			return new CustomTaskPaneEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomTaskPaneEvents COM interface which is disposible
		/// </summary>
		public static _CustomTaskPaneEventsProxy ToProxy(this Microsoft.Office.Core._CustomTaskPaneEvents resource)
		{
			return new _CustomTaskPaneEventsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for _CustomTaskPaneEvents_Event COM interface which is disposible
		/// </summary>
		public static _CustomTaskPaneEvents_EventProxy ToProxy(this Microsoft.Office.Core._CustomTaskPaneEvents_Event resource)
		{
			return new _CustomTaskPaneEvents_EventProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for CustomTaskPane COM interface which is disposible
		/// </summary>
		public static CustomTaskPaneProxy ToProxy(this Microsoft.Office.Core.CustomTaskPane resource)
		{
			return new CustomTaskPaneProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICTPFactory COM interface which is disposible
		/// </summary>
		public static ICTPFactoryProxy ToProxy(this Microsoft.Office.Core.ICTPFactory resource)
		{
			return new ICTPFactoryProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ICustomTaskPaneConsumer COM interface which is disposible
		/// </summary>
		public static ICustomTaskPaneConsumerProxy ToProxy(this Microsoft.Office.Core.ICustomTaskPaneConsumer resource)
		{
			return new ICustomTaskPaneConsumerProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IRibbonUI COM interface which is disposible
		/// </summary>
		public static IRibbonUIProxy ToProxy(this Microsoft.Office.Core.IRibbonUI resource)
		{
			return new IRibbonUIProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IRibbonControl COM interface which is disposible
		/// </summary>
		public static IRibbonControlProxy ToProxy(this Microsoft.Office.Core.IRibbonControl resource)
		{
			return new IRibbonControlProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IRibbonExtensibility COM interface which is disposible
		/// </summary>
		public static IRibbonExtensibilityProxy ToProxy(this Microsoft.Office.Core.IRibbonExtensibility resource)
		{
			return new IRibbonExtensibilityProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IAssistance COM interface which is disposible
		/// </summary>
		public static IAssistanceProxy ToProxy(this Microsoft.Office.Core.IAssistance resource)
		{
			return new IAssistanceProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoChartData COM interface which is disposible
		/// </summary>
		public static IMsoChartDataProxy ToProxy(this Microsoft.Office.Core.IMsoChartData resource)
		{
			return new IMsoChartDataProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoChart COM interface which is disposible
		/// </summary>
		public static IMsoChartProxy ToProxy(this Microsoft.Office.Core.IMsoChart resource)
		{
			return new IMsoChartProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoCorners COM interface which is disposible
		/// </summary>
		public static IMsoCornersProxy ToProxy(this Microsoft.Office.Core.IMsoCorners resource)
		{
			return new IMsoCornersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoLegend COM interface which is disposible
		/// </summary>
		public static IMsoLegendProxy ToProxy(this Microsoft.Office.Core.IMsoLegend resource)
		{
			return new IMsoLegendProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoBorder COM interface which is disposible
		/// </summary>
		public static IMsoBorderProxy ToProxy(this Microsoft.Office.Core.IMsoBorder resource)
		{
			return new IMsoBorderProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoWalls COM interface which is disposible
		/// </summary>
		public static IMsoWallsProxy ToProxy(this Microsoft.Office.Core.IMsoWalls resource)
		{
			return new IMsoWallsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoFloor COM interface which is disposible
		/// </summary>
		public static IMsoFloorProxy ToProxy(this Microsoft.Office.Core.IMsoFloor resource)
		{
			return new IMsoFloorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoPlotArea COM interface which is disposible
		/// </summary>
		public static IMsoPlotAreaProxy ToProxy(this Microsoft.Office.Core.IMsoPlotArea resource)
		{
			return new IMsoPlotAreaProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoChartArea COM interface which is disposible
		/// </summary>
		public static IMsoChartAreaProxy ToProxy(this Microsoft.Office.Core.IMsoChartArea resource)
		{
			return new IMsoChartAreaProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoSeriesLines COM interface which is disposible
		/// </summary>
		public static IMsoSeriesLinesProxy ToProxy(this Microsoft.Office.Core.IMsoSeriesLines resource)
		{
			return new IMsoSeriesLinesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoLeaderLines COM interface which is disposible
		/// </summary>
		public static IMsoLeaderLinesProxy ToProxy(this Microsoft.Office.Core.IMsoLeaderLines resource)
		{
			return new IMsoLeaderLinesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for GridLines COM interface which is disposible
		/// </summary>
		public static GridLinesProxy ToProxy(this Microsoft.Office.Core.GridLines resource)
		{
			return new GridLinesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoUpBars COM interface which is disposible
		/// </summary>
		public static IMsoUpBarsProxy ToProxy(this Microsoft.Office.Core.IMsoUpBars resource)
		{
			return new IMsoUpBarsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDownBars COM interface which is disposible
		/// </summary>
		public static IMsoDownBarsProxy ToProxy(this Microsoft.Office.Core.IMsoDownBars resource)
		{
			return new IMsoDownBarsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoInterior COM interface which is disposible
		/// </summary>
		public static IMsoInteriorProxy ToProxy(this Microsoft.Office.Core.IMsoInterior resource)
		{
			return new IMsoInteriorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ChartFillFormat COM interface which is disposible
		/// </summary>
		public static ChartFillFormatProxy ToProxy(this Microsoft.Office.Core.ChartFillFormat resource)
		{
			return new ChartFillFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for LegendEntries COM interface which is disposible
		/// </summary>
		public static LegendEntriesProxy ToProxy(this Microsoft.Office.Core.LegendEntries resource)
		{
			return new LegendEntriesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ChartFont COM interface which is disposible
		/// </summary>
		public static ChartFontProxy ToProxy(this Microsoft.Office.Core.ChartFont resource)
		{
			return new ChartFontProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ChartColorFormat COM interface which is disposible
		/// </summary>
		public static ChartColorFormatProxy ToProxy(this Microsoft.Office.Core.ChartColorFormat resource)
		{
			return new ChartColorFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for LegendEntry COM interface which is disposible
		/// </summary>
		public static LegendEntryProxy ToProxy(this Microsoft.Office.Core.LegendEntry resource)
		{
			return new LegendEntryProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoLegendKey COM interface which is disposible
		/// </summary>
		public static IMsoLegendKeyProxy ToProxy(this Microsoft.Office.Core.IMsoLegendKey resource)
		{
			return new IMsoLegendKeyProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SeriesCollection COM interface which is disposible
		/// </summary>
		public static SeriesCollectionProxy ToProxy(this Microsoft.Office.Core.SeriesCollection resource)
		{
			return new SeriesCollectionProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoSeries COM interface which is disposible
		/// </summary>
		public static IMsoSeriesProxy ToProxy(this Microsoft.Office.Core.IMsoSeries resource)
		{
			return new IMsoSeriesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoErrorBars COM interface which is disposible
		/// </summary>
		public static IMsoErrorBarsProxy ToProxy(this Microsoft.Office.Core.IMsoErrorBars resource)
		{
			return new IMsoErrorBarsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoTrendline COM interface which is disposible
		/// </summary>
		public static IMsoTrendlineProxy ToProxy(this Microsoft.Office.Core.IMsoTrendline resource)
		{
			return new IMsoTrendlineProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Trendlines COM interface which is disposible
		/// </summary>
		public static TrendlinesProxy ToProxy(this Microsoft.Office.Core.Trendlines resource)
		{
			return new TrendlinesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDataLabels COM interface which is disposible
		/// </summary>
		public static IMsoDataLabelsProxy ToProxy(this Microsoft.Office.Core.IMsoDataLabels resource)
		{
			return new IMsoDataLabelsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDataLabel COM interface which is disposible
		/// </summary>
		public static IMsoDataLabelProxy ToProxy(this Microsoft.Office.Core.IMsoDataLabel resource)
		{
			return new IMsoDataLabelProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Points COM interface which is disposible
		/// </summary>
		public static PointsProxy ToProxy(this Microsoft.Office.Core.Points resource)
		{
			return new PointsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ChartPoint COM interface which is disposible
		/// </summary>
		public static ChartPointProxy ToProxy(this Microsoft.Office.Core.ChartPoint resource)
		{
			return new ChartPointProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Axes COM interface which is disposible
		/// </summary>
		public static AxesProxy ToProxy(this Microsoft.Office.Core.Axes resource)
		{
			return new AxesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoAxis COM interface which is disposible
		/// </summary>
		public static IMsoAxisProxy ToProxy(this Microsoft.Office.Core.IMsoAxis resource)
		{
			return new IMsoAxisProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDataTable COM interface which is disposible
		/// </summary>
		public static IMsoDataTableProxy ToProxy(this Microsoft.Office.Core.IMsoDataTable resource)
		{
			return new IMsoDataTableProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoChartTitle COM interface which is disposible
		/// </summary>
		public static IMsoChartTitleProxy ToProxy(this Microsoft.Office.Core.IMsoChartTitle resource)
		{
			return new IMsoChartTitleProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoAxisTitle COM interface which is disposible
		/// </summary>
		public static IMsoAxisTitleProxy ToProxy(this Microsoft.Office.Core.IMsoAxisTitle resource)
		{
			return new IMsoAxisTitleProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDisplayUnitLabel COM interface which is disposible
		/// </summary>
		public static IMsoDisplayUnitLabelProxy ToProxy(this Microsoft.Office.Core.IMsoDisplayUnitLabel resource)
		{
			return new IMsoDisplayUnitLabelProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoTickLabels COM interface which is disposible
		/// </summary>
		public static IMsoTickLabelsProxy ToProxy(this Microsoft.Office.Core.IMsoTickLabels resource)
		{
			return new IMsoTickLabelsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoHyperlinks COM interface which is disposible
		/// </summary>
		public static IMsoHyperlinksProxy ToProxy(this Microsoft.Office.Core.IMsoHyperlinks resource)
		{
			return new IMsoHyperlinksProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoDropLines COM interface which is disposible
		/// </summary>
		public static IMsoDropLinesProxy ToProxy(this Microsoft.Office.Core.IMsoDropLines resource)
		{
			return new IMsoDropLinesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoHiLoLines COM interface which is disposible
		/// </summary>
		public static IMsoHiLoLinesProxy ToProxy(this Microsoft.Office.Core.IMsoHiLoLines resource)
		{
			return new IMsoHiLoLinesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoChartGroup COM interface which is disposible
		/// </summary>
		public static IMsoChartGroupProxy ToProxy(this Microsoft.Office.Core.IMsoChartGroup resource)
		{
			return new IMsoChartGroupProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ChartGroups COM interface which is disposible
		/// </summary>
		public static ChartGroupsProxy ToProxy(this Microsoft.Office.Core.ChartGroups resource)
		{
			return new ChartGroupsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoCharacters COM interface which is disposible
		/// </summary>
		public static IMsoCharactersProxy ToProxy(this Microsoft.Office.Core.IMsoCharacters resource)
		{
			return new IMsoCharactersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoChartFormat COM interface which is disposible
		/// </summary>
		public static IMsoChartFormatProxy ToProxy(this Microsoft.Office.Core.IMsoChartFormat resource)
		{
			return new IMsoChartFormatProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for BulletFormat2 COM interface which is disposible
		/// </summary>
		public static BulletFormat2Proxy ToProxy(this Microsoft.Office.Core.BulletFormat2 resource)
		{
			return new BulletFormat2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TabStops2 COM interface which is disposible
		/// </summary>
		public static TabStops2Proxy ToProxy(this Microsoft.Office.Core.TabStops2 resource)
		{
			return new TabStops2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for TabStop2 COM interface which is disposible
		/// </summary>
		public static TabStop2Proxy ToProxy(this Microsoft.Office.Core.TabStop2 resource)
		{
			return new TabStop2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Ruler2 COM interface which is disposible
		/// </summary>
		public static Ruler2Proxy ToProxy(this Microsoft.Office.Core.Ruler2 resource)
		{
			return new Ruler2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for RulerLevels2 COM interface which is disposible
		/// </summary>
		public static RulerLevels2Proxy ToProxy(this Microsoft.Office.Core.RulerLevels2 resource)
		{
			return new RulerLevels2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for RulerLevel2 COM interface which is disposible
		/// </summary>
		public static RulerLevel2Proxy ToProxy(this Microsoft.Office.Core.RulerLevel2 resource)
		{
			return new RulerLevel2Proxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for EncryptionProvider COM interface which is disposible
		/// </summary>
		public static EncryptionProviderProxy ToProxy(this Microsoft.Office.Core.EncryptionProvider resource)
		{
			return new EncryptionProviderProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IBlogExtensibility COM interface which is disposible
		/// </summary>
		public static IBlogExtensibilityProxy ToProxy(this Microsoft.Office.Core.IBlogExtensibility resource)
		{
			return new IBlogExtensibilityProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IBlogPictureExtensibility COM interface which is disposible
		/// </summary>
		public static IBlogPictureExtensibilityProxy ToProxy(this Microsoft.Office.Core.IBlogPictureExtensibility resource)
		{
			return new IBlogPictureExtensibilityProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IConverterPreferences COM interface which is disposible
		/// </summary>
		public static IConverterPreferencesProxy ToProxy(this Microsoft.Office.Core.IConverterPreferences resource)
		{
			return new IConverterPreferencesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IConverterApplicationPreferences COM interface which is disposible
		/// </summary>
		public static IConverterApplicationPreferencesProxy ToProxy(this Microsoft.Office.Core.IConverterApplicationPreferences resource)
		{
			return new IConverterApplicationPreferencesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IConverterUICallback COM interface which is disposible
		/// </summary>
		public static IConverterUICallbackProxy ToProxy(this Microsoft.Office.Core.IConverterUICallback resource)
		{
			return new IConverterUICallbackProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IConverter COM interface which is disposible
		/// </summary>
		public static IConverterProxy ToProxy(this Microsoft.Office.Core.IConverter resource)
		{
			return new IConverterProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArt COM interface which is disposible
		/// </summary>
		public static SmartArtProxy ToProxy(this Microsoft.Office.Core.SmartArt resource)
		{
			return new SmartArtProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtNodes COM interface which is disposible
		/// </summary>
		public static SmartArtNodesProxy ToProxy(this Microsoft.Office.Core.SmartArtNodes resource)
		{
			return new SmartArtNodesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtNode COM interface which is disposible
		/// </summary>
		public static SmartArtNodeProxy ToProxy(this Microsoft.Office.Core.SmartArtNode resource)
		{
			return new SmartArtNodeProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtLayouts COM interface which is disposible
		/// </summary>
		public static SmartArtLayoutsProxy ToProxy(this Microsoft.Office.Core.SmartArtLayouts resource)
		{
			return new SmartArtLayoutsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtLayout COM interface which is disposible
		/// </summary>
		public static SmartArtLayoutProxy ToProxy(this Microsoft.Office.Core.SmartArtLayout resource)
		{
			return new SmartArtLayoutProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtQuickStyles COM interface which is disposible
		/// </summary>
		public static SmartArtQuickStylesProxy ToProxy(this Microsoft.Office.Core.SmartArtQuickStyles resource)
		{
			return new SmartArtQuickStylesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtQuickStyle COM interface which is disposible
		/// </summary>
		public static SmartArtQuickStyleProxy ToProxy(this Microsoft.Office.Core.SmartArtQuickStyle resource)
		{
			return new SmartArtQuickStyleProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtColors COM interface which is disposible
		/// </summary>
		public static SmartArtColorsProxy ToProxy(this Microsoft.Office.Core.SmartArtColors resource)
		{
			return new SmartArtColorsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for SmartArtColor COM interface which is disposible
		/// </summary>
		public static SmartArtColorProxy ToProxy(this Microsoft.Office.Core.SmartArtColor resource)
		{
			return new SmartArtColorProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerField COM interface which is disposible
		/// </summary>
		public static PickerFieldProxy ToProxy(this Microsoft.Office.Core.PickerField resource)
		{
			return new PickerFieldProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerFields COM interface which is disposible
		/// </summary>
		public static PickerFieldsProxy ToProxy(this Microsoft.Office.Core.PickerFields resource)
		{
			return new PickerFieldsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerProperty COM interface which is disposible
		/// </summary>
		public static PickerPropertyProxy ToProxy(this Microsoft.Office.Core.PickerProperty resource)
		{
			return new PickerPropertyProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerProperties COM interface which is disposible
		/// </summary>
		public static PickerPropertiesProxy ToProxy(this Microsoft.Office.Core.PickerProperties resource)
		{
			return new PickerPropertiesProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerResult COM interface which is disposible
		/// </summary>
		public static PickerResultProxy ToProxy(this Microsoft.Office.Core.PickerResult resource)
		{
			return new PickerResultProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerResults COM interface which is disposible
		/// </summary>
		public static PickerResultsProxy ToProxy(this Microsoft.Office.Core.PickerResults resource)
		{
			return new PickerResultsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PickerDialog COM interface which is disposible
		/// </summary>
		public static PickerDialogProxy ToProxy(this Microsoft.Office.Core.PickerDialog resource)
		{
			return new PickerDialogProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for IMsoContactCard COM interface which is disposible
		/// </summary>
		public static IMsoContactCardProxy ToProxy(this Microsoft.Office.Core.IMsoContactCard resource)
		{
			return new IMsoContactCardProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for EffectParameter COM interface which is disposible
		/// </summary>
		public static EffectParameterProxy ToProxy(this Microsoft.Office.Core.EffectParameter resource)
		{
			return new EffectParameterProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for EffectParameters COM interface which is disposible
		/// </summary>
		public static EffectParametersProxy ToProxy(this Microsoft.Office.Core.EffectParameters resource)
		{
			return new EffectParametersProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PictureEffect COM interface which is disposible
		/// </summary>
		public static PictureEffectProxy ToProxy(this Microsoft.Office.Core.PictureEffect resource)
		{
			return new PictureEffectProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for PictureEffects COM interface which is disposible
		/// </summary>
		public static PictureEffectsProxy ToProxy(this Microsoft.Office.Core.PictureEffects resource)
		{
			return new PictureEffectsProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for Crop COM interface which is disposible
		/// </summary>
		public static CropProxy ToProxy(this Microsoft.Office.Core.Crop resource)
		{
			return new CropProxy(resource);
		}

		/// <summary>
		/// Wrapper Proxy for ContactCard COM interface which is disposible
		/// </summary>
		public static ContactCardProxy ToProxy(this Microsoft.Office.Core.ContactCard resource)
		{
			return new ContactCardProxy(resource);
		}

	}
}
