using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = Microsoft.Office.Interop.Excel.Window;
using Windows = Microsoft.Office.Interop.Excel.Windows;

namespace Excel.TestDoubles
{
#pragma warning disable 0067
    public class WorkbookTestDouble : Workbook
    {
        public WorkbookTestDouble(ApplicationTestDouble application, WindowTestDouble window)
        {
            Application = application;
            Windows = new WindowsTestDouble(application)
            {
                window
            };
        }

        void _Workbook.Activate()
        {
            throw new NotImplementedException();
        }

        public void ChangeFileAccess(XlFileAccess mode, object writePassword, object notify)
        {
            throw new NotImplementedException();
        }

        public void ChangeLink(string name, string newName, XlLinkType type = XlLinkType.xlLinkTypeExcelLinks)
        {
            throw new NotImplementedException();
        }

        public void Close(object saveChanges, object filename, object routeWorkbook)
        {
            var applicationTestDouble = ((ApplicationTestDouble)Application);
            ((WorkbooksTestDouble)applicationTestDouble.Workbooks).Remove(this);
            applicationTestDouble.RaiseWorkbookBeforeClose(this);
            applicationTestDouble.RaiseWorkbookDeactivate(this);
        }

        public void DeleteNumberFormat(string numberFormat)
        {
            throw new NotImplementedException();
        }

        public bool ExclusiveAccess()
        {
            throw new NotImplementedException();
        }

        public void ForwardMailer()
        {
            throw new NotImplementedException();
        }

        public object LinkInfo(string name, XlLinkInfo linkInfo, object type, object editionRef)
        {
            throw new NotImplementedException();
        }

        public object LinkSources(object type)
        {
            throw new NotImplementedException();
        }

        public void MergeWorkbook(object filename)
        {
            throw new NotImplementedException();
        }

        public Window NewWindow()
        {
            throw new NotImplementedException();
        }

        public void OpenLinks(string name, object readOnly, object type)
        {
            throw new NotImplementedException();
        }

        public PivotCaches PivotCaches()
        {
            throw new NotImplementedException();
        }

        public void Post(object destName)
        {
            throw new NotImplementedException();
        }

        public void _PrintOut(object @from, object to, object copies, object preview, object activePrinter, object printToFile,
            object collate)
        {
            throw new NotImplementedException();
        }

        public void PrintPreview(object enableChanges)
        {
            throw new NotImplementedException();
        }

        public void _Protect(object password, object structure, object windows)
        {
            throw new NotImplementedException();
        }

        public void ProtectSharing(object filename, object password, object writeResPassword, object readOnlyRecommended,
            object createBackup, object sharingPassword)
        {
            throw new NotImplementedException();
        }

        public void RefreshAll()
        {
            throw new NotImplementedException();
        }

        public void Reply()
        {
            throw new NotImplementedException();
        }

        public void ReplyAll()
        {
            throw new NotImplementedException();
        }

        public void RemoveUser(int index)
        {
            throw new NotImplementedException();
        }

        public void Route()
        {
            throw new NotImplementedException();
        }

        public void RunAutoMacros(XlRunAutoMacro which)
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            throw new NotImplementedException();
        }

        public void _SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended,
            object createBackup, XlSaveAsAccessMode accessMode, object conflictResolution, object addToMru,
            object textCodepage, object textVisualLayout)
        {
            throw new NotImplementedException();
        }

        public void SaveCopyAs(object filename)
        {
            throw new NotImplementedException();
        }

        public void SendMail(object recipients, object subject, object returnReceipt)
        {
            throw new NotImplementedException();
        }

        public void SendMailer(object fileFormat, XlPriority priority = XlPriority.xlPriorityNormal)
        {
            throw new NotImplementedException();
        }

        public void SetLinkOnData(string name, object procedure)
        {
            throw new NotImplementedException();
        }

        public void Unprotect(object password)
        {
            throw new NotImplementedException();
        }

        public void UnprotectSharing(object sharingPassword)
        {
            throw new NotImplementedException();
        }

        public void UpdateFromFile()
        {
            throw new NotImplementedException();
        }

        public void UpdateLink(object name, object type)
        {
            throw new NotImplementedException();
        }

        public void HighlightChangesOptions(object when, object who, object where)
        {
            throw new NotImplementedException();
        }

        public void PurgeChangeHistoryNow(int days, object sharingPassword)
        {
            throw new NotImplementedException();
        }

        public void AcceptAllChanges(object when, object who, object where)
        {
            throw new NotImplementedException();
        }

        public void RejectAllChanges(object when, object who, object where)
        {
            throw new NotImplementedException();
        }

        public void PivotTableWizard(object sourceType, object sourceData, object tableDestination, object tableName, object rowGrand,
            object columnGrand, object saveData, object hasAutoFormat, object autoPage, object reserved, object backgroundQuery,
            object optimizeCache, object pageFieldOrder, object pageFieldWrapCount, object readData, object connection)
        {
            throw new NotImplementedException();
        }

        public void ResetColors()
        {
            throw new NotImplementedException();
        }

        public void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo,
            object method, object headerInfo)
        {
            throw new NotImplementedException();
        }

        public void AddToFavorites()
        {
            throw new NotImplementedException();
        }

        public void PrintOut(object @from, object to, object copies, object preview, object activePrinter, object printToFile,
            object collate, object prToFileName)
        {
            throw new NotImplementedException();
        }

        public void WebPagePreview()
        {
            throw new NotImplementedException();
        }

        public void ReloadAs(MsoEncoding encoding)
        {
            throw new NotImplementedException();
        }

        public void Dummy17(int calcid)
        {
            throw new NotImplementedException();
        }

        public void sblt(string s)
        {
            throw new NotImplementedException();
        }

        public void BreakLink(string name, XlLinkType type)
        {
            throw new NotImplementedException();
        }

        public void Dummy16()
        {
            throw new NotImplementedException();
        }

        public void SaveAs(object filename, object fileFormat, object password, object writeResPassword, object readOnlyRecommended,
            object createBackup, XlSaveAsAccessMode accessMode, object conflictResolution, object addToMru,
            object textCodepage, object textVisualLayout, object local)
        {
            throw new NotImplementedException();
        }

        public void CheckIn(object saveChanges, object comments, object makePublic)
        {
            throw new NotImplementedException();
        }

        public bool CanCheckIn()
        {
            throw new NotImplementedException();
        }

        public void SendForReview(object recipients, object subject, object showMessage, object includeAttachment)
        {
            throw new NotImplementedException();
        }

        public void ReplyWithChanges(object showMessage)
        {
            throw new NotImplementedException();
        }

        public void EndReview()
        {
            throw new NotImplementedException();
        }

        public void SetPasswordEncryptionOptions(object passwordEncryptionProvider, object passwordEncryptionAlgorithm,
            object passwordEncryptionKeyLength, object passwordEncryptionFileProperties)
        {
            throw new NotImplementedException();
        }

        public void Protect(object password, object structure, object windows)
        {
            throw new NotImplementedException();
        }

        public void RecheckSmartTags()
        {
            throw new NotImplementedException();
        }

        public void SendFaxOverInternet(object recipients, object subject, object showMessage)
        {
            throw new NotImplementedException();
        }

        public XlXmlImportResult XmlImport(string url, out XmlMap importMap, object overwrite, object destination)
        {
            throw new NotImplementedException();
        }

        public XlXmlImportResult XmlImportXml(string data, out XmlMap importMap, object overwrite, object destination)
        {
            throw new NotImplementedException();
        }

        public void SaveAsXMLData(string filename, XmlMap map)
        {
            throw new NotImplementedException();
        }

        public void ToggleFormsDesign()
        {
            throw new NotImplementedException();
        }

        public void RemoveDocumentInformation(XlRemoveDocInfoType removeDocInfoType)
        {
            throw new NotImplementedException();
        }

        public void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType)
        {
            throw new NotImplementedException();
        }

        public void LockServerFile()
        {
            throw new NotImplementedException();
        }

        public WorkflowTasks GetWorkflowTasks()
        {
            throw new NotImplementedException();
        }

        public WorkflowTemplates GetWorkflowTemplates()
        {
            throw new NotImplementedException();
        }

        public void PrintOutEx(object @from, object to, object copies, object preview, object activePrinter, object printToFile,
            object collate, object prToFileName, object ignorePrintAreas)
        {
            throw new NotImplementedException();
        }

        public void ApplyTheme(string filename)
        {
            throw new NotImplementedException();
        }

        public void EnableConnections()
        {
            throw new NotImplementedException();
        }

        public void ExportAsFixedFormat(XlFixedFormatType type, object filename, object quality, object includeDocProperties,
            object ignorePrintAreas, object from, object to, object openAfterPublish, object fixedFormatExtClassPtr)
        {
            throw new NotImplementedException();
        }

        public void ProtectSharingEx(object filename, object password, object writeResPassword, object readOnlyRecommended,
            object createBackup, object sharingPassword, object fileFormat)
        {
            throw new NotImplementedException();
        }

        public void Dummy26()
        {
            throw new NotImplementedException();
        }

        public void Dummy27()
        {
            throw new NotImplementedException();
        }

        public Application Application { get; private set; }
        public XlCreator Creator { get; private set; }
        public object Parent { get; private set; }
        public bool AcceptLabelsInFormulas { get; set; }
        public Chart ActiveChart { get; private set; }
        public object ActiveSheet { get; private set; }
        public string Author { get; set; }
        public int AutoUpdateFrequency { get; set; }
        public bool AutoUpdateSaveChanges { get; set; }
        public int ChangeHistoryDuration { get; set; }
        public object BuiltinDocumentProperties { get; private set; }
        public Sheets Charts { get; private set; }
        public string CodeName { get; private set; }
        public string _CodeName { get; set; }
        public object get_Colors(object index)
        {
            throw new NotImplementedException();
        }

        public void set_Colors(object index, object rhs)
        {
            throw new NotImplementedException();
        }

        public CommandBars CommandBars { get; private set; }
        public string Comments { get; set; }
        public XlSaveConflictResolution ConflictResolution { get; set; }
        public object Container { get; private set; }
        public bool CreateBackup { get; private set; }
        public object CustomDocumentProperties { get; private set; }
        public bool Date1904 { get; set; }
        public Sheets DialogSheets { get; private set; }
        public XlDisplayDrawingObjects DisplayDrawingObjects { get; set; }
        public XlFileFormat FileFormat { get; private set; }
        public string FullName { get; private set; }
        public bool HasMailer { get; set; }
        public bool HasPassword { get; private set; }
        public bool HasRoutingSlip { get; set; }
        public bool IsAddin { get; set; }
        public string Keywords { get; set; }
        public Mailer Mailer { get; private set; }
        public Sheets Modules { get; private set; }
        public bool MultiUserEditing { get; private set; }
        public string Name { get; private set; }
        public Names Names { get; private set; }
        public string OnSave { get; set; }
        public string OnSheetActivate { get; set; }
        public string OnSheetDeactivate { get; set; }
        public string Path { get; private set; }
        public bool PersonalViewListSettings { get; set; }
        public bool PersonalViewPrintSettings { get; set; }
        public bool PrecisionAsDisplayed { get; set; }
        public bool ProtectStructure { get; private set; }
        public bool ProtectWindows { get; private set; }
        public bool ReadOnly { get; private set; }
        public bool _ReadOnlyRecommended { get; private set; }
        public int RevisionNumber { get; private set; }
        public bool Routed { get; private set; }
        public RoutingSlip RoutingSlip { get; private set; }
        public bool Saved { get; set; }
        public bool SaveLinkValues { get; set; }
        public Sheets Sheets { get; private set; }
        public bool ShowConflictHistory { get; set; }
        public Styles Styles { get; private set; }
        public string Subject { get; set; }
        public string Title { get; set; }
        public bool UpdateRemoteReferences { get; set; }
        public bool UserControl { get; set; }
        public object UserStatus { get; private set; }
        public CustomViews CustomViews { get; private set; }
        public Windows Windows { get; private set; }
        public Sheets Worksheets { get; private set; }
        public bool WriteReserved { get; private set; }
        public string WriteReservedBy { get; private set; }
        public Sheets Excel4IntlMacroSheets { get; private set; }
        public Sheets Excel4MacroSheets { get; private set; }
        public bool TemplateRemoveExtData { get; set; }
        public bool HighlightChangesOnScreen { get; set; }
        public bool KeepChangeHistory { get; set; }
        public bool ListChangesOnNewSheet { get; set; }
        public VBProject VBProject { get; private set; }
        public bool IsInplace { get; private set; }
        public PublishObjects PublishObjects { get; private set; }
        public WebOptions WebOptions { get; private set; }
        public HTMLProject HTMLProject { get; private set; }
        public bool EnvelopeVisible { get; set; }
        public int CalculationVersion { get; private set; }
        public bool VBASigned { get; private set; }
        public bool ShowPivotTableFieldList { get; set; }
        public XlUpdateLinks UpdateLinks { get; set; }
        public bool EnableAutoRecover { get; set; }
        public bool RemovePersonalInformation { get; set; }
        public string FullNameURLEncoded { get; private set; }
        public string Password { get; set; }
        public string WritePassword { get; set; }
        public string PasswordEncryptionProvider { get; private set; }
        public string PasswordEncryptionAlgorithm { get; private set; }
        public int PasswordEncryptionKeyLength { get; private set; }
        public bool PasswordEncryptionFileProperties { get; private set; }
        public bool ReadOnlyRecommended { get; set; }
        public SmartTagOptions SmartTagOptions { get; private set; }
        public Permission Permission { get; private set; }
        public SharedWorkspace SharedWorkspace { get; private set; }
        Sync _Workbook.Sync { get { throw new NotImplementedException();} }
        public XmlNamespaces XmlNamespaces { get; private set; }
        public XmlMaps XmlMaps { get; private set; }
        public SmartDocument SmartDocument { get; private set; }
        public DocumentLibraryVersions DocumentLibraryVersions { get; private set; }
        public bool InactiveListBorderVisible { get; set; }
        public bool DisplayInkComments { get; set; }
        public MetaProperties ContentTypeProperties { get; private set; }
        public Connections Connections { get; private set; }
        public SignatureSet Signatures { get; private set; }
        public ServerPolicy ServerPolicy { get; private set; }
        public DocumentInspectors DocumentInspectors { get; private set; }
        public ServerViewableItems ServerViewableItems { get; private set; }
        public TableStyles TableStyles { get; private set; }
        public object DefaultTableStyle { get; set; }
        public object DefaultPivotTableStyle { get; set; }
        public bool CheckCompatibility { get; set; }
        public bool HasVBProject { get; private set; }
        public CustomXMLParts CustomXMLParts { get; private set; }
        public bool Final { get; set; }
        public Research Research { get; private set; }
        public OfficeTheme Theme { get; private set; }
        public bool Excel8CompatibilityMode { get; private set; }
        public bool ConnectionsDisabled { get; private set; }
        public bool ShowPivotChartActiveFields { get; set; }
        public IconSets IconSets { get; private set; }
        public string EncryptionProvider { get; set; }
        public bool DoNotPromptForConvert { get; set; }
        public bool ForceFullCalculation { get; set; }
        public SlicerCaches SlicerCaches { get; private set; }
        public Slicer ActiveSlicer { get; private set; }
        public object DefaultSlicerStyle { get; set; }
        public int AccuracyVersion { get; set; }
        public bool CaseSensitive { get; private set; }
        public bool UseWholeCellCriteria { get; private set; }
        public bool UseWildcards { get; private set; }
        public object PivotTables { get; private set; }
        public Model Model { get; private set; }
        public bool ChartDataPointTrack { get; set; }
        public object DefaultTimelineStyle { get; set; }
        public event WorkbookEvents_OpenEventHandler Open;
        public event WorkbookEvents_ActivateEventHandler Activate;
        public event WorkbookEvents_DeactivateEventHandler Deactivate;
        public event WorkbookEvents_BeforeCloseEventHandler BeforeClose;
        public event WorkbookEvents_BeforeSaveEventHandler BeforeSave;
        public event WorkbookEvents_BeforePrintEventHandler BeforePrint;
        public event WorkbookEvents_NewSheetEventHandler NewSheet;
        public event WorkbookEvents_AddinInstallEventHandler AddinInstall;
        public event WorkbookEvents_AddinUninstallEventHandler AddinUninstall;
        public event WorkbookEvents_WindowResizeEventHandler WindowResize;
        public event WorkbookEvents_WindowActivateEventHandler WindowActivate;
        public event WorkbookEvents_WindowDeactivateEventHandler WindowDeactivate;
        public event WorkbookEvents_SheetSelectionChangeEventHandler SheetSelectionChange;
        public event WorkbookEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        public event WorkbookEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        public event WorkbookEvents_SheetActivateEventHandler SheetActivate;
        public event WorkbookEvents_SheetDeactivateEventHandler SheetDeactivate;
        public event WorkbookEvents_SheetCalculateEventHandler SheetCalculate;
        public event WorkbookEvents_SheetChangeEventHandler SheetChange;
        public event WorkbookEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlink;
        public event WorkbookEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        public event WorkbookEvents_PivotTableCloseConnectionEventHandler PivotTableCloseConnection;
        public event WorkbookEvents_PivotTableOpenConnectionEventHandler PivotTableOpenConnection;
        event WorkbookEvents_SyncEventHandler WorkbookEvents_Event.Sync
        {
            add { throw new NotImplementedException(); }
            remove { throw new NotImplementedException(); }
        }

        public event WorkbookEvents_BeforeXmlImportEventHandler BeforeXmlImport;
        public event WorkbookEvents_AfterXmlImportEventHandler AfterXmlImport;
        public event WorkbookEvents_BeforeXmlExportEventHandler BeforeXmlExport;
        public event WorkbookEvents_AfterXmlExportEventHandler AfterXmlExport;
        public event WorkbookEvents_RowsetCompleteEventHandler RowsetComplete;
        public event WorkbookEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChange;
        public event WorkbookEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChanges;
        public event WorkbookEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChanges;
        public event WorkbookEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChanges;
        public event WorkbookEvents_SheetPivotTableChangeSyncEventHandler SheetPivotTableChangeSync;
        public event WorkbookEvents_AfterSaveEventHandler AfterSave;
        public event WorkbookEvents_NewChartEventHandler NewChart;
        public event WorkbookEvents_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderComplete;
        public event WorkbookEvents_SheetTableUpdateEventHandler SheetTableUpdate;
        public event WorkbookEvents_ModelChangeEventHandler ModelChange;
        public event WorkbookEvents_SheetBeforeDeleteEventHandler SheetBeforeDelete;
    }
}