using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Excel.Application;
using Window = Microsoft.Office.Interop.Excel.Window;
using Windows = Microsoft.Office.Interop.Excel.Windows;

namespace Excel.TestDoubles
{
    public class ApplicationTestDouble : Application
    {
        public ApplicationTestDouble(Version version)
        {
            Workbooks = new WorkbooksTestDouble(this);
            Windows = new WindowsTestDouble(this)
            {
                new WindowTestDouble()
            };
            Version = version.ToString();
        }

        public void Calculate()
        {
            throw new NotImplementedException();
        }

        public void DDEExecute(int channel, string @string)
        {
            throw new NotImplementedException();
        }

        public int DDEInitiate(string app, string topic)
        {
            throw new NotImplementedException();
        }

        public void DDEPoke(int channel, object item, object data)
        {
            throw new NotImplementedException();
        }

        public object DDERequest(int channel, string Item)
        {
            throw new NotImplementedException();
        }

        public void DDETerminate(int channel)
        {
            throw new NotImplementedException();
        }

        public object Evaluate(object Name)
        {
            throw new NotImplementedException();
        }

        public object _Evaluate(object Name)
        {
            throw new NotImplementedException();
        }

        public object ExecuteExcel4Macro(string String)
        {
            throw new NotImplementedException();
        }

        public Range Intersect(Range Arg1, Range Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7,
            object Arg8,
            object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16,
            object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23,
            object Arg24,
            object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public object Run(object Macro, object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6,
            object Arg7,
            object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22,
            object Arg23,
            object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public object _Run2(object Macro, object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6,
            object Arg7,
            object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22,
            object Arg23,
            object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public void SendKeys(object Keys, object Wait)
        {
            throw new NotImplementedException();
        }

        public Range Union(Range Arg1, Range Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7,
            object Arg8,
            object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16,
            object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23,
            object Arg24,
            object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public void ActivateMicrosoftApp(XlMSApplication Index)
        {
            throw new NotImplementedException();
        }

        public void AddChartAutoFormat(object Chart, string Name, object Description)
        {
            throw new NotImplementedException();
        }

        public void AddCustomList(object ListArray, object ByRow)
        {
            throw new NotImplementedException();
        }

        public double CentimetersToPoints(double Centimeters)
        {
            throw new NotImplementedException();
        }

        public bool CheckSpelling(string Word, object CustomDictionary, object IgnoreUppercase)
        {
            throw new NotImplementedException();
        }

        public object ConvertFormula(object Formula, XlReferenceStyle FromReferenceStyle, object ToReferenceStyle,
            object ToAbsolute,
            object RelativeTo)
        {
            throw new NotImplementedException();
        }

        public object Dummy1(object Arg1, object Arg2, object Arg3, object Arg4)
        {
            throw new NotImplementedException();
        }

        public object Dummy2(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7,
            object Arg8)
        {
            throw new NotImplementedException();
        }

        public object Dummy3()
        {
            throw new NotImplementedException();
        }

        public object Dummy4(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7,
            object Arg8,
            object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15)
        {
            throw new NotImplementedException();
        }

        public object Dummy5(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7,
            object Arg8,
            object Arg9, object Arg10, object Arg11, object Arg12, object Arg13)
        {
            throw new NotImplementedException();
        }

        public object Dummy6()
        {
            throw new NotImplementedException();
        }

        public object Dummy7()
        {
            throw new NotImplementedException();
        }

        public object Dummy8(object Arg1)
        {
            throw new NotImplementedException();
        }

        public object Dummy9()
        {
            throw new NotImplementedException();
        }

        public bool Dummy10(object arg)
        {
            throw new NotImplementedException();
        }

        public void Dummy11()
        {
            throw new NotImplementedException();
        }

        public void DeleteChartAutoFormat(string Name)
        {
            throw new NotImplementedException();
        }

        public void DeleteCustomList(int ListNum)
        {
            throw new NotImplementedException();
        }

        public void DoubleClick()
        {
            throw new NotImplementedException();
        }

        public void _FindFile()
        {
            throw new NotImplementedException();
        }

        public object GetCustomListContents(int ListNum)
        {
            throw new NotImplementedException();
        }

        public int GetCustomListNum(object ListArray)
        {
            throw new NotImplementedException();
        }

        public object GetOpenFilename(object FileFilter, object FilterIndex, object Title, object ButtonText,
            object MultiSelect)
        {
            throw new NotImplementedException();
        }

        public object GetSaveAsFilename(object InitialFilename, object FileFilter, object FilterIndex, object Title,
            object ButtonText)
        {
            throw new NotImplementedException();
        }

        public void Goto(object Reference, object Scroll)
        {
            throw new NotImplementedException();
        }

        public void Help(object HelpFile, object HelpContextID)
        {
            throw new NotImplementedException();
        }

        public double InchesToPoints(double Inches)
        {
            throw new NotImplementedException();
        }

        public object InputBox(string Prompt, object Title, object Default, object Left, object Top, object HelpFile,
            object HelpContextID, object Type)
        {
            throw new NotImplementedException();
        }

        public void MacroOptions(object Macro, object Description, object HasMenu, object MenuText,
            object HasShortcutKey,
            object ShortcutKey, object Category, object StatusBar, object HelpContextID, object HelpFile)
        {
            throw new NotImplementedException();
        }

        public void MailLogoff()
        {
            throw new NotImplementedException();
        }

        public void MailLogon(object Name, object Password, object DownloadNewMail)
        {
            throw new NotImplementedException();
        }

        public Workbook NextLetter()
        {
            throw new NotImplementedException();
        }

        public void OnKey(string Key, object Procedure)
        {
            throw new NotImplementedException();
        }

        public void OnRepeat(string Text, string Procedure)
        {
            throw new NotImplementedException();
        }

        public void OnTime(object EarliestTime, string Procedure, object LatestTime, object Schedule)
        {
            throw new NotImplementedException();
        }

        public void OnUndo(string Text, string Procedure)
        {
            throw new NotImplementedException();
        }

        public void Quit()
        {
            throw new NotImplementedException();
        }

        public void RecordMacro(object BasicCode, object XlmCode)
        {
            throw new NotImplementedException();
        }

        public bool RegisterXLL(string Filename)
        {
            throw new NotImplementedException();
        }

        public void Repeat()
        {
            throw new NotImplementedException();
        }

        public void ResetTipWizard()
        {
            throw new NotImplementedException();
        }

        public void Save(object Filename)
        {
            throw new NotImplementedException();
        }

        public void SaveWorkspace(object Filename)
        {
            throw new NotImplementedException();
        }

        public void SetDefaultChart(object FormatName, object Gallery)
        {
            throw new NotImplementedException();
        }

        public void Undo()
        {
            throw new NotImplementedException();
        }

        public void Volatile(object Volatile)
        {
            throw new NotImplementedException();
        }

        public void _Wait(object Time)
        {
            throw new NotImplementedException();
        }

        public object _WSFunction(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6,
            object Arg7,
            object Arg8, object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16, object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22,
            object Arg23,
            object Arg24, object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public bool Wait(object Time)
        {
            throw new NotImplementedException();
        }

        public string GetPhonetic(object Text)
        {
            throw new NotImplementedException();
        }

        public void Dummy12(PivotTable p1, PivotTable p2)
        {
            throw new NotImplementedException();
        }

        public void CalculateFull()
        {
            throw new NotImplementedException();
        }

        public bool FindFile()
        {
            throw new NotImplementedException();
        }

        public object Dummy13(object Arg1, object Arg2, object Arg3, object Arg4, object Arg5, object Arg6, object Arg7,
            object Arg8,
            object Arg9, object Arg10, object Arg11, object Arg12, object Arg13, object Arg14, object Arg15,
            object Arg16,
            object Arg17, object Arg18, object Arg19, object Arg20, object Arg21, object Arg22, object Arg23,
            object Arg24,
            object Arg25, object Arg26, object Arg27, object Arg28, object Arg29, object Arg30)
        {
            throw new NotImplementedException();
        }

        public void Dummy14()
        {
            throw new NotImplementedException();
        }

        public void CalculateFullRebuild()
        {
            throw new NotImplementedException();
        }

        public void CheckAbort(object KeepAbort)
        {
            throw new NotImplementedException();
        }

        public void DisplayXMLSourcePane(object XmlMap)
        {
            throw new NotImplementedException();
        }

        public object Support(object Object, int ID, object arg)
        {
            throw new NotImplementedException();
        }

        public object Dummy20(int grfCompareFunctions)
        {
            throw new NotImplementedException();
        }

        public void CalculateUntilAsyncQueriesDone()
        {
            throw new NotImplementedException();
        }

        public int SharePointVersion(string bstrUrl)
        {
            throw new NotImplementedException();
        }

        public void MacroOptions2(object Macro, object Description, object HasMenu, object MenuText,
            object hasShortcutKey,
            object shortcutKey, object category, object statusBar, object helpContextId, object helpFile,
            object argumentDescriptions)
        {
            throw new NotImplementedException();
        }

        public Application Application { get; private set; }
        public XlCreator Creator { get; private set; }
        public Application Parent { get; private set; }
        public Range ActiveCell { get; private set; }
        public Chart ActiveChart { get; private set; }
        public DialogSheet ActiveDialog { get; private set; }
        public MenuBar ActiveMenuBar { get; private set; }
        public string ActivePrinter { get; set; }
        public object ActiveSheet { get; private set; }
        public Window ActiveWindow { get; private set; }
        public Workbook ActiveWorkbook { get; private set; }
        public AddIns AddIns { get; private set; }
        public Assistant Assistant { get; private set; }
        public Range Cells { get; private set; }
        public Sheets Charts { get; private set; }
        public Range Columns { get; private set; }
        public CommandBars CommandBars { get; private set; }
        public int DDEAppReturnCode { get; private set; }
        public Sheets DialogSheets { get; private set; }
        public MenuBars MenuBars { get; private set; }
        public Modules Modules { get; private set; }
        public Names Names { get; private set; }

        public Range get_Range(object Cell1, object Cell2)
        {
            throw new NotImplementedException();
        }

        public Range Rows { get; private set; }
        public object Selection { get; private set; }
        public Sheets Sheets { get; private set; }

        public Menu get_ShortcutMenus(int Index)
        {
            throw new NotImplementedException();
        }

        public Workbook ThisWorkbook { get; private set; }
        public Toolbars Toolbars { get; private set; }
        public Windows Windows { get; private set; }
        public Workbooks Workbooks { get; private set; }
        public WorksheetFunction WorksheetFunction { get; private set; }
        public Sheets Worksheets { get; private set; }
        public Sheets Excel4IntlMacroSheets { get; private set; }
        public Sheets Excel4MacroSheets { get; private set; }
        public bool AlertBeforeOverwriting { get; set; }
        public string AltStartupPath { get; set; }
        public bool AskToUpdateLinks { get; set; }
        public bool EnableAnimations { get; set; }
        public AutoCorrect AutoCorrect { get; private set; }
        public int Build { get; private set; }
        public bool CalculateBeforeSave { get; set; }
        public XlCalculation Calculation { get; set; }

        public object get_Caller(object Index)
        {
            throw new NotImplementedException();
        }

        public bool CanPlaySounds { get; private set; }
        public bool CanRecordSounds { get; private set; }
        public string Caption { get; set; }
        public bool CellDragAndDrop { get; set; }

        public object get_ClipboardFormats(object Index)
        {
            throw new NotImplementedException();
        }

        public bool DisplayClipboardWindow { get; set; }
        public bool ColorButtons { get; set; }
        public XlCommandUnderlines CommandUnderlines { get; set; }
        public bool ConstrainNumeric { get; set; }
        public bool CopyObjectsWithCells { get; set; }
        public XlMousePointer Cursor { get; set; }
        public int CustomListCount { get; private set; }
        public XlCutCopyMode CutCopyMode { get; set; }
        public int DataEntryMode { get; set; }
        public string _Default { get; private set; }
        public string DefaultFilePath { get; set; }
        public Dialogs Dialogs { get; private set; }
        public bool DisplayAlerts { get; set; }
        public bool DisplayFormulaBar { get; set; }
        public bool DisplayFullScreen { get; set; }
        public bool DisplayNoteIndicator { get; set; }
        public XlCommentDisplayMode DisplayCommentIndicator { get; set; }
        public bool DisplayExcel4Menus { get; set; }
        public bool DisplayRecentFiles { get; set; }
        public bool DisplayScrollBars { get; set; }
        public bool DisplayStatusBar { get; set; }
        public bool EditDirectlyInCell { get; set; }
        public bool EnableAutoComplete { get; set; }
        public XlEnableCancelKey EnableCancelKey { get; set; }
        public bool EnableSound { get; set; }
        public bool EnableTipWizard { get; set; }

        public object get_FileConverters(object Index1, object Index2)
        {
            throw new NotImplementedException();
        }

        public FileSearch FileSearch { get; private set; }
        public IFind FileFind { get; private set; }
        public bool FixedDecimal { get; set; }
        public int FixedDecimalPlaces { get; set; }
        public double Height { get; set; }
        public bool IgnoreRemoteRequests { get; set; }
        public bool Interactive { get; set; }

        public object get_International(object Index)
        {
            throw new NotImplementedException();
        }

        public bool Iteration { get; set; }
        public bool LargeButtons { get; set; }
        public double Left { get; set; }
        public string LibraryPath { get; private set; }
        public object MailSession { get; private set; }
        public XlMailSystem MailSystem { get; private set; }
        public bool MathCoprocessorAvailable { get; private set; }
        public double MaxChange { get; set; }
        public int MaxIterations { get; set; }
        public int MemoryFree { get; private set; }
        public int MemoryTotal { get; private set; }
        public int MemoryUsed { get; private set; }
        public bool MouseAvailable { get; private set; }
        public bool MoveAfterReturn { get; set; }
        public XlDirection MoveAfterReturnDirection { get; set; }
        public RecentFiles RecentFiles { get; private set; }
        public string Name { get; private set; }
        public string NetworkTemplatesPath { get; private set; }
        public ODBCErrors ODBCErrors { get; private set; }
        public int ODBCTimeout { get; set; }
        public string OnCalculate { get; set; }
        public string OnData { get; set; }
        public string OnDoubleClick { get; set; }
        public string OnEntry { get; set; }
        public string OnSheetActivate { get; set; }
        public string OnSheetDeactivate { get; set; }
        public string OnWindow { get; set; }
        public string OperatingSystem { get; private set; }
        public string OrganizationName { get; private set; }
        public string Path { get; private set; }
        public string PathSeparator { get; private set; }

        public object get_PreviousSelections(object Index)
        {
            throw new NotImplementedException();
        }

        public bool PivotTableSelection { get; set; }
        public bool PromptForSummaryInfo { get; set; }
        public bool RecordRelative { get; private set; }
        public XlReferenceStyle ReferenceStyle { get; set; }

        public object get_RegisteredFunctions(object Index1, object Index2)
        {
            throw new NotImplementedException();
        }

        public bool RollZoom { get; set; }
        public bool ScreenUpdating { get; set; }
        public int SheetsInNewWorkbook { get; set; }
        public bool ShowChartTipNames { get; set; }
        public bool ShowChartTipValues { get; set; }
        public string StandardFont { get; set; }
        public double StandardFontSize { get; set; }
        public string StartupPath { get; private set; }
        public object StatusBar { get; set; }
        public string TemplatesPath { get; private set; }
        public bool ShowToolTips { get; set; }
        public double Top { get; set; }
        public XlFileFormat DefaultSaveFormat { get; set; }
        public string TransitionMenuKey { get; set; }
        public int TransitionMenuKeyAction { get; set; }
        public bool TransitionNavigKeys { get; set; }
        public double UsableHeight { get; private set; }
        public double UsableWidth { get; private set; }
        public bool UserControl { get; set; }
        public string UserName { get; set; }
        public string Value { get; private set; }
        public VBE VBE { get; private set; }
        public string Version { get; private set; }
        public bool Visible { get; set; }
        public double Width { get; set; }
        public bool WindowsForPens { get; private set; }
        public XlWindowState WindowState { get; set; }
        public int UILanguage { get; set; }
        public int DefaultSheetDirection { get; set; }
        public int CursorMovement { get; set; }
        public bool ControlCharacters { get; set; }
        public bool EnableEvents { get; set; }
        public bool DisplayInfoWindow { get; set; }
        public bool ExtendList { get; set; }
        public OLEDBErrors OLEDBErrors { get; private set; }
        public COMAddIns COMAddIns { get; private set; }
        public DefaultWebOptions DefaultWebOptions { get; private set; }
        public string ProductCode { get; private set; }
        public string UserLibraryPath { get; private set; }
        public bool AutoPercentEntry { get; set; }
        public LanguageSettings LanguageSettings { get; private set; }
        public object Dummy101 { get; private set; }
        public AnswerWizard AnswerWizard { get; private set; }
        public int CalculationVersion { get; private set; }
        public bool ShowWindowsInTaskbar { get; set; }
        public MsoFeatureInstall FeatureInstall { get; set; }
        public bool Ready { get; private set; }
        public CellFormat FindFormat { get; set; }
        public CellFormat ReplaceFormat { get; set; }
        public UsedObjects UsedObjects { get; private set; }
        public XlCalculationState CalculationState { get; private set; }
        public XlCalculationInterruptKey CalculationInterruptKey { get; set; }
        public Watches Watches { get; private set; }
        public bool DisplayFunctionToolTips { get; set; }
        public MsoAutomationSecurity AutomationSecurity { get; set; }
        public bool DisplayPasteOptions { get; set; }
        public bool DisplayInsertOptions { get; set; }
        public bool GenerateGetPivotData { get; set; }
        public AutoRecover AutoRecover { get; private set; }
        public int Hwnd { get; private set; }
        public int Hinstance { get; private set; }
        public ErrorCheckingOptions ErrorCheckingOptions { get; private set; }
        public bool AutoFormatAsYouTypeReplaceHyperlinks { get; set; }
        public SmartTagRecognizers SmartTagRecognizers { get; private set; }

        NewFile _Application.NewWorkbook
        {
            get { throw new NotImplementedException(); }
        }

        public SpellingOptions SpellingOptions { get; private set; }
        public Speech Speech { get; private set; }
        public bool MapPaperSize { get; set; }
        public bool ShowStartupDialog { get; set; }
        public string DecimalSeparator { get; set; }
        public string ThousandsSeparator { get; set; }
        public bool UseSystemSeparators { get; set; }
        public Range ThisCell { get; private set; }
        public RTD RTD { get; private set; }
        public bool DisplayDocumentActionTaskPane { get; set; }
        public bool ArbitraryXMLSupportAvailable { get; private set; }
        public int MeasurementUnit { get; set; }
        public bool ShowSelectionFloaties { get; set; }
        public bool ShowMenuFloaties { get; set; }
        public bool ShowDevTools { get; set; }
        public bool EnableLivePreview { get; set; }
        public bool DisplayDocumentInformationPanel { get; set; }
        public bool AlwaysUseClearType { get; set; }
        public bool WarnOnFunctionNameConflict { get; set; }
        public int FormulaBarHeight { get; set; }
        public bool DisplayFormulaAutoComplete { get; set; }
        public XlGenerateTableRefs GenerateTableRefs { get; set; }
        public IAssistance Assistance { get; private set; }
        public bool EnableLargeOperationAlert { get; set; }
        public int LargeOperationCellThousandCount { get; set; }
        public bool DeferAsyncQueries { get; set; }
        public MultiThreadedCalculation MultiThreadedCalculation { get; private set; }
        public int ActiveEncryptionSession { get; private set; }
        public bool HighQualityModeForGraphics { get; set; }
        public FileExportConverters FileExportConverters { get; private set; }
        public SmartArtLayouts SmartArtLayouts { get; private set; }
        public SmartArtQuickStyles SmartArtQuickStyles { get; private set; }
        public SmartArtColors SmartArtColors { get; private set; }
        public AddIns2 AddIns2 { get; private set; }
        public bool PrintCommunication { get; set; }
        public bool UseClusterConnector { get; set; }
        public string ClusterConnector { get; set; }
        public bool Quitting { get; private set; }
        public bool Dummy22 { get; set; }
        public bool Dummy23 { get; set; }
        public ProtectedViewWindows ProtectedViewWindows { get; private set; }
        public ProtectedViewWindow ActiveProtectedViewWindow { get; private set; }
        public bool IsSandboxed { get; private set; }
        public bool SaveISO8601Dates { get; set; }
        public object HinstancePtr { get; private set; }
        public MsoFileValidationMode FileValidation { get; set; }
        public XlFileValidationPivotMode FileValidationPivot { get; set; }
        public bool ShowQuickAnalysis { get; set; }
        public QuickAnalysis QuickAnalysis { get; private set; }
        public bool FlashFill { get; set; }
        public bool EnableMacroAnimations { get; set; }
        public bool ChartDataPointTrack { get; set; }
        public bool FlashFillMode { get; set; }
        public bool MergeInstances { get; set; }
        public bool EnableCheckFileExtensions { get; set; }

        public FileDialog get_FileDialog(MsoFileDialogType fileDialogType)
        {
            throw new NotImplementedException();
        }

        AppEvents_NewWorkbookEventHandler newWorkbookHandler = doc => { };
        event AppEvents_NewWorkbookEventHandler AppEvents_Event.NewWorkbook
        {
            add { newWorkbookHandler += value; }
            remove { newWorkbookHandler -= value; }
        }

        public event AppEvents_SheetSelectionChangeEventHandler SheetSelectionChange;
        public event AppEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        public event AppEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        public event AppEvents_SheetActivateEventHandler SheetActivate;
        public event AppEvents_SheetDeactivateEventHandler SheetDeactivate;
        public event AppEvents_SheetCalculateEventHandler SheetCalculate;
        public event AppEvents_SheetChangeEventHandler SheetChange;
        public event AppEvents_WorkbookOpenEventHandler WorkbookOpen;
        public event AppEvents_WorkbookActivateEventHandler WorkbookActivate;
        public event AppEvents_WorkbookDeactivateEventHandler WorkbookDeactivate;
        public event AppEvents_WorkbookBeforeCloseEventHandler WorkbookBeforeClose;
        public event AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSave;
        public event AppEvents_WorkbookBeforePrintEventHandler WorkbookBeforePrint;
        public event AppEvents_WorkbookNewSheetEventHandler WorkbookNewSheet;
        public event AppEvents_WorkbookAddinInstallEventHandler WorkbookAddinInstall;
        public event AppEvents_WorkbookAddinUninstallEventHandler WorkbookAddinUninstall;
        public event AppEvents_WindowResizeEventHandler WindowResize;
        public event AppEvents_WindowActivateEventHandler WindowActivate;
        public event AppEvents_WindowDeactivateEventHandler WindowDeactivate;
        public event AppEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlink;
        public event AppEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        public event AppEvents_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnection;
        public event AppEvents_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnection;
        public event AppEvents_WorkbookSyncEventHandler WorkbookSync;
        public event AppEvents_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImport;
        public event AppEvents_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImport;
        public event AppEvents_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExport;
        public event AppEvents_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExport;
        public event AppEvents_WorkbookRowsetCompleteEventHandler WorkbookRowsetComplete;
        public event AppEvents_AfterCalculateEventHandler AfterCalculate;
        public event AppEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChange;
        public event AppEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChanges;
        public event AppEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChanges;
        public event AppEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChanges;
        public event AppEvents_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpen;
        public event AppEvents_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEdit;
        public event AppEvents_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeClose;
        public event AppEvents_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResize;
        public event AppEvents_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivate;
        public event AppEvents_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivate;
        public event AppEvents_WorkbookAfterSaveEventHandler WorkbookAfterSave;
        public event AppEvents_WorkbookNewChartEventHandler WorkbookNewChart;
        public event AppEvents_SheetLensGalleryRenderCompleteEventHandler SheetLensGalleryRenderComplete;
        public event AppEvents_SheetTableUpdateEventHandler SheetTableUpdate;
        public event AppEvents_WorkbookModelChangeEventHandler WorkbookModelChange;
        public event AppEvents_SheetBeforeDeleteEventHandler SheetBeforeDelete;

        public void RaiseNewWorkbook(WorkbookTestDouble workbookTestDouble)
        {
            newWorkbookHandler(workbookTestDouble);
        }

        public void RaiseWorkbookOpen(WorkbookTestDouble workbookTestDouble)
        {
            WorkbookOpen(workbookTestDouble);
        }

        public void RaiseWorkbookBeforeClose(WorkbookTestDouble workbookTestDouble)
        {
            var cancel = false;
            WorkbookBeforeClose(workbookTestDouble, ref cancel);
        }

        public void RaiseWorkbookDeactivate(WorkbookTestDouble workbookTestDouble)
        {
            WorkbookDeactivate(workbookTestDouble);
        }
    }
}