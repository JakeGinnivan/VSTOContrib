using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Word.Application;
using Window = Microsoft.Office.Interop.Word.Window;
using Windows = Microsoft.Office.Interop.Word.Windows;

namespace Word.TestDoubles
{
    public class ApplicationTestDouble : Application
    {
        public ApplicationTestDouble()
        {
            Documents = new DocumentsTestDouble(this);
            Windows = new WindowsTestDouble(this);
        }


        void _Application.Quit(ref object saveChanges, ref object originalFormat, ref object routeDocument)
        {
            throw new NotImplementedException();
        }

        public void ScreenRefresh()
        {
            throw new NotImplementedException();
        }

        public void PrintOutOld(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object FileName, ref object ActivePrinterMacGX, ref object ManualDuplexPrint)
        {
            throw new NotImplementedException();
        }

        public void LookupNameProperties(string Name)
        {
            throw new NotImplementedException();
        }

        public void SubstituteFont(string UnavailableFont, string SubstituteFont)
        {
            throw new NotImplementedException();
        }

        public bool Repeat(ref object Times)
        {
            throw new NotImplementedException();
        }

        public void DDEExecute(int Channel, string Command)
        {
            throw new NotImplementedException();
        }

        public int DDEInitiate(string App, string Topic)
        {
            throw new NotImplementedException();
        }

        public void DDEPoke(int Channel, string Item, string Data)
        {
            throw new NotImplementedException();
        }

        public string DDERequest(int Channel, string Item)
        {
            throw new NotImplementedException();
        }

        public void DDETerminate(int Channel)
        {
            throw new NotImplementedException();
        }

        public void DDETerminateAll()
        {
            throw new NotImplementedException();
        }

        public int BuildKeyCode(WdKey Arg1, ref object Arg2, ref object Arg3, ref object Arg4)
        {
            throw new NotImplementedException();
        }

        public string KeyString(int KeyCode, ref object KeyCode2)
        {
            throw new NotImplementedException();
        }

        public void OrganizerCopy(string Source, string Destination, string Name, WdOrganizerObject Object)
        {
            throw new NotImplementedException();
        }

        public void OrganizerDelete(string Source, string Name, WdOrganizerObject Object)
        {
            throw new NotImplementedException();
        }

        public void OrganizerRename(string Source, string Name, string NewName, WdOrganizerObject Object)
        {
            throw new NotImplementedException();
        }

        public void AddAddress(ref Array TagID, ref Array Value)
        {
            throw new NotImplementedException();
        }

        public string GetAddress(ref object Name, ref object AddressProperties, ref object UseAutoText, ref object DisplaySelectDialog,
            ref object SelectDialog, ref object CheckNamesDialog, ref object RecentAddressesChoice,
            ref object UpdateRecentAddresses)
        {
            throw new NotImplementedException();
        }

        public bool CheckGrammar(string String)
        {
            throw new NotImplementedException();
        }

        public bool CheckSpelling(string Word, ref object CustomDictionary, ref object IgnoreUppercase, ref object MainDictionary,
            ref object CustomDictionary2, ref object CustomDictionary3, ref object CustomDictionary4,
            ref object CustomDictionary5, ref object CustomDictionary6, ref object CustomDictionary7,
            ref object CustomDictionary8, ref object CustomDictionary9, ref object CustomDictionary10)
        {
            throw new NotImplementedException();
        }

        public void ResetIgnoreAll()
        {
            throw new NotImplementedException();
        }

        public SpellingSuggestions GetSpellingSuggestions(string Word, ref object CustomDictionary, ref object IgnoreUppercase,
            ref object MainDictionary, ref object SuggestionMode, ref object CustomDictionary2, ref object CustomDictionary3,
            ref object CustomDictionary4, ref object CustomDictionary5, ref object CustomDictionary6,
            ref object CustomDictionary7, ref object CustomDictionary8, ref object CustomDictionary9,
            ref object CustomDictionary10)
        {
            throw new NotImplementedException();
        }

        public void GoBack()
        {
            throw new NotImplementedException();
        }

        public void Help(ref object HelpType)
        {
            throw new NotImplementedException();
        }

        public void AutomaticChange()
        {
            throw new NotImplementedException();
        }

        public void ShowMe()
        {
            throw new NotImplementedException();
        }

        public void HelpTool()
        {
            throw new NotImplementedException();
        }

        public Window NewWindow()
        {
            var windowTestDouble = new WindowTestDouble();
            Windows.Add(windowTestDouble);
            return windowTestDouble;
        }

        public void ListCommands(bool listAllCommands)
        {
            throw new NotImplementedException();
        }

        public void ShowClipboard()
        {
            throw new NotImplementedException();
        }

        public void OnTime(ref object when, string name, ref object tolerance)
        {
            throw new NotImplementedException();
        }

        public void NextLetter()
        {
            throw new NotImplementedException();
        }

        public short MountVolume(string zone, string Server, string Volume, ref object User, ref object UserPassword,
            ref object VolumePassword)
        {
            throw new NotImplementedException();
        }

        public string CleanString(string String)
        {
            throw new NotImplementedException();
        }

        public void SendFax()
        {
            throw new NotImplementedException();
        }

        public void ChangeFileOpenDirectory(string Path)
        {
            throw new NotImplementedException();
        }

        public void RunOld(string MacroName)
        {
            throw new NotImplementedException();
        }

        public void GoForward()
        {
            throw new NotImplementedException();
        }

        public void Move(int Left, int Top)
        {
            throw new NotImplementedException();
        }

        public void Resize(int Width, int Height)
        {
            throw new NotImplementedException();
        }

        public float InchesToPoints(float Inches)
        {
            throw new NotImplementedException();
        }

        public float CentimetersToPoints(float Centimeters)
        {
            throw new NotImplementedException();
        }

        public float MillimetersToPoints(float Millimeters)
        {
            throw new NotImplementedException();
        }

        public float PicasToPoints(float Picas)
        {
            throw new NotImplementedException();
        }

        public float LinesToPoints(float Lines)
        {
            throw new NotImplementedException();
        }

        public float PointsToInches(float Points)
        {
            throw new NotImplementedException();
        }

        public float PointsToCentimeters(float Points)
        {
            throw new NotImplementedException();
        }

        public float PointsToMillimeters(float Points)
        {
            throw new NotImplementedException();
        }

        public float PointsToPicas(float Points)
        {
            throw new NotImplementedException();
        }

        public float PointsToLines(float Points)
        {
            throw new NotImplementedException();
        }

        public void Activate()
        {
            throw new NotImplementedException();
        }

        public float PointsToPixels(float Points, ref object fVertical)
        {
            throw new NotImplementedException();
        }

        public float PixelsToPoints(float Pixels, ref object fVertical)
        {
            throw new NotImplementedException();
        }

        public void KeyboardLatin()
        {
            throw new NotImplementedException();
        }

        public void KeyboardBidi()
        {
            throw new NotImplementedException();
        }

        public void ToggleKeyboard()
        {
            throw new NotImplementedException();
        }

        public int Keyboard(int LangId = 0)
        {
            throw new NotImplementedException();
        }

        public string ProductCode()
        {
            throw new NotImplementedException();
        }

        public DefaultWebOptions DefaultWebOptions()
        {
            throw new NotImplementedException();
        }

        public void DiscussionSupport(ref object Range, ref object cid, ref object piCSE)
        {
            throw new NotImplementedException();
        }

        public void SetDefaultTheme(string Name, WdDocumentMedium DocumentType)
        {
            throw new NotImplementedException();
        }

        public string GetDefaultTheme(WdDocumentMedium DocumentType)
        {
            throw new NotImplementedException();
        }

        public void PrintOut2000(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object FileName, ref object ActivePrinterMacGX, ref object ManualDuplexPrint,
            ref object PrintZoomColumn, ref object PrintZoomRow, ref object PrintZoomPaperWidth, ref object PrintZoomPaperHeight)
        {
            throw new NotImplementedException();
        }

        public object Run(string MacroName, ref object varg1, ref object varg2, ref object varg3, ref object varg4, ref object varg5,
            ref object varg6, ref object varg7, ref object varg8, ref object varg9, ref object varg10, ref object varg11,
            ref object varg12, ref object varg13, ref object varg14, ref object varg15, ref object varg16, ref object varg17,
            ref object varg18, ref object varg19, ref object varg20, ref object varg21, ref object varg22, ref object varg23,
            ref object varg24, ref object varg25, ref object varg26, ref object varg27, ref object varg28, ref object varg29,
            ref object varg30)
        {
            throw new NotImplementedException();
        }

        public void PrintOut(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object FileName, ref object ActivePrinterMacGX, ref object ManualDuplexPrint,
            ref object PrintZoomColumn, ref object PrintZoomRow, ref object PrintZoomPaperWidth, ref object PrintZoomPaperHeight)
        {
            throw new NotImplementedException();
        }

        public bool Dummy2()
        {
            throw new NotImplementedException();
        }

        public void PutFocusInMailHeader()
        {
            throw new NotImplementedException();
        }

        public void LoadMasterList(string FileName)
        {
            throw new NotImplementedException();
        }

        public Document CompareDocuments(Document OriginalDocument, Document RevisedDocument,
            WdCompareDestination Destination = WdCompareDestination.wdCompareDestinationNew, WdGranularity Granularity = WdGranularity.wdGranularityWordLevel,
            bool CompareFormatting = true, bool CompareCaseChanges = true, bool CompareWhitespace = true,
            bool CompareTables = true, bool CompareHeaders = true, bool CompareFootnotes = true, bool CompareTextboxes = true,
            bool CompareFields = true, bool CompareComments = true, bool CompareMoves = true, string RevisedAuthor = "",
            bool IgnoreAllComparisonWarnings = false)
        {
            throw new NotImplementedException();
        }

        public Document MergeDocuments(Document OriginalDocument, Document RevisedDocument,
            WdCompareDestination Destination = WdCompareDestination.wdCompareDestinationNew, WdGranularity Granularity = WdGranularity.wdGranularityWordLevel,
            bool CompareFormatting = true, bool CompareCaseChanges = true, bool CompareWhitespace = true,
            bool CompareTables = true, bool CompareHeaders = true, bool CompareFootnotes = true, bool CompareTextboxes = true,
            bool CompareFields = true, bool CompareComments = true, bool CompareMoves = true, string OriginalAuthor = "",
            string RevisedAuthor = "", WdMergeFormatFrom FormatFrom = WdMergeFormatFrom.wdMergeFormatFromPrompt)
        {
            throw new NotImplementedException();
        }

        public void ThreeWayMerge(Document localDocument, Document serverDocument, Document baseDocument, bool favorSource)
        {
            throw new NotImplementedException();
        }

        public void Dummy4()
        {
            throw new NotImplementedException();
        }

        public Application Application { get { return this; } }
        public int Creator { get; set; }
        public object Parent { get; set; }
        public string Name { get; set; }
        public Documents Documents { get; private set; }
        public Windows Windows { get; private set; }
        public Document ActiveDocument { get; set; }
        public Window ActiveWindow { get; set; }
        public Selection Selection { get; set; }
        public object WordBasic { get; set; }
        public RecentFiles RecentFiles { get; set; }
        public Template NormalTemplate { get; set; }
        public Microsoft.Office.Interop.Word.System System { get; set; }
        public AutoCorrect AutoCorrect { get; set; }
        public FontNames FontNames { get; set; }
        public FontNames LandscapeFontNames { get; set; }
        public FontNames PortraitFontNames { get; set; }
        public Languages Languages { get; set; }
        public Assistant Assistant { get; set; }
        public Browser Browser { get; set; }
        public FileConverters FileConverters { get; set; }
        public MailingLabel MailingLabel { get; set; }
        public Dialogs Dialogs { get; set; }
        public CaptionLabels CaptionLabels { get; set; }
        public AutoCaptions AutoCaptions { get; set; }
        public AddIns AddIns { get; set; }
        public bool Visible { get; set; }
        public string Version { get; set; }
        public bool ScreenUpdating { get; set; }
        public bool PrintPreview { get; set; }
        public Tasks Tasks { get; set; }
        public bool DisplayStatusBar { get; set; }
        public bool SpecialMode { get; set; }
        public int UsableWidth { get; set; }
        public int UsableHeight { get; set; }
        public bool MathCoprocessorAvailable { get; set; }
        public bool MouseAvailable { get; set; }
        public object get_International(WdInternationalIndex Index)
        {
            throw new NotImplementedException();
        }

        public string Build { get; set; }
        public bool CapsLock { get; set; }
        public bool NumLock { get; set; }
        public string UserName { get; set; }
        public string UserInitials { get; set; }
        public string UserAddress { get; set; }
        public object MacroContainer { get; set; }
        public bool DisplayRecentFiles { get; set; }
        public CommandBars CommandBars { get; set; }
        public SynonymInfo get_SynonymInfo(string Word, ref object LanguageID)
        {
            throw new NotImplementedException();
        }

        public VBE VBE { get; set; }
        public string DefaultSaveFormat { get; set; }
        public ListGalleries ListGalleries { get; set; }
        public string ActivePrinter { get; set; }
        public Templates Templates { get; set; }
        public object CustomizationContext { get; set; }
        public KeyBindings KeyBindings { get; set; }
        public KeysBoundTo get_KeysBoundTo(WdKeyCategory KeyCategory, string Command, ref object CommandParameter)
        {
            throw new NotImplementedException();
        }

        public KeyBinding get_FindKey(int keyCode, ref object KeyCode2)
        {
            throw new NotImplementedException();
        }

        public string Caption { get; set; }
        public string Path { get; set; }
        public bool DisplayScrollBars { get; set; }
        public string StartupPath { get; set; }
        public int BackgroundSavingStatus { get; set; }
        public int BackgroundPrintingStatus { get; set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public WdWindowState WindowState { get; set; }
        public bool DisplayAutoCompleteTips { get; set; }
        public Options Options { get; set; }
        public WdAlertLevel DisplayAlerts { get; set; }
        public Dictionaries CustomDictionaries { get; set; }
        public string PathSeparator { get; set; }
        public string StatusBar { set; get; }
        public bool MAPIAvailable { get; set; }
        public bool DisplayScreenTips { get; set; }
        public WdEnableCancelKey EnableCancelKey { get; set; }
        public bool UserControl { get; set; }
        public FileSearch FileSearch { get; set; }
        public WdMailSystem MailSystem { get; set; }
        public string DefaultTableSeparator { get; set; }
        public bool ShowVisualBasicEditor { get; set; }
        public string BrowseExtraFileTypes { get; set; }
        public bool get_IsObjectValid(object Object)
        {
            throw new NotImplementedException();
        }

        public HangulHanjaConversionDictionaries HangulHanjaDictionaries { get; set; }
        public MailMessage MailMessage { get; set; }
        public bool FocusInMailHeader { get; set; }
        public EmailOptions EmailOptions { get; set; }
        public MsoLanguageID Language { get; set; }
        public COMAddIns COMAddIns { get; set; }
        public bool CheckLanguage { get; set; }
        public LanguageSettings LanguageSettings { get; set; }
        public bool Dummy1 { get; set; }
        public AnswerWizard AnswerWizard { get; set; }
        public MsoFeatureInstall FeatureInstall { get; set; }
        public MsoAutomationSecurity AutomationSecurity { get; set; }
        public FileDialog get_FileDialog(MsoFileDialogType fileDialogType)
        {
            throw new NotImplementedException();
        }

        public string EmailTemplate { get; set; }
        public bool ShowWindowsInTaskbar { get; set; }
        public NewFile NewDocument
        {
            get { return ((_Application)Application).NewDocument; }
        }

        public bool ShowStartupDialog { get; set; }
        public AutoCorrect AutoCorrectEmail { get; set; }
        public TaskPanes TaskPanes { get; set; }
        public bool DefaultLegalBlackline { get; set; }
        public SmartTagRecognizers SmartTagRecognizers { get; set; }
        public SmartTagTypes SmartTagTypes { get; set; }
        public XMLNamespaces XMLNamespaces { get; set; }
        public bool ArbitraryXMLSupportAvailable { get; set; }
        public string BuildFull { get; set; }
        public string BuildFeatureCrew { get; set; }
        public Bibliography Bibliography { get; set; }
        public bool ShowStylePreviews { get; set; }
        public bool RestrictLinkedStyles { get; set; }
        public OMathAutoCorrect OMathAutoCorrect { get; set; }
        public bool DisplayDocumentInformationPanel { get; set; }
        public IAssistance Assistance { get; set; }
        public bool OpenAttachmentsInFullScreen { get; set; }
        public int ActiveEncryptionSession { get; set; }
        public bool DontResetInsertionPointProperties { get; set; }
        public SmartArtLayouts SmartArtLayouts { get; set; }
        public SmartArtQuickStyles SmartArtQuickStyles { get; set; }
        public SmartArtColors SmartArtColors { get; set; }
        public UndoRecord UndoRecord { get; set; }
        public PickerDialog PickerDialog { get; set; }
        public ProtectedViewWindows ProtectedViewWindows { get; set; }
        public ProtectedViewWindow ActiveProtectedViewWindow { get; set; }
        public bool IsSandboxed { get; set; }
        public MsoFileValidationMode FileValidation { get; set; }
        public bool ChartDataPointTrack { get; set; }
        public bool ShowAnimation { get; set; }
        public event ApplicationEvents4_StartupEventHandler Startup;
        public event ApplicationEvents4_QuitEventHandler Quit;
        public event ApplicationEvents4_DocumentChangeEventHandler DocumentChange;
        public event ApplicationEvents4_DocumentOpenEventHandler DocumentOpen = doc => { };
        public event ApplicationEvents4_DocumentBeforeCloseEventHandler DocumentBeforeClose;
        public event ApplicationEvents4_DocumentBeforePrintEventHandler DocumentBeforePrint;
        public event ApplicationEvents4_DocumentBeforeSaveEventHandler DocumentBeforeSave;
        ApplicationEvents4_NewDocumentEventHandler newDocumentHandler = doc => { };
        event ApplicationEvents4_NewDocumentEventHandler ApplicationEvents4_Event.NewDocument
        {
            add { newDocumentHandler += value; }
            remove { newDocumentHandler -= value; }
        }

        public event ApplicationEvents4_WindowActivateEventHandler WindowActivate;
        public event ApplicationEvents4_WindowDeactivateEventHandler WindowDeactivate;
        public event ApplicationEvents4_WindowSelectionChangeEventHandler WindowSelectionChange;
        public event ApplicationEvents4_WindowBeforeRightClickEventHandler WindowBeforeRightClick;
        public event ApplicationEvents4_WindowBeforeDoubleClickEventHandler WindowBeforeDoubleClick;
        public event ApplicationEvents4_EPostagePropertyDialogEventHandler EPostagePropertyDialog;
        public event ApplicationEvents4_EPostageInsertEventHandler EPostageInsert;
        public event ApplicationEvents4_MailMergeAfterMergeEventHandler MailMergeAfterMerge;
        public event ApplicationEvents4_MailMergeAfterRecordMergeEventHandler MailMergeAfterRecordMerge;
        public event ApplicationEvents4_MailMergeBeforeMergeEventHandler MailMergeBeforeMerge;
        public event ApplicationEvents4_MailMergeBeforeRecordMergeEventHandler MailMergeBeforeRecordMerge;
        public event ApplicationEvents4_MailMergeDataSourceLoadEventHandler MailMergeDataSourceLoad;
        public event ApplicationEvents4_MailMergeDataSourceValidateEventHandler MailMergeDataSourceValidate;
        public event ApplicationEvents4_MailMergeWizardSendToCustomEventHandler MailMergeWizardSendToCustom;
        public event ApplicationEvents4_MailMergeWizardStateChangeEventHandler MailMergeWizardStateChange;
        public event ApplicationEvents4_WindowSizeEventHandler WindowSize;
        public event ApplicationEvents4_XMLSelectionChangeEventHandler XMLSelectionChange;
        public event ApplicationEvents4_XMLValidationErrorEventHandler XMLValidationError;
        public event ApplicationEvents4_DocumentSyncEventHandler DocumentSync;
        public event ApplicationEvents4_EPostageInsertExEventHandler EPostageInsertEx;
        public event ApplicationEvents4_MailMergeDataSourceValidate2EventHandler MailMergeDataSourceValidate2;
        public event ApplicationEvents4_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpen;
        public event ApplicationEvents4_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEdit;
        public event ApplicationEvents4_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeClose;
        public event ApplicationEvents4_ProtectedViewWindowSizeEventHandler ProtectedViewWindowSize;
        public event ApplicationEvents4_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivate;
        public event ApplicationEvents4_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivate;

        public void OnDocumentOpen(DocumentTestDouble document)
        {
            DocumentOpen(document);
        }
    }
}