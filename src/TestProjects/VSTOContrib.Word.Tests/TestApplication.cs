using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Word.Application;
using Window = Microsoft.Office.Interop.Word.Window;
using Windows = Microsoft.Office.Interop.Word.Windows;

namespace VSTOContrib.Word.Tests
{
    public class TestApplication : Application
    {
        void _Application.Quit(ref object SaveChanges, ref object originalFormat, ref object RouteDocument)
        {
        }

        public void ScreenRefresh()
        {
        }

        public void PrintOutOld(ref object Background, ref object Append, ref object Range, ref object OutputFileName,
            ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType,
            ref object PrintToFile,
            ref object Collate, ref object FileName, ref object ActivePrinterMacGX, ref object ManualDuplexPrint)
        {
        }

        public void LookupNameProperties(string Name)
        {
        }

        public void SubstituteFont(string UnavailableFont, string SubstituteFont)
        {
        }

        public bool Repeat(ref object Times)
        {
            return false;
        }

        public void DDEExecute(int Channel, string Command)
        {
        }

        public int DDEInitiate(string App, string Topic)
        {
            return 0;
        }

        public void DDEPoke(int Channel, string Item, string Data)
        {
        }

        public string DDERequest(int Channel, string Item)
        {
            return null;
        }

        public void DDETerminate(int Channel)
        {
        }

        public void DDETerminateAll()
        {
        }

        public int BuildKeyCode(WdKey Arg1, ref object Arg2, ref object Arg3, ref object Arg4)
        {
            return 0;
        }

        public string KeyString(int KeyCode, ref object KeyCode2)
        {
            return null;
        }

        public void OrganizerCopy(string Source, string Destination, string Name, WdOrganizerObject Object)
        {
        }

        public void OrganizerDelete(string Source, string Name, WdOrganizerObject Object)
        {
        }

        public void OrganizerRename(string Source, string Name, string NewName, WdOrganizerObject Object)
        {
        }

        public void AddAddress(ref Array TagID, ref Array Value)
        {
        }

        public string GetAddress(ref object Name, ref object AddressProperties, ref object UseAutoText,
            ref object DisplaySelectDialog,
            ref object SelectDialog, ref object CheckNamesDialog, ref object RecentAddressesChoice,
            ref object UpdateRecentAddresses)
        {
            return null;
        }

        public bool CheckGrammar(string String)
        {
            return false;
        }

        public bool CheckSpelling(string Word, ref object CustomDictionary, ref object IgnoreUppercase,
            ref object MainDictionary,
            ref object CustomDictionary2, ref object CustomDictionary3, ref object CustomDictionary4,
            ref object CustomDictionary5, ref object CustomDictionary6, ref object CustomDictionary7,
            ref object CustomDictionary8, ref object CustomDictionary9, ref object CustomDictionary10)
        {
            return false;
        }

        public void ResetIgnoreAll()
        {
        }

        public SpellingSuggestions GetSpellingSuggestions(string Word, ref object CustomDictionary,
            ref object IgnoreUppercase,
            ref object MainDictionary, ref object SuggestionMode, ref object CustomDictionary2,
            ref object CustomDictionary3,
            ref object CustomDictionary4, ref object CustomDictionary5, ref object CustomDictionary6,
            ref object CustomDictionary7, ref object CustomDictionary8, ref object CustomDictionary9,
            ref object CustomDictionary10)
        {
            return null;
        }

        public void GoBack()
        {
        }

        public void Help(ref object HelpType)
        {
        }

        public void AutomaticChange()
        {
        }

        public void ShowMe()
        {
        }

        public void HelpTool()
        {
        }

        public Window NewWindow()
        {
            throw new NotImplementedException();
        }

        public void ListCommands(bool ListAllCommands)
        {
        }

        public void ShowClipboard()
        {
        }

        public void OnTime(ref object When, string Name, ref object Tolerance)
        {
        }

        public void NextLetter()
        {
        }

        public short MountVolume(string Zone, string Server, string Volume, ref object User, ref object UserPassword,
            ref object VolumePassword)
        {
            return 0;
        }

        public string CleanString(string String)
        {
            return null;
        }

        public void SendFax()
        {
        }

        public void ChangeFileOpenDirectory(string Path)
        {
        }

        public void RunOld(string MacroName)
        {
        }

        public void GoForward()
        {
        }

        public void Move(int Left, int Top)
        {
        }

        public void Resize(int Width, int Height)
        {
        }

        public float InchesToPoints(float Inches)
        {
            return 0;
        }

        public float CentimetersToPoints(float Centimeters)
        {
            return 0;
        }

        public float MillimetersToPoints(float Millimeters)
        {
            return 0;
        }

        public float PicasToPoints(float Picas)
        {
            return 0;
        }

        public float LinesToPoints(float Lines)
        {
            return 0;
        }

        public float PointsToInches(float Points)
        {
            return 0;
        }

        public float PointsToCentimeters(float Points)
        {
            return 0;
        }

        public float PointsToMillimeters(float Points)
        {
            return 0;
        }

        public float PointsToPicas(float Points)
        {
            return 0;
        }

        public float PointsToLines(float Points)
        {
            return 0;
        }

        public void Activate()
        {
        }

        public float PointsToPixels(float Points, ref object fVertical)
        {
            return 0;
        }

        public float PixelsToPoints(float Pixels, ref object fVertical)
        {
            return 0;
        }

        public void KeyboardLatin()
        {
        }

        public void KeyboardBidi()
        {
        }

        public void ToggleKeyboard()
        {
        }

        public int Keyboard(int LangId = 0)
        {
            return 0;
        }

        public string ProductCode()
        {
            return null;
        }

        public DefaultWebOptions DefaultWebOptions()
        {
            return null;
        }

        public void DiscussionSupport(ref object Range, ref object cid, ref object piCSE)
        {
        }

        public void SetDefaultTheme(string Name, WdDocumentMedium DocumentType)
        {
        }

        public string GetDefaultTheme(WdDocumentMedium DocumentType)
        {
            return null;
        }

        public void PrintOut2000(ref object Background, ref object Append, ref object Range, ref object OutputFileName,
            ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType,
            ref object PrintToFile,
            ref object Collate, ref object FileName, ref object ActivePrinterMacGX, ref object ManualDuplexPrint,
            ref object PrintZoomColumn, ref object PrintZoomRow, ref object PrintZoomPaperWidth,
            ref object PrintZoomPaperHeight)
        {
        }

        public object Run(string MacroName, ref object varg1, ref object varg2, ref object varg3, ref object varg4,
            ref object varg5,
            ref object varg6, ref object varg7, ref object varg8, ref object varg9, ref object varg10, ref object varg11,
            ref object varg12, ref object varg13, ref object varg14, ref object varg15, ref object varg16,
            ref object varg17,
            ref object varg18, ref object varg19, ref object varg20, ref object varg21, ref object varg22,
            ref object varg23,
            ref object varg24, ref object varg25, ref object varg26, ref object varg27, ref object varg28,
            ref object varg29,
            ref object varg30)
        {
            return null;
        }

        public void PrintOut(ref object Background, ref object Append, ref object Range, ref object OutputFileName,
            ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType,
            ref object PrintToFile,
            ref object Collate, ref object FileName, ref object ActivePrinterMacGX, ref object ManualDuplexPrint,
            ref object PrintZoomColumn, ref object PrintZoomRow, ref object PrintZoomPaperWidth,
            ref object PrintZoomPaperHeight)
        {
        }

        public bool Dummy2()
        {
            return false;
        }

        public void PutFocusInMailHeader()
        {
        }

        public void LoadMasterList(string FileName)
        {
        }

        public Document CompareDocuments(Document OriginalDocument, Document RevisedDocument,
            WdCompareDestination Destination = WdCompareDestination.wdCompareDestinationNew,
            WdGranularity Granularity = WdGranularity.wdGranularityWordLevel,
            bool CompareFormatting = true, bool CompareCaseChanges = true, bool CompareWhitespace = true,
            bool CompareTables = true, bool CompareHeaders = true, bool CompareFootnotes = true,
            bool CompareTextboxes = true,
            bool CompareFields = true, bool CompareComments = true, bool CompareMoves = true, string RevisedAuthor = "",
            bool IgnoreAllComparisonWarnings = false)
        {
            return null;
        }

        public Document MergeDocuments(Document OriginalDocument, Document RevisedDocument,
            WdCompareDestination Destination = WdCompareDestination.wdCompareDestinationNew,
            WdGranularity Granularity = WdGranularity.wdGranularityWordLevel,
            bool CompareFormatting = true, bool CompareCaseChanges = true, bool CompareWhitespace = true,
            bool CompareTables = true, bool CompareHeaders = true, bool CompareFootnotes = true,
            bool CompareTextboxes = true,
            bool CompareFields = true, bool CompareComments = true, bool CompareMoves = true, string OriginalAuthor = "",
            string RevisedAuthor = "", WdMergeFormatFrom FormatFrom = WdMergeFormatFrom.wdMergeFormatFromPrompt)
        {
            return null;
        }

        public void ThreeWayMerge(Document LocalDocument, Document ServerDocument, Document BaseDocument,
            bool FavorSource)
        {
        }

        public void Dummy4()
        {
        }

        public Application Application { get; private set; }
        public int Creator { get; private set; }
        public object Parent { get; private set; }
        public string Name { get; private set; }
        public Documents Documents { get; set; }
        public Windows Windows { get; private set; }
        public Document ActiveDocument { get; private set; }
        public Window ActiveWindow { get; private set; }
        public Selection Selection { get; private set; }
        public object WordBasic { get; private set; }
        public RecentFiles RecentFiles { get; private set; }
        public Template NormalTemplate { get; private set; }
        public Microsoft.Office.Interop.Word.System System { get; private set; }
        public AutoCorrect AutoCorrect { get; private set; }
        public FontNames FontNames { get; private set; }
        public FontNames LandscapeFontNames { get; private set; }
        public FontNames PortraitFontNames { get; private set; }
        public Languages Languages { get; private set; }
        public Assistant Assistant { get; private set; }
        public Browser Browser { get; private set; }
        public FileConverters FileConverters { get; private set; }
        public MailingLabel MailingLabel { get; private set; }
        public Dialogs Dialogs { get; private set; }
        public CaptionLabels CaptionLabels { get; private set; }
        public AutoCaptions AutoCaptions { get; private set; }
        public AddIns AddIns { get; private set; }
        public bool Visible { get; set; }
        public string Version { get; private set; }
        public bool ScreenUpdating { get; set; }
        public bool PrintPreview { get; set; }
        public Tasks Tasks { get; private set; }
        public bool DisplayStatusBar { get; set; }
        public bool SpecialMode { get; private set; }
        public int UsableWidth { get; private set; }
        public int UsableHeight { get; private set; }
        public bool MathCoprocessorAvailable { get; private set; }
        public bool MouseAvailable { get; private set; }

        public object get_International(WdInternationalIndex Index)
        {
            return null;
        }

        public string Build { get; private set; }
        public bool CapsLock { get; private set; }
        public bool NumLock { get; private set; }
        public string UserName { get; set; }
        public string UserInitials { get; set; }
        public string UserAddress { get; set; }
        public object MacroContainer { get; private set; }
        public bool DisplayRecentFiles { get; set; }
        public CommandBars CommandBars { get; private set; }

        public SynonymInfo get_SynonymInfo(string Word, ref object LanguageID)
        {
            return null;
        }

        public VBE VBE { get; private set; }
        public string DefaultSaveFormat { get; set; }
        public ListGalleries ListGalleries { get; private set; }
        public string ActivePrinter { get; set; }
        public Templates Templates { get; private set; }
        public object CustomizationContext { get; set; }
        public KeyBindings KeyBindings { get; private set; }

        public KeysBoundTo get_KeysBoundTo(WdKeyCategory KeyCategory, string Command, ref object CommandParameter)
        {
            return null;
        }

        public KeyBinding get_FindKey(int KeyCode, ref object KeyCode2)
        {
            return null;
        }

        public string Caption { get; set; }
        public string Path { get; private set; }
        public bool DisplayScrollBars { get; set; }
        public string StartupPath { get; set; }
        public int BackgroundSavingStatus { get; private set; }
        public int BackgroundPrintingStatus { get; private set; }
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public WdWindowState WindowState { get; set; }
        public bool DisplayAutoCompleteTips { get; set; }
        public Options Options { get; private set; }
        public WdAlertLevel DisplayAlerts { get; set; }
        public Dictionaries CustomDictionaries { get; private set; }
        public string PathSeparator { get; private set; }
        public string StatusBar { set; private get; }
        public bool MAPIAvailable { get; private set; }
        public bool DisplayScreenTips { get; set; }
        public WdEnableCancelKey EnableCancelKey { get; set; }
        public bool UserControl { get; private set; }
        public FileSearch FileSearch { get; private set; }
        public WdMailSystem MailSystem { get; private set; }
        public string DefaultTableSeparator { get; set; }
        public bool ShowVisualBasicEditor { get; set; }
        public string BrowseExtraFileTypes { get; set; }

        public bool get_IsObjectValid(object Object)
        {
            return false;
        }

        public HangulHanjaConversionDictionaries HangulHanjaDictionaries { get; private set; }
        public MailMessage MailMessage { get; private set; }
        public bool FocusInMailHeader { get; private set; }
        public EmailOptions EmailOptions { get; private set; }
        public MsoLanguageID Language { get; private set; }
        public COMAddIns COMAddIns { get; private set; }
        public bool CheckLanguage { get; set; }
        public LanguageSettings LanguageSettings { get; private set; }
        public bool Dummy1 { get; private set; }
        public AnswerWizard AnswerWizard { get; private set; }
        public MsoFeatureInstall FeatureInstall { get; set; }
        public MsoAutomationSecurity AutomationSecurity { get; set; }

        public FileDialog get_FileDialog(MsoFileDialogType FileDialogType)
        {
            return null;
        }

        public string EmailTemplate { get; set; }
        public bool ShowWindowsInTaskbar { get; set; }

        NewFile _Application.NewDocument
        {
            get { return null; }
        }

        public bool ShowStartupDialog { get; set; }
        public AutoCorrect AutoCorrectEmail { get; private set; }
        public TaskPanes TaskPanes { get; private set; }
        public bool DefaultLegalBlackline { get; set; }
        public SmartTagRecognizers SmartTagRecognizers { get; private set; }
        public SmartTagTypes SmartTagTypes { get; private set; }
        public XMLNamespaces XMLNamespaces { get; private set; }
        public bool ArbitraryXMLSupportAvailable { get; private set; }
        public string BuildFull { get; private set; }
        public string BuildFeatureCrew { get; private set; }
        public Bibliography Bibliography { get; private set; }
        public bool ShowStylePreviews { get; set; }
        public bool RestrictLinkedStyles { get; set; }
        public OMathAutoCorrect OMathAutoCorrect { get; private set; }
        public bool DisplayDocumentInformationPanel { get; set; }
        public IAssistance Assistance { get; private set; }
        public bool OpenAttachmentsInFullScreen { get; set; }
        public int ActiveEncryptionSession { get; private set; }
        public bool DontResetInsertionPointProperties { get; set; }
        public SmartArtLayouts SmartArtLayouts { get; private set; }
        public SmartArtQuickStyles SmartArtQuickStyles { get; private set; }
        public SmartArtColors SmartArtColors { get; private set; }
        public UndoRecord UndoRecord { get; private set; }
        public PickerDialog PickerDialog { get; private set; }
        public ProtectedViewWindows ProtectedViewWindows { get; private set; }
        public ProtectedViewWindow ActiveProtectedViewWindow { get; private set; }
        public bool IsSandboxed { get; private set; }
        public MsoFileValidationMode FileValidation { get; set; }
        public bool ChartDataPointTrack { get; set; }
        public bool ShowAnimation { get; set; }
        public event ApplicationEvents4_StartupEventHandler Startup;
        public event ApplicationEvents4_QuitEventHandler Quit;
        public event ApplicationEvents4_DocumentChangeEventHandler DocumentChange;
        public event ApplicationEvents4_DocumentOpenEventHandler DocumentOpen;
        public event ApplicationEvents4_DocumentBeforeCloseEventHandler DocumentBeforeClose;
        public event ApplicationEvents4_DocumentBeforePrintEventHandler DocumentBeforePrint;
        public event ApplicationEvents4_DocumentBeforeSaveEventHandler DocumentBeforeSave;

        event ApplicationEvents4_NewDocumentEventHandler ApplicationEvents4_Event.NewDocument
        {
            add { }
            remove { }
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
    }
}