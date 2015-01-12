using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Word.Application;
using Shape = Microsoft.Office.Interop.Word.Shape;
using Shapes = Microsoft.Office.Interop.Word.Shapes;
using Window = Microsoft.Office.Interop.Word.Window;
using Windows = Microsoft.Office.Interop.Word.Windows;

namespace Word.TestDoubles
{
#pragma warning disable 0067
    public class DocumentTestDouble : Document
    {
        public DocumentTestDouble(ApplicationTestDouble application, WindowTestDouble window)
        {
            Windows = new WindowsTestDouble(application);
            Windows.Add(window);
            Application = application;
        }

        void _Document.Close(ref object saveChanges, ref object originalFormat, ref object routeDocument)
        {
            throw new System.NotImplementedException();
        }

        public void SaveAs2000(ref object fileName, ref object fileFormat, ref object lockComments, ref object password,
            ref object addToRecentFiles, ref object writePassword, ref object readOnlyRecommended, ref object embedTrueTypeFonts,
            ref object saveNativePictureFormat, ref object saveFormsData, ref object saveAsAoceLetter)
        {
            throw new System.NotImplementedException();
        }

        public void Repaginate()
        {
            throw new System.NotImplementedException();
        }

        public void FitToPages()
        {
            throw new System.NotImplementedException();
        }

        public void ManualHyphenation()
        {
            throw new System.NotImplementedException();
        }

        public void Select()
        {
            throw new System.NotImplementedException();
        }

        public void DataForm()
        {
            throw new System.NotImplementedException();
        }

        public void Route()
        {
            throw new System.NotImplementedException();
        }

        public void Save()
        {
            throw new System.NotImplementedException();
        }

        public void PrintOutOld(ref object background, ref object append, ref object range, ref object outputFileName, ref object @from,
            ref object to, ref object item, ref object copies, ref object pages, ref object pageType, ref object printToFile,
            ref object collate, ref object activePrinterMacGx, ref object manualDuplexPrint)
        {
            throw new System.NotImplementedException();
        }

        public void SendMail()
        {
            throw new System.NotImplementedException();
        }

        public Range Range(ref object start, ref object end)
        {
            throw new System.NotImplementedException();
        }

        public void RunAutoMacro(WdAutoMacros which)
        {
            throw new System.NotImplementedException();
        }

        public void Activate()
        {
            throw new System.NotImplementedException();
        }

        public void PrintPreview()
        {
            throw new System.NotImplementedException();
        }

        public Range GoTo(ref object what, ref object which, ref object count, ref object name)
        {
            throw new System.NotImplementedException();
        }

        public bool Undo(ref object times)
        {
            throw new System.NotImplementedException();
        }

        public bool Redo(ref object Times)
        {
            throw new System.NotImplementedException();
        }

        public int ComputeStatistics(WdStatistic Statistic, ref object IncludeFootnotesAndEndnotes)
        {
            throw new System.NotImplementedException();
        }

        public void MakeCompatibilityDefault()
        {
            throw new System.NotImplementedException();
        }

        public void Protect2002(WdProtectionType Type, ref object NoReset, ref object Password)
        {
            throw new System.NotImplementedException();
        }

        public void Unprotect(ref object Password)
        {
            throw new System.NotImplementedException();
        }

        public void EditionOptions(WdEditionType Type, WdEditionOption Option, string Name, ref object Format)
        {
            throw new System.NotImplementedException();
        }

        public void RunLetterWizard(ref object LetterContent, ref object WizardMode)
        {
            throw new System.NotImplementedException();
        }

        public LetterContent GetLetterContent()
        {
            throw new System.NotImplementedException();
        }

        public void SetLetterContent(ref object LetterContent)
        {
            throw new System.NotImplementedException();
        }

        public void CopyStylesFromTemplate(string Template)
        {
            throw new System.NotImplementedException();
        }

        public void UpdateStyles()
        {
            throw new System.NotImplementedException();
        }

        public void CheckGrammar()
        {
            throw new System.NotImplementedException();
        }

        public void CheckSpelling(ref object CustomDictionary, ref object IgnoreUppercase, ref object AlwaysSuggest,
            ref object CustomDictionary2, ref object CustomDictionary3, ref object CustomDictionary4,
            ref object CustomDictionary5, ref object CustomDictionary6, ref object CustomDictionary7,
            ref object CustomDictionary8, ref object CustomDictionary9, ref object CustomDictionary10)
        {
            throw new System.NotImplementedException();
        }

        public void FollowHyperlink(ref object Address, ref object SubAddress, ref object NewWindow, ref object AddHistory,
            ref object ExtraInfo, ref object Method, ref object HeaderInfo)
        {
            throw new System.NotImplementedException();
        }

        public void AddToFavorites()
        {
            throw new System.NotImplementedException();
        }

        public void Reload()
        {
            throw new System.NotImplementedException();
        }

        public Range AutoSummarize(ref object Length, ref object Mode, ref object UpdateProperties)
        {
            throw new System.NotImplementedException();
        }

        public void RemoveNumbers(ref object NumberType)
        {
            throw new System.NotImplementedException();
        }

        public void ConvertNumbersToText(ref object NumberType)
        {
            throw new System.NotImplementedException();
        }

        public int CountNumberedItems(ref object NumberType, ref object Level)
        {
            throw new System.NotImplementedException();
        }

        public void Post()
        {
            throw new System.NotImplementedException();
        }

        public void ToggleFormsDesign()
        {
            throw new System.NotImplementedException();
        }

        public void Compare2000(string Name)
        {
            throw new System.NotImplementedException();
        }

        public void UpdateSummaryProperties()
        {
            throw new System.NotImplementedException();
        }

        public object GetCrossReferenceItems(ref object ReferenceType)
        {
            throw new System.NotImplementedException();
        }

        public void AutoFormat()
        {
            throw new System.NotImplementedException();
        }

        public void ViewCode()
        {
            throw new System.NotImplementedException();
        }

        public void ViewPropertyBrowser()
        {
            throw new System.NotImplementedException();
        }

        public void ForwardMailer()
        {
            throw new System.NotImplementedException();
        }

        public void Reply()
        {
            throw new System.NotImplementedException();
        }

        public void ReplyAll()
        {
            throw new System.NotImplementedException();
        }

        public void SendMailer(ref object FileFormat, ref object Priority)
        {
            throw new System.NotImplementedException();
        }

        public void UndoClear()
        {
            throw new System.NotImplementedException();
        }

        public void PresentIt()
        {
            throw new System.NotImplementedException();
        }

        public void SendFax(string Address, ref object Subject)
        {
            throw new System.NotImplementedException();
        }

        public void Merge2000(string FileName)
        {
            throw new System.NotImplementedException();
        }

        public void ClosePrintPreview()
        {
            throw new System.NotImplementedException();
        }

        public void CheckConsistency()
        {
            throw new System.NotImplementedException();
        }

        public LetterContent CreateLetterContent(string DateFormat, bool IncludeHeaderFooter, string PageDesign,
            WdLetterStyle LetterStyle, bool Letterhead, WdLetterheadLocation LetterheadLocation, float LetterheadSize,
            string RecipientName, string RecipientAddress, string Salutation, WdSalutationType SalutationType,
            string RecipientReference, string MailingInstructions, string AttentionLine, string Subject, string CCList,
            string ReturnAddress, string SenderName, string Closing, string SenderCompany, string SenderJobTitle,
            string SenderInitials, int EnclosureNumber, ref object InfoBlock, ref object RecipientCode,
            ref object RecipientGender, ref object ReturnAddressShortForm, ref object SenderCity, ref object SenderCode,
            ref object SenderGender, ref object SenderReference)
        {
            throw new System.NotImplementedException();
        }

        public void AcceptAllRevisions()
        {
            throw new System.NotImplementedException();
        }

        public void RejectAllRevisions()
        {
            throw new System.NotImplementedException();
        }

        public void DetectLanguage()
        {
            throw new System.NotImplementedException();
        }

        public void ApplyTheme(string Name)
        {
            throw new System.NotImplementedException();
        }

        public void RemoveTheme()
        {
            throw new System.NotImplementedException();
        }

        public void WebPagePreview()
        {
            throw new System.NotImplementedException();
        }

        public void ReloadAs(MsoEncoding Encoding)
        {
            throw new System.NotImplementedException();
        }

        public void PrintOut2000(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object ActivePrinterMacGX, ref object ManualDuplexPrint, ref object PrintZoomColumn,
            ref object PrintZoomRow, ref object PrintZoomPaperWidth, ref object PrintZoomPaperHeight)
        {
            throw new System.NotImplementedException();
        }

        public void sblt(string s)
        {
            throw new System.NotImplementedException();
        }

        public void ConvertVietDoc(int CodePageOrigin)
        {
            throw new System.NotImplementedException();
        }

        public void PrintOut(ref object Background, ref object Append, ref object Range, ref object OutputFileName, ref object From,
            ref object To, ref object Item, ref object Copies, ref object Pages, ref object PageType, ref object PrintToFile,
            ref object Collate, ref object ActivePrinterMacGX, ref object ManualDuplexPrint, ref object PrintZoomColumn,
            ref object PrintZoomRow, ref object PrintZoomPaperWidth, ref object PrintZoomPaperHeight)
        {
            throw new System.NotImplementedException();
        }

        public void Compare2002(string Name, ref object AuthorName, ref object CompareTarget, ref object DetectFormatChanges,
            ref object IgnoreAllComparisonWarnings, ref object AddToRecentFiles)
        {
            throw new System.NotImplementedException();
        }

        public void CheckIn(bool SaveChanges, ref object Comments, bool MakePublic = false)
        {
            throw new System.NotImplementedException();
        }

        public bool CanCheckin()
        {
            throw new System.NotImplementedException();
        }

        public void Merge(string FileName, ref object MergeTarget, ref object DetectFormatChanges, ref object UseFormattingFrom,
            ref object AddToRecentFiles)
        {
            throw new System.NotImplementedException();
        }

        public void SendForReview(ref object Recipients, ref object Subject, ref object ShowMessage, ref object IncludeAttachment)
        {
            throw new System.NotImplementedException();
        }

        public void ReplyWithChanges(ref object ShowMessage)
        {
            throw new System.NotImplementedException();
        }

        public void EndReview()
        {
            throw new System.NotImplementedException();
        }

        public void SetPasswordEncryptionOptions(string PasswordEncryptionProvider, string PasswordEncryptionAlgorithm,
            int PasswordEncryptionKeyLength, ref object PasswordEncryptionFileProperties)
        {
            throw new System.NotImplementedException();
        }

        public void RecheckSmartTags()
        {
            throw new System.NotImplementedException();
        }

        public void RemoveSmartTags()
        {
            throw new System.NotImplementedException();
        }

        public void SetDefaultTableStyle(ref object Style, bool SetInTemplate)
        {
            throw new System.NotImplementedException();
        }

        public void DeleteAllComments()
        {
            throw new System.NotImplementedException();
        }

        public void AcceptAllRevisionsShown()
        {
            throw new System.NotImplementedException();
        }

        public void RejectAllRevisionsShown()
        {
            throw new System.NotImplementedException();
        }

        public void DeleteAllCommentsShown()
        {
            throw new System.NotImplementedException();
        }

        public void ResetFormFields()
        {
            throw new System.NotImplementedException();
        }

        public void SaveAs(ref object FileName, ref object FileFormat, ref object LockComments, ref object Password,
            ref object AddToRecentFiles, ref object WritePassword, ref object ReadOnlyRecommended, ref object EmbedTrueTypeFonts,
            ref object SaveNativePictureFormat, ref object SaveFormsData, ref object SaveAsAOCELetter, ref object Encoding,
            ref object InsertLineBreaks, ref object AllowSubstitutions, ref object LineEnding, ref object AddBiDiMarks)
        {
            throw new System.NotImplementedException();
        }

        public void CheckNewSmartTags()
        {
            throw new System.NotImplementedException();
        }

        public void SendFaxOverInternet(ref object Recipients, ref object Subject, ref object ShowMessage)
        {
            throw new System.NotImplementedException();
        }

        public void TransformDocument(string Path, bool DataOnly = true)
        {
            throw new System.NotImplementedException();
        }

        public void Protect(WdProtectionType Type, ref object NoReset, ref object Password, ref object UseIRM,
            ref object EnforceStyleLock)
        {
            throw new System.NotImplementedException();
        }

        public void SelectAllEditableRanges(ref object EditorID)
        {
            throw new System.NotImplementedException();
        }

        public void DeleteAllEditableRanges(ref object EditorID)
        {
            throw new System.NotImplementedException();
        }

        public void DeleteAllInkAnnotations()
        {
            throw new System.NotImplementedException();
        }

        public void AddDocumentWorkspaceHeader(bool RichFormat, string Url, string Title, string Description, string ID)
        {
            throw new System.NotImplementedException();
        }

        public void RemoveDocumentWorkspaceHeader(string ID)
        {
            throw new System.NotImplementedException();
        }

        public void Compare(string Name, ref object AuthorName, ref object CompareTarget, ref object DetectFormatChanges,
            ref object IgnoreAllComparisonWarnings, ref object AddToRecentFiles, ref object RemovePersonalInformation,
            ref object RemoveDateAndTime)
        {
            throw new System.NotImplementedException();
        }

        public void RemoveLockedStyles()
        {
            throw new System.NotImplementedException();
        }

        public XMLNode SelectSingleNode(string XPath, string PrefixMapping = "", bool FastSearchSkippingTextNodes = true)
        {
            throw new System.NotImplementedException();
        }

        public XMLNodes SelectNodes(string XPath, string PrefixMapping = "", bool FastSearchSkippingTextNodes = true)
        {
            throw new System.NotImplementedException();
        }

        public void Dummy1()
        {
            throw new System.NotImplementedException();
        }

        public void RemoveDocumentInformation(WdRemoveDocInfoType RemoveDocInfoType)
        {
            throw new System.NotImplementedException();
        }

        public void CheckInWithVersion(bool SaveChanges, ref object Comments, bool MakePublic, ref object VersionType)
        {
            throw new System.NotImplementedException();
        }

        public void Dummy2()
        {
            throw new System.NotImplementedException();
        }

        public void Dummy3()
        {
            throw new System.NotImplementedException();
        }

        public void LockServerFile()
        {
            throw new System.NotImplementedException();
        }

        public WorkflowTasks GetWorkflowTasks()
        {
            throw new System.NotImplementedException();
        }

        public WorkflowTemplates GetWorkflowTemplates()
        {
            throw new System.NotImplementedException();
        }

        public void Dummy4()
        {
            throw new System.NotImplementedException();
        }

        public void AddMeetingWorkspaceHeader(bool SkipIfAbsent, string Url, string Title, string Description, string ID)
        {
            throw new System.NotImplementedException();
        }

        public void SaveAsQuickStyleSet(string FileName)
        {
            throw new System.NotImplementedException();
        }

        public void ApplyQuickStyleSet(string Name)
        {
            throw new System.NotImplementedException();
        }

        public void ApplyDocumentTheme(string FileName)
        {
            throw new System.NotImplementedException();
        }

        public ContentControls SelectLinkedControls(CustomXMLNode Node)
        {
            throw new System.NotImplementedException();
        }

        public ContentControls SelectUnlinkedControls(CustomXMLPart Stream = null)
        {
            throw new System.NotImplementedException();
        }

        public ContentControls SelectContentControlsByTitle(string Title)
        {
            throw new System.NotImplementedException();
        }

        public void ExportAsFixedFormat(string OutputFileName, WdExportFormat ExportFormat, bool OpenAfterExport,
            WdExportOptimizeFor OptimizeFor, WdExportRange Range, int From,
            int To, WdExportItem Item, bool IncludeDocProps, bool KeepIRM,
            WdExportCreateBookmarks CreateBookmarks, bool DocStructureTags,
            bool BitmapMissingFonts, bool UseISO19005_1, ref object FixedFormatExtClassPtr)
        {
            throw new System.NotImplementedException();
        }

        public void FreezeLayout()
        {
            throw new System.NotImplementedException();
        }

        public void UnfreezeLayout()
        {
            throw new System.NotImplementedException();
        }

        public void DowngradeDocument()
        {
            throw new System.NotImplementedException();
        }

        public void Convert()
        {
            throw new System.NotImplementedException();
        }

        public ContentControls SelectContentControlsByTag(string Tag)
        {
            throw new System.NotImplementedException();
        }

        public void ConvertAutoHyphens()
        {
            throw new System.NotImplementedException();
        }

        public void ApplyQuickStyleSet2(ref object Style)
        {
            throw new System.NotImplementedException();
        }

        public void SaveAs2(ref object FileName, ref object FileFormat, ref object LockComments, ref object Password,
            ref object AddToRecentFiles, ref object WritePassword, ref object ReadOnlyRecommended, ref object EmbedTrueTypeFonts,
            ref object SaveNativePictureFormat, ref object SaveFormsData, ref object SaveAsAOCELetter, ref object Encoding,
            ref object InsertLineBreaks, ref object AllowSubstitutions, ref object LineEnding, ref object AddBiDiMarks,
            ref object CompatibilityMode)
        {
            throw new System.NotImplementedException();
        }

        public void SetCompatibilityMode(int Mode)
        {
            throw new System.NotImplementedException();
        }

        public int ReturnToLastReadPosition()
        {
            throw new System.NotImplementedException();
        }

        public void SaveCopyAs(ref object FileName, ref object FileFormat, ref object LockComments, ref object Password,
            ref object AddToRecentFiles, ref object WritePassword, ref object ReadOnlyRecommended, ref object EmbedTrueTypeFonts,
            ref object SaveNativePictureFormat, ref object SaveFormsData, ref object SaveAsAOCELetter, ref object Encoding,
            ref object InsertLineBreaks, ref object AllowSubstitutions, ref object LineEnding, ref object AddBiDiMarks,
            ref object CompatibilityMode)
        {
            throw new System.NotImplementedException();
        }

        public string Name { get; private set; }
        public Application Application { get; private set; }
        public int Creator { get; private set; }
        public object Parent { get; private set; }
        public object BuiltInDocumentProperties { get; private set; }
        public object CustomDocumentProperties { get; private set; }
        public string Path { get; private set; }
        public Bookmarks Bookmarks { get; private set; }
        public Tables Tables { get; private set; }
        public Footnotes Footnotes { get; private set; }
        public Endnotes Endnotes { get; private set; }
        public Comments Comments { get; private set; }
        public WdDocumentType Type { get; private set; }
        public bool AutoHyphenation { get; set; }
        public bool HyphenateCaps { get; set; }
        public int HyphenationZone { get; set; }
        public int ConsecutiveHyphensLimit { get; set; }
        public Sections Sections { get; private set; }
        public Paragraphs Paragraphs { get; private set; }
        public Words Words { get; private set; }
        public Sentences Sentences { get; private set; }
        public Characters Characters { get; private set; }
        public Fields Fields { get; private set; }
        public FormFields FormFields { get; private set; }
        public Styles Styles { get; private set; }
        public Frames Frames { get; private set; }
        public TablesOfFigures TablesOfFigures { get; private set; }
        public Variables Variables { get; private set; }
        public MailMerge MailMerge { get; private set; }
        public Envelope Envelope { get; private set; }
        public string FullName { get; private set; }
        public Revisions Revisions { get; private set; }
        public TablesOfContents TablesOfContents { get; private set; }
        public TablesOfAuthorities TablesOfAuthorities { get; private set; }
        public PageSetup PageSetup { get; set; }
        public Windows Windows { get; private set; }
        public bool HasRoutingSlip { get; set; }
        public RoutingSlip RoutingSlip { get; private set; }
        public bool Routed { get; private set; }
        public TablesOfAuthoritiesCategories TablesOfAuthoritiesCategories { get; private set; }
        public Indexes Indexes { get; private set; }
        public bool Saved { get; set; }
        public Range Content { get; private set; }
        public Window ActiveWindow { get; private set; }
        public WdDocumentKind Kind { get; set; }
        public bool ReadOnly { get; private set; }
        public Subdocuments Subdocuments { get; private set; }
        public bool IsMasterDocument { get; private set; }
        public float DefaultTabStop { get; set; }
        public bool EmbedTrueTypeFonts { get; set; }
        public bool SaveFormsData { get; set; }
        public bool ReadOnlyRecommended { get; set; }
        public bool SaveSubsetFonts { get; set; }
        public bool get_Compatibility(WdCompatibility Type)
        {
            throw new System.NotImplementedException();
        }

        public void set_Compatibility(WdCompatibility Type, bool prop)
        {
            throw new System.NotImplementedException();
        }

        public StoryRanges StoryRanges { get; private set; }
        public CommandBars CommandBars { get; private set; }
        public bool IsSubdocument { get; private set; }
        public int SaveFormat { get; private set; }
        public WdProtectionType ProtectionType { get; private set; }
        public Hyperlinks Hyperlinks { get; private set; }
        public Shapes Shapes { get; private set; }
        public ListTemplates ListTemplates { get; private set; }
        public Lists Lists { get; private set; }
        public bool UpdateStylesOnOpen { get; set; }
        public object get_AttachedTemplate()
        {
            throw new System.NotImplementedException();
        }

        public void set_AttachedTemplate(ref object prop)
        {
            throw new System.NotImplementedException();
        }

        public InlineShapes InlineShapes { get; private set; }
        public Shape Background { get; set; }
        public bool GrammarChecked { get; set; }
        public bool SpellingChecked { get; set; }
        public bool ShowGrammaticalErrors { get; set; }
        public bool ShowSpellingErrors { get; set; }
        public Versions Versions { get; private set; }
        public bool ShowSummary { get; set; }
        public WdSummaryMode SummaryViewMode { get; set; }
        public int SummaryLength { get; set; }
        public bool PrintFractionalWidths { get; set; }
        public bool PrintPostScriptOverText { get; set; }
        public object Container { get; private set; }
        public bool PrintFormsData { get; set; }
        public ListParagraphs ListParagraphs { get; private set; }
        public string Password { set; private get; }
        public string WritePassword { set; private get; }
        public bool HasPassword { get; private set; }
        public bool WriteReserved { get; private set; }
        public string get_ActiveWritingStyle(ref object LanguageID)
        {
            throw new System.NotImplementedException();
        }

        public void set_ActiveWritingStyle(ref object LanguageID, string prop)
        {
            throw new System.NotImplementedException();
        }

        public bool UserControl { get; set; }
        public bool HasMailer { get; set; }
        public Mailer Mailer { get; private set; }
        public ReadabilityStatistics ReadabilityStatistics { get; private set; }
        public ProofreadingErrors GrammaticalErrors { get; private set; }
        public ProofreadingErrors SpellingErrors { get; private set; }
        public VBProject VBProject { get; private set; }
        public bool FormsDesign { get; private set; }
        public string _CodeName { get; set; }
        public string CodeName { get; private set; }
        public bool SnapToGrid { get; set; }
        public bool SnapToShapes { get; set; }
        public float GridDistanceHorizontal { get; set; }
        public float GridDistanceVertical { get; set; }
        public float GridOriginHorizontal { get; set; }
        public float GridOriginVertical { get; set; }
        public int GridSpaceBetweenHorizontalLines { get; set; }
        public int GridSpaceBetweenVerticalLines { get; set; }
        public bool GridOriginFromMargin { get; set; }
        public bool KerningByAlgorithm { get; set; }
        public WdJustificationMode JustificationMode { get; set; }
        public WdFarEastLineBreakLevel FarEastLineBreakLevel { get; set; }
        public string NoLineBreakBefore { get; set; }
        public string NoLineBreakAfter { get; set; }
        public bool TrackRevisions { get; set; }
        public bool PrintRevisions { get; set; }
        public bool ShowRevisions { get; set; }
        public string ActiveTheme { get; private set; }
        public string ActiveThemeDisplayName { get; private set; }
        public Email Email { get; private set; }
        public Scripts Scripts { get; private set; }
        public bool LanguageDetected { get; set; }
        public WdFarEastLineBreakLanguageID FarEastLineBreakLanguage { get; set; }
        public Frameset Frameset { get; private set; }
        public object get_ClickAndTypeParagraphStyle()
        {
            throw new System.NotImplementedException();
        }

        public void set_ClickAndTypeParagraphStyle(ref object prop)
        {
            throw new System.NotImplementedException();
        }

        public HTMLProject HTMLProject { get; private set; }
        public WebOptions WebOptions { get; private set; }
        public MsoEncoding OpenEncoding { get; private set; }
        public MsoEncoding SaveEncoding { get; set; }
        public bool OptimizeForWord97 { get; set; }
        public bool VBASigned { get; private set; }
        public MsoEnvelope MailEnvelope { get; private set; }
        public bool DisableFeatures { get; set; }
        public bool DoNotEmbedSystemFonts { get; set; }
        public SignatureSet Signatures { get; private set; }
        public string DefaultTargetFrame { get; set; }
        public HTMLDivisions HTMLDivisions { get; private set; }
        public WdDisableFeaturesIntroducedAfter DisableFeaturesIntroducedAfter { get; set; }
        public bool RemovePersonalInformation { get; set; }
        public SmartTags SmartTags { get; private set; }
        public bool EmbedSmartTags { get; set; }
        public bool SmartTagsAsXMLProps { get; set; }
        public MsoEncoding TextEncoding { get; set; }
        public WdLineEndingType TextLineEnding { get; set; }
        public StyleSheets StyleSheets { get; private set; }
        public object DefaultTableStyle { get; private set; }
        public string PasswordEncryptionProvider { get; private set; }
        public string PasswordEncryptionAlgorithm { get; private set; }
        public int PasswordEncryptionKeyLength { get; private set; }
        public bool PasswordEncryptionFileProperties { get; private set; }
        public bool EmbedLinguisticData { get; set; }
        public bool FormattingShowFont { get; set; }
        public bool FormattingShowClear { get; set; }
        public bool FormattingShowParagraph { get; set; }
        public bool FormattingShowNumbering { get; set; }
        public WdShowFilter FormattingShowFilter { get; set; }
        public Permission Permission { get; private set; }
        public XMLNodes XMLNodes { get; private set; }
        public XMLSchemaReferences XMLSchemaReferences { get; private set; }
        public SmartDocument SmartDocument { get; private set; }
        public SharedWorkspace SharedWorkspace { get; private set; }

        Sync _Document.Sync
        {
            get { throw new System.NotImplementedException(); }
        }

        public bool EnforceStyle { get; set; }
        public bool AutoFormatOverride { get; set; }
        public bool XMLSaveDataOnly { get; set; }
        public bool XMLHideNamespaces { get; set; }
        public bool XMLShowAdvancedErrors { get; set; }
        public bool XMLUseXSLTWhenSaving { get; set; }
        public string XMLSaveThroughXSLT { get; set; }
        public DocumentLibraryVersions DocumentLibraryVersions { get; private set; }
        public bool ReadingModeLayoutFrozen { get; set; }
        public bool RemoveDateAndTime { get; set; }
        public XMLChildNodeSuggestions ChildNodeSuggestions { get; private set; }
        public XMLNodes XMLSchemaViolations { get; private set; }
        public int ReadingLayoutSizeX { get; set; }
        public int ReadingLayoutSizeY { get; set; }
        public WdStyleSort StyleSortMethod { get; set; }
        public MetaProperties ContentTypeProperties { get; private set; }
        public bool TrackMoves { get; set; }
        public bool TrackFormatting { get; set; }
        public OMaths OMaths { get; private set; }
        public ServerPolicy ServerPolicy { get; private set; }
        public ContentControls ContentControls { get; private set; }
        public DocumentInspectors DocumentInspectors { get; private set; }
        public Bibliography Bibliography { get; private set; }
        public bool LockTheme { get; set; }
        public bool LockQuickStyleSet { get; set; }
        public string OriginalDocumentTitle { get; private set; }
        public string RevisedDocumentTitle { get; private set; }
        public CustomXMLParts CustomXMLParts { get; private set; }
        public bool FormattingShowNextLevel { get; set; }
        public bool FormattingShowUserStyleName { get; set; }
        public Research Research { get; private set; }
        public bool Final { get; set; }
        public WdOMathBreakBin OMathBreakBin { get; set; }
        public WdOMathBreakSub OMathBreakSub { get; set; }
        public WdOMathJc OMathJc { get; set; }
        public float OMathLeftMargin { get; set; }
        public float OMathRightMargin { get; set; }
        public float OMathWrap { get; set; }
        public bool OMathIntSubSupLim { get; set; }
        public bool OMathNarySupSubLim { get; set; }
        public bool OMathSmallFrac { get; set; }
        public string WordOpenXML { get; private set; }
        public OfficeTheme DocumentTheme { get; private set; }
        public bool HasVBProject { get; private set; }
        public string OMathFontName { get; set; }
        public string EncryptionProvider { get; set; }
        public bool UseMathDefaults { get; set; }
        public int CurrentRsid { get; private set; }
        public int DocID { get; private set; }
        public int CompatibilityMode { get; private set; }
        public CoAuthoring CoAuthoring { get; private set; }
        public Broadcast Broadcast { get; private set; }
        public bool ChartDataPointTrack { get; set; }
        public bool IsInAutosave { get; private set; }
        public event DocumentEvents2_NewEventHandler New;
        public event DocumentEvents2_OpenEventHandler Open;
        public event DocumentEvents2_CloseEventHandler Close;
        event DocumentEvents2_SyncEventHandler DocumentEvents2_Event.Sync
        {
            add { throw new System.NotImplementedException(); }
            remove { throw new System.NotImplementedException(); }
        }

        public event DocumentEvents2_XMLAfterInsertEventHandler XMLAfterInsert;
        public event DocumentEvents2_XMLBeforeDeleteEventHandler XMLBeforeDelete;
        public event DocumentEvents2_ContentControlAfterAddEventHandler ContentControlAfterAdd;
        public event DocumentEvents2_ContentControlBeforeDeleteEventHandler ContentControlBeforeDelete;
        public event DocumentEvents2_ContentControlOnExitEventHandler ContentControlOnExit;
        public event DocumentEvents2_ContentControlOnEnterEventHandler ContentControlOnEnter;
        public event DocumentEvents2_ContentControlBeforeStoreUpdateEventHandler ContentControlBeforeStoreUpdate;
        public event DocumentEvents2_ContentControlBeforeContentUpdateEventHandler ContentControlBeforeContentUpdate;
        public event DocumentEvents2_BuildingBlockInsertEventHandler BuildingBlockInsert;
    }
}