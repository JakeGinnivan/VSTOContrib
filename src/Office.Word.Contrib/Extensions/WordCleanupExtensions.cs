using Office.Contrib.Extensions;

namespace Office.Word.Contrib.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Application WithComCleanup(this Microsoft.Office.Interop.Word._Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Application, Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Global WithComCleanup(this Microsoft.Office.Interop.Word._Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Global, Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for FontNames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFontNames WithComCleanup(this Microsoft.Office.Interop.Word.FontNames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FontNames, Interfaces.IFontNames>();
		}

		/// <summary>
		/// Wrapper interface for Languages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILanguages WithComCleanup(this Microsoft.Office.Interop.Word.Languages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Languages, Interfaces.ILanguages>();
		}

		/// <summary>
		/// Wrapper interface for Language which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILanguage WithComCleanup(this Microsoft.Office.Interop.Word.Language resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Language, Interfaces.ILanguage>();
		}

		/// <summary>
		/// Wrapper interface for Documents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocuments WithComCleanup(this Microsoft.Office.Interop.Word.Documents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Documents, Interfaces.IDocuments>();
		}

		/// <summary>
		/// Wrapper interface for _Document which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Document WithComCleanup(this Microsoft.Office.Interop.Word._Document resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Document, Interfaces.I_Document>();
		}

		/// <summary>
		/// Wrapper interface for Template which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITemplate WithComCleanup(this Microsoft.Office.Interop.Word.Template resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Template, Interfaces.ITemplate>();
		}

		/// <summary>
		/// Wrapper interface for Templates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITemplates WithComCleanup(this Microsoft.Office.Interop.Word.Templates resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Templates, Interfaces.ITemplates>();
		}

		/// <summary>
		/// Wrapper interface for RoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRoutingSlip WithComCleanup(this Microsoft.Office.Interop.Word.RoutingSlip resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.RoutingSlip, Interfaces.IRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for Bookmark which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBookmark WithComCleanup(this Microsoft.Office.Interop.Word.Bookmark resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Bookmark, Interfaces.IBookmark>();
		}

		/// <summary>
		/// Wrapper interface for Bookmarks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBookmarks WithComCleanup(this Microsoft.Office.Interop.Word.Bookmarks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Bookmarks, Interfaces.IBookmarks>();
		}

		/// <summary>
		/// Wrapper interface for Variable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVariable WithComCleanup(this Microsoft.Office.Interop.Word.Variable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Variable, Interfaces.IVariable>();
		}

		/// <summary>
		/// Wrapper interface for Variables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVariables WithComCleanup(this Microsoft.Office.Interop.Word.Variables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Variables, Interfaces.IVariables>();
		}

		/// <summary>
		/// Wrapper interface for RecentFile which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecentFile WithComCleanup(this Microsoft.Office.Interop.Word.RecentFile resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.RecentFile, Interfaces.IRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for RecentFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecentFiles WithComCleanup(this Microsoft.Office.Interop.Word.RecentFiles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.RecentFiles, Interfaces.IRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for Window which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWindow WithComCleanup(this Microsoft.Office.Interop.Word.Window resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Window, Interfaces.IWindow>();
		}

		/// <summary>
		/// Wrapper interface for Windows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWindows WithComCleanup(this Microsoft.Office.Interop.Word.Windows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Windows, Interfaces.IWindows>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPane WithComCleanup(this Microsoft.Office.Interop.Word.Pane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Pane, Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPanes WithComCleanup(this Microsoft.Office.Interop.Word.Panes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Panes, Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Range which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRange WithComCleanup(this Microsoft.Office.Interop.Word.Range resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Range, Interfaces.IRange>();
		}

		/// <summary>
		/// Wrapper interface for ListFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListFormat WithComCleanup(this Microsoft.Office.Interop.Word.ListFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListFormat, Interfaces.IListFormat>();
		}

		/// <summary>
		/// Wrapper interface for Find which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFind WithComCleanup(this Microsoft.Office.Interop.Word.Find resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Find, Interfaces.IFind>();
		}

		/// <summary>
		/// Wrapper interface for Replacement which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReplacement WithComCleanup(this Microsoft.Office.Interop.Word.Replacement resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Replacement, Interfaces.IReplacement>();
		}

		/// <summary>
		/// Wrapper interface for Characters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICharacters WithComCleanup(this Microsoft.Office.Interop.Word.Characters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Characters, Interfaces.ICharacters>();
		}

		/// <summary>
		/// Wrapper interface for Words which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWords WithComCleanup(this Microsoft.Office.Interop.Word.Words resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Words, Interfaces.IWords>();
		}

		/// <summary>
		/// Wrapper interface for Sentences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISentences WithComCleanup(this Microsoft.Office.Interop.Word.Sentences resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Sentences, Interfaces.ISentences>();
		}

		/// <summary>
		/// Wrapper interface for Sections which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISections WithComCleanup(this Microsoft.Office.Interop.Word.Sections resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Sections, Interfaces.ISections>();
		}

		/// <summary>
		/// Wrapper interface for Section which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISection WithComCleanup(this Microsoft.Office.Interop.Word.Section resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Section, Interfaces.ISection>();
		}

		/// <summary>
		/// Wrapper interface for Paragraphs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphs WithComCleanup(this Microsoft.Office.Interop.Word.Paragraphs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Paragraphs, Interfaces.IParagraphs>();
		}

		/// <summary>
		/// Wrapper interface for Paragraph which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraph WithComCleanup(this Microsoft.Office.Interop.Word.Paragraph resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Paragraph, Interfaces.IParagraph>();
		}

		/// <summary>
		/// Wrapper interface for DropCap which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropCap WithComCleanup(this Microsoft.Office.Interop.Word.DropCap resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DropCap, Interfaces.IDropCap>();
		}

		/// <summary>
		/// Wrapper interface for TabStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStops WithComCleanup(this Microsoft.Office.Interop.Word.TabStops resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TabStops, Interfaces.ITabStops>();
		}

		/// <summary>
		/// Wrapper interface for TabStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStop WithComCleanup(this Microsoft.Office.Interop.Word.TabStop resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TabStop, Interfaces.ITabStop>();
		}

		/// <summary>
		/// Wrapper interface for _ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ParagraphFormat WithComCleanup(this Microsoft.Office.Interop.Word._ParagraphFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._ParagraphFormat, Interfaces.I_ParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for _Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Font WithComCleanup(this Microsoft.Office.Interop.Word._Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Font, Interfaces.I_Font>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITable WithComCleanup(this Microsoft.Office.Interop.Word.Table resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Table, Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRow WithComCleanup(this Microsoft.Office.Interop.Word.Row resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Row, Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumn WithComCleanup(this Microsoft.Office.Interop.Word.Column resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Column, Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for Cell which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICell WithComCleanup(this Microsoft.Office.Interop.Word.Cell resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Cell, Interfaces.ICell>();
		}

		/// <summary>
		/// Wrapper interface for Tables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITables WithComCleanup(this Microsoft.Office.Interop.Word.Tables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Tables, Interfaces.ITables>();
		}

		/// <summary>
		/// Wrapper interface for Rows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRows WithComCleanup(this Microsoft.Office.Interop.Word.Rows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Rows, Interfaces.IRows>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumns WithComCleanup(this Microsoft.Office.Interop.Word.Columns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Columns, Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for Cells which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICells WithComCleanup(this Microsoft.Office.Interop.Word.Cells resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Cells, Interfaces.ICells>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrect WithComCleanup(this Microsoft.Office.Interop.Word.AutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCorrect, Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrectEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrectEntries WithComCleanup(this Microsoft.Office.Interop.Word.AutoCorrectEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCorrectEntries, Interfaces.IAutoCorrectEntries>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrectEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrectEntry WithComCleanup(this Microsoft.Office.Interop.Word.AutoCorrectEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCorrectEntry, Interfaces.IAutoCorrectEntry>();
		}

		/// <summary>
		/// Wrapper interface for FirstLetterExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFirstLetterExceptions WithComCleanup(this Microsoft.Office.Interop.Word.FirstLetterExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FirstLetterExceptions, Interfaces.IFirstLetterExceptions>();
		}

		/// <summary>
		/// Wrapper interface for FirstLetterException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFirstLetterException WithComCleanup(this Microsoft.Office.Interop.Word.FirstLetterException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FirstLetterException, Interfaces.IFirstLetterException>();
		}

		/// <summary>
		/// Wrapper interface for TwoInitialCapsExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITwoInitialCapsExceptions WithComCleanup(this Microsoft.Office.Interop.Word.TwoInitialCapsExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TwoInitialCapsExceptions, Interfaces.ITwoInitialCapsExceptions>();
		}

		/// <summary>
		/// Wrapper interface for TwoInitialCapsException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITwoInitialCapsException WithComCleanup(this Microsoft.Office.Interop.Word.TwoInitialCapsException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TwoInitialCapsException, Interfaces.ITwoInitialCapsException>();
		}

		/// <summary>
		/// Wrapper interface for Footnotes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFootnotes WithComCleanup(this Microsoft.Office.Interop.Word.Footnotes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Footnotes, Interfaces.IFootnotes>();
		}

		/// <summary>
		/// Wrapper interface for Endnotes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEndnotes WithComCleanup(this Microsoft.Office.Interop.Word.Endnotes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Endnotes, Interfaces.IEndnotes>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComments WithComCleanup(this Microsoft.Office.Interop.Word.Comments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Comments, Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Footnote which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFootnote WithComCleanup(this Microsoft.Office.Interop.Word.Footnote resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Footnote, Interfaces.IFootnote>();
		}

		/// <summary>
		/// Wrapper interface for Endnote which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEndnote WithComCleanup(this Microsoft.Office.Interop.Word.Endnote resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Endnote, Interfaces.IEndnote>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComment WithComCleanup(this Microsoft.Office.Interop.Word.Comment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Comment, Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorders WithComCleanup(this Microsoft.Office.Interop.Word.Borders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Borders, Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for Border which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorder WithComCleanup(this Microsoft.Office.Interop.Word.Border resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Border, Interfaces.IBorder>();
		}

		/// <summary>
		/// Wrapper interface for Shading which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShading WithComCleanup(this Microsoft.Office.Interop.Word.Shading resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Shading, Interfaces.IShading>();
		}

		/// <summary>
		/// Wrapper interface for TextRetrievalMode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextRetrievalMode WithComCleanup(this Microsoft.Office.Interop.Word.TextRetrievalMode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextRetrievalMode, Interfaces.ITextRetrievalMode>();
		}

		/// <summary>
		/// Wrapper interface for AutoTextEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoTextEntries WithComCleanup(this Microsoft.Office.Interop.Word.AutoTextEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoTextEntries, Interfaces.IAutoTextEntries>();
		}

		/// <summary>
		/// Wrapper interface for AutoTextEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoTextEntry WithComCleanup(this Microsoft.Office.Interop.Word.AutoTextEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoTextEntry, Interfaces.IAutoTextEntry>();
		}

		/// <summary>
		/// Wrapper interface for System which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISystem WithComCleanup(this Microsoft.Office.Interop.Word.System resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.System, Interfaces.ISystem>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEFormat WithComCleanup(this Microsoft.Office.Interop.Word.OLEFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OLEFormat, Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinkFormat WithComCleanup(this Microsoft.Office.Interop.Word.LinkFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LinkFormat, Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for _OLEControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OLEControl WithComCleanup(this Microsoft.Office.Interop.Word._OLEControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._OLEControl, Interfaces.I_OLEControl>();
		}

		/// <summary>
		/// Wrapper interface for Fields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFields WithComCleanup(this Microsoft.Office.Interop.Word.Fields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Fields, Interfaces.IFields>();
		}

		/// <summary>
		/// Wrapper interface for Field which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IField WithComCleanup(this Microsoft.Office.Interop.Word.Field resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Field, Interfaces.IField>();
		}

		/// <summary>
		/// Wrapper interface for Browser which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBrowser WithComCleanup(this Microsoft.Office.Interop.Word.Browser resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Browser, Interfaces.IBrowser>();
		}

		/// <summary>
		/// Wrapper interface for Styles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyles WithComCleanup(this Microsoft.Office.Interop.Word.Styles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Styles, Interfaces.IStyles>();
		}

		/// <summary>
		/// Wrapper interface for Style which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyle WithComCleanup(this Microsoft.Office.Interop.Word.Style resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Style, Interfaces.IStyle>();
		}

		/// <summary>
		/// Wrapper interface for Frames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFrames WithComCleanup(this Microsoft.Office.Interop.Word.Frames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Frames, Interfaces.IFrames>();
		}

		/// <summary>
		/// Wrapper interface for Frame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFrame WithComCleanup(this Microsoft.Office.Interop.Word.Frame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Frame, Interfaces.IFrame>();
		}

		/// <summary>
		/// Wrapper interface for FormFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormFields WithComCleanup(this Microsoft.Office.Interop.Word.FormFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FormFields, Interfaces.IFormFields>();
		}

		/// <summary>
		/// Wrapper interface for FormField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormField WithComCleanup(this Microsoft.Office.Interop.Word.FormField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FormField, Interfaces.IFormField>();
		}

		/// <summary>
		/// Wrapper interface for TextInput which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextInput WithComCleanup(this Microsoft.Office.Interop.Word.TextInput resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextInput, Interfaces.ITextInput>();
		}

		/// <summary>
		/// Wrapper interface for CheckBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICheckBox WithComCleanup(this Microsoft.Office.Interop.Word.CheckBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CheckBox, Interfaces.ICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for DropDown which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropDown WithComCleanup(this Microsoft.Office.Interop.Word.DropDown resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DropDown, Interfaces.IDropDown>();
		}

		/// <summary>
		/// Wrapper interface for ListEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListEntries WithComCleanup(this Microsoft.Office.Interop.Word.ListEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListEntries, Interfaces.IListEntries>();
		}

		/// <summary>
		/// Wrapper interface for ListEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListEntry WithComCleanup(this Microsoft.Office.Interop.Word.ListEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListEntry, Interfaces.IListEntry>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfFigures which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfFigures WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfFigures resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfFigures, Interfaces.ITablesOfFigures>();
		}

		/// <summary>
		/// Wrapper interface for TableOfFigures which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfFigures WithComCleanup(this Microsoft.Office.Interop.Word.TableOfFigures resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfFigures, Interfaces.ITableOfFigures>();
		}

		/// <summary>
		/// Wrapper interface for MailMerge which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMerge WithComCleanup(this Microsoft.Office.Interop.Word.MailMerge resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMerge, Interfaces.IMailMerge>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeFields WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeFields, Interfaces.IMailMergeFields>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeField WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeField, Interfaces.IMailMergeField>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataSource which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeDataSource WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeDataSource resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeDataSource, Interfaces.IMailMergeDataSource>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFieldNames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeFieldNames WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeFieldNames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeFieldNames, Interfaces.IMailMergeFieldNames>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFieldName which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeFieldName WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeFieldName resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeFieldName, Interfaces.IMailMergeFieldName>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeDataFields WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeDataFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeDataFields, Interfaces.IMailMergeDataFields>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeDataField WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeDataField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeDataField, Interfaces.IMailMergeDataField>();
		}

		/// <summary>
		/// Wrapper interface for Envelope which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEnvelope WithComCleanup(this Microsoft.Office.Interop.Word.Envelope resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Envelope, Interfaces.IEnvelope>();
		}

		/// <summary>
		/// Wrapper interface for MailingLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailingLabel WithComCleanup(this Microsoft.Office.Interop.Word.MailingLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailingLabel, Interfaces.IMailingLabel>();
		}

		/// <summary>
		/// Wrapper interface for CustomLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLabels WithComCleanup(this Microsoft.Office.Interop.Word.CustomLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomLabels, Interfaces.ICustomLabels>();
		}

		/// <summary>
		/// Wrapper interface for CustomLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLabel WithComCleanup(this Microsoft.Office.Interop.Word.CustomLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomLabel, Interfaces.ICustomLabel>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfContents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfContents WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfContents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfContents, Interfaces.ITablesOfContents>();
		}

		/// <summary>
		/// Wrapper interface for TableOfContents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfContents WithComCleanup(this Microsoft.Office.Interop.Word.TableOfContents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfContents, Interfaces.ITableOfContents>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfAuthorities which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfAuthorities WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfAuthorities resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfAuthorities, Interfaces.ITablesOfAuthorities>();
		}

		/// <summary>
		/// Wrapper interface for TableOfAuthorities which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfAuthorities WithComCleanup(this Microsoft.Office.Interop.Word.TableOfAuthorities resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfAuthorities, Interfaces.ITableOfAuthorities>();
		}

		/// <summary>
		/// Wrapper interface for Dialogs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogs WithComCleanup(this Microsoft.Office.Interop.Word.Dialogs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dialogs, Interfaces.IDialogs>();
		}

		/// <summary>
		/// Wrapper interface for Dialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialog WithComCleanup(this Microsoft.Office.Interop.Word.Dialog resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dialog, Interfaces.IDialog>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageSetup WithComCleanup(this Microsoft.Office.Interop.Word.PageSetup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PageSetup, Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for LineNumbering which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineNumbering WithComCleanup(this Microsoft.Office.Interop.Word.LineNumbering resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LineNumbering, Interfaces.ILineNumbering>();
		}

		/// <summary>
		/// Wrapper interface for TextColumns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextColumns WithComCleanup(this Microsoft.Office.Interop.Word.TextColumns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextColumns, Interfaces.ITextColumns>();
		}

		/// <summary>
		/// Wrapper interface for TextColumn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextColumn WithComCleanup(this Microsoft.Office.Interop.Word.TextColumn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextColumn, Interfaces.ITextColumn>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISelection WithComCleanup(this Microsoft.Office.Interop.Word.Selection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Selection, Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfAuthoritiesCategories which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfAuthoritiesCategories WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories, Interfaces.ITablesOfAuthoritiesCategories>();
		}

		/// <summary>
		/// Wrapper interface for TableOfAuthoritiesCategory which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfAuthoritiesCategory WithComCleanup(this Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory, Interfaces.ITableOfAuthoritiesCategory>();
		}

		/// <summary>
		/// Wrapper interface for CaptionLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICaptionLabels WithComCleanup(this Microsoft.Office.Interop.Word.CaptionLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CaptionLabels, Interfaces.ICaptionLabels>();
		}

		/// <summary>
		/// Wrapper interface for CaptionLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICaptionLabel WithComCleanup(this Microsoft.Office.Interop.Word.CaptionLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CaptionLabel, Interfaces.ICaptionLabel>();
		}

		/// <summary>
		/// Wrapper interface for AutoCaptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCaptions WithComCleanup(this Microsoft.Office.Interop.Word.AutoCaptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCaptions, Interfaces.IAutoCaptions>();
		}

		/// <summary>
		/// Wrapper interface for AutoCaption which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCaption WithComCleanup(this Microsoft.Office.Interop.Word.AutoCaption resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCaption, Interfaces.IAutoCaption>();
		}

		/// <summary>
		/// Wrapper interface for Indexes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIndexes WithComCleanup(this Microsoft.Office.Interop.Word.Indexes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Indexes, Interfaces.IIndexes>();
		}

		/// <summary>
		/// Wrapper interface for Index which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIndex WithComCleanup(this Microsoft.Office.Interop.Word.Index resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Index, Interfaces.IIndex>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIn WithComCleanup(this Microsoft.Office.Interop.Word.AddIn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AddIn, Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIns WithComCleanup(this Microsoft.Office.Interop.Word.AddIns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AddIns, Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for Revisions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRevisions WithComCleanup(this Microsoft.Office.Interop.Word.Revisions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Revisions, Interfaces.IRevisions>();
		}

		/// <summary>
		/// Wrapper interface for Revision which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRevision WithComCleanup(this Microsoft.Office.Interop.Word.Revision resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Revision, Interfaces.IRevision>();
		}

		/// <summary>
		/// Wrapper interface for Task which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITask WithComCleanup(this Microsoft.Office.Interop.Word.Task resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Task, Interfaces.ITask>();
		}

		/// <summary>
		/// Wrapper interface for Tasks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITasks WithComCleanup(this Microsoft.Office.Interop.Word.Tasks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Tasks, Interfaces.ITasks>();
		}

		/// <summary>
		/// Wrapper interface for HeadersFooters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadersFooters WithComCleanup(this Microsoft.Office.Interop.Word.HeadersFooters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeadersFooters, Interfaces.IHeadersFooters>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeaderFooter WithComCleanup(this Microsoft.Office.Interop.Word.HeaderFooter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeaderFooter, Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for PageNumbers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageNumbers WithComCleanup(this Microsoft.Office.Interop.Word.PageNumbers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PageNumbers, Interfaces.IPageNumbers>();
		}

		/// <summary>
		/// Wrapper interface for PageNumber which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageNumber WithComCleanup(this Microsoft.Office.Interop.Word.PageNumber resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PageNumber, Interfaces.IPageNumber>();
		}

		/// <summary>
		/// Wrapper interface for Subdocuments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISubdocuments WithComCleanup(this Microsoft.Office.Interop.Word.Subdocuments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Subdocuments, Interfaces.ISubdocuments>();
		}

		/// <summary>
		/// Wrapper interface for Subdocument which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISubdocument WithComCleanup(this Microsoft.Office.Interop.Word.Subdocument resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Subdocument, Interfaces.ISubdocument>();
		}

		/// <summary>
		/// Wrapper interface for HeadingStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadingStyles WithComCleanup(this Microsoft.Office.Interop.Word.HeadingStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeadingStyles, Interfaces.IHeadingStyles>();
		}

		/// <summary>
		/// Wrapper interface for HeadingStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadingStyle WithComCleanup(this Microsoft.Office.Interop.Word.HeadingStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeadingStyle, Interfaces.IHeadingStyle>();
		}

		/// <summary>
		/// Wrapper interface for StoryRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStoryRanges WithComCleanup(this Microsoft.Office.Interop.Word.StoryRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.StoryRanges, Interfaces.IStoryRanges>();
		}

		/// <summary>
		/// Wrapper interface for ListLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListLevel WithComCleanup(this Microsoft.Office.Interop.Word.ListLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListLevel, Interfaces.IListLevel>();
		}

		/// <summary>
		/// Wrapper interface for ListLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListLevels WithComCleanup(this Microsoft.Office.Interop.Word.ListLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListLevels, Interfaces.IListLevels>();
		}

		/// <summary>
		/// Wrapper interface for ListTemplate which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListTemplate WithComCleanup(this Microsoft.Office.Interop.Word.ListTemplate resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListTemplate, Interfaces.IListTemplate>();
		}

		/// <summary>
		/// Wrapper interface for ListTemplates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListTemplates WithComCleanup(this Microsoft.Office.Interop.Word.ListTemplates resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListTemplates, Interfaces.IListTemplates>();
		}

		/// <summary>
		/// Wrapper interface for ListParagraphs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListParagraphs WithComCleanup(this Microsoft.Office.Interop.Word.ListParagraphs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListParagraphs, Interfaces.IListParagraphs>();
		}

		/// <summary>
		/// Wrapper interface for List which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IList WithComCleanup(this Microsoft.Office.Interop.Word.List resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.List, Interfaces.IList>();
		}

		/// <summary>
		/// Wrapper interface for Lists which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILists WithComCleanup(this Microsoft.Office.Interop.Word.Lists resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Lists, Interfaces.ILists>();
		}

		/// <summary>
		/// Wrapper interface for ListGallery which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListGallery WithComCleanup(this Microsoft.Office.Interop.Word.ListGallery resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListGallery, Interfaces.IListGallery>();
		}

		/// <summary>
		/// Wrapper interface for ListGalleries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListGalleries WithComCleanup(this Microsoft.Office.Interop.Word.ListGalleries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListGalleries, Interfaces.IListGalleries>();
		}

		/// <summary>
		/// Wrapper interface for KeyBindings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IKeyBindings WithComCleanup(this Microsoft.Office.Interop.Word.KeyBindings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.KeyBindings, Interfaces.IKeyBindings>();
		}

		/// <summary>
		/// Wrapper interface for KeysBoundTo which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IKeysBoundTo WithComCleanup(this Microsoft.Office.Interop.Word.KeysBoundTo resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.KeysBoundTo, Interfaces.IKeysBoundTo>();
		}

		/// <summary>
		/// Wrapper interface for KeyBinding which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IKeyBinding WithComCleanup(this Microsoft.Office.Interop.Word.KeyBinding resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.KeyBinding, Interfaces.IKeyBinding>();
		}

		/// <summary>
		/// Wrapper interface for FileConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverter WithComCleanup(this Microsoft.Office.Interop.Word.FileConverter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FileConverter, Interfaces.IFileConverter>();
		}

		/// <summary>
		/// Wrapper interface for FileConverters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverters WithComCleanup(this Microsoft.Office.Interop.Word.FileConverters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FileConverters, Interfaces.IFileConverters>();
		}

		/// <summary>
		/// Wrapper interface for SynonymInfo which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISynonymInfo WithComCleanup(this Microsoft.Office.Interop.Word.SynonymInfo resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SynonymInfo, Interfaces.ISynonymInfo>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlinks WithComCleanup(this Microsoft.Office.Interop.Word.Hyperlinks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Hyperlinks, Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlink WithComCleanup(this Microsoft.Office.Interop.Word.Hyperlink resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Hyperlink, Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapes WithComCleanup(this Microsoft.Office.Interop.Word.Shapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Shapes, Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeRange WithComCleanup(this Microsoft.Office.Interop.Word.ShapeRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShapeRange, Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupShapes WithComCleanup(this Microsoft.Office.Interop.Word.GroupShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.GroupShapes, Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShape WithComCleanup(this Microsoft.Office.Interop.Word.Shape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Shape, Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame WithComCleanup(this Microsoft.Office.Interop.Word.TextFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextFrame, Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for _LetterContent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_LetterContent WithComCleanup(this Microsoft.Office.Interop.Word._LetterContent resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._LetterContent, Interfaces.I_LetterContent>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IView WithComCleanup(this Microsoft.Office.Interop.Word.View resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.View, Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for Zoom which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IZoom WithComCleanup(this Microsoft.Office.Interop.Word.Zoom resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Zoom, Interfaces.IZoom>();
		}

		/// <summary>
		/// Wrapper interface for Zooms which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IZooms WithComCleanup(this Microsoft.Office.Interop.Word.Zooms resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Zooms, Interfaces.IZooms>();
		}

		/// <summary>
		/// Wrapper interface for InlineShape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInlineShape WithComCleanup(this Microsoft.Office.Interop.Word.InlineShape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.InlineShape, Interfaces.IInlineShape>();
		}

		/// <summary>
		/// Wrapper interface for InlineShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInlineShapes WithComCleanup(this Microsoft.Office.Interop.Word.InlineShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.InlineShapes, Interfaces.IInlineShapes>();
		}

		/// <summary>
		/// Wrapper interface for SpellingSuggestions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpellingSuggestions WithComCleanup(this Microsoft.Office.Interop.Word.SpellingSuggestions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SpellingSuggestions, Interfaces.ISpellingSuggestions>();
		}

		/// <summary>
		/// Wrapper interface for SpellingSuggestion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpellingSuggestion WithComCleanup(this Microsoft.Office.Interop.Word.SpellingSuggestion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SpellingSuggestion, Interfaces.ISpellingSuggestion>();
		}

		/// <summary>
		/// Wrapper interface for Dictionaries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDictionaries WithComCleanup(this Microsoft.Office.Interop.Word.Dictionaries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dictionaries, Interfaces.IDictionaries>();
		}

		/// <summary>
		/// Wrapper interface for HangulHanjaConversionDictionaries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHangulHanjaConversionDictionaries WithComCleanup(this Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries, Interfaces.IHangulHanjaConversionDictionaries>();
		}

		/// <summary>
		/// Wrapper interface for Dictionary which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDictionary WithComCleanup(this Microsoft.Office.Interop.Word.Dictionary resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dictionary, Interfaces.IDictionary>();
		}

		/// <summary>
		/// Wrapper interface for ReadabilityStatistics which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReadabilityStatistics WithComCleanup(this Microsoft.Office.Interop.Word.ReadabilityStatistics resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ReadabilityStatistics, Interfaces.IReadabilityStatistics>();
		}

		/// <summary>
		/// Wrapper interface for ReadabilityStatistic which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReadabilityStatistic WithComCleanup(this Microsoft.Office.Interop.Word.ReadabilityStatistic resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ReadabilityStatistic, Interfaces.IReadabilityStatistic>();
		}

		/// <summary>
		/// Wrapper interface for Versions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVersions WithComCleanup(this Microsoft.Office.Interop.Word.Versions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Versions, Interfaces.IVersions>();
		}

		/// <summary>
		/// Wrapper interface for Version which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVersion WithComCleanup(this Microsoft.Office.Interop.Word.Version resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Version, Interfaces.IVersion>();
		}

		/// <summary>
		/// Wrapper interface for Options which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOptions WithComCleanup(this Microsoft.Office.Interop.Word.Options resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Options, Interfaces.IOptions>();
		}

		/// <summary>
		/// Wrapper interface for MailMessage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMessage WithComCleanup(this Microsoft.Office.Interop.Word.MailMessage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMessage, Interfaces.IMailMessage>();
		}

		/// <summary>
		/// Wrapper interface for ProofreadingErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProofreadingErrors WithComCleanup(this Microsoft.Office.Interop.Word.ProofreadingErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ProofreadingErrors, Interfaces.IProofreadingErrors>();
		}

		/// <summary>
		/// Wrapper interface for Mailer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailer WithComCleanup(this Microsoft.Office.Interop.Word.Mailer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Mailer, Interfaces.IMailer>();
		}

		/// <summary>
		/// Wrapper interface for WrapFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWrapFormat WithComCleanup(this Microsoft.Office.Interop.Word.WrapFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.WrapFormat, Interfaces.IWrapFormat>();
		}

		/// <summary>
		/// Wrapper interface for HangulAndAlphabetExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHangulAndAlphabetExceptions WithComCleanup(this Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions, Interfaces.IHangulAndAlphabetExceptions>();
		}

		/// <summary>
		/// Wrapper interface for HangulAndAlphabetException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHangulAndAlphabetException WithComCleanup(this Microsoft.Office.Interop.Word.HangulAndAlphabetException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HangulAndAlphabetException, Interfaces.IHangulAndAlphabetException>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAdjustments WithComCleanup(this Microsoft.Office.Interop.Word.Adjustments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Adjustments, Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalloutFormat WithComCleanup(this Microsoft.Office.Interop.Word.CalloutFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CalloutFormat, Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorFormat WithComCleanup(this Microsoft.Office.Interop.Word.ColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ColorFormat, Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnectorFormat WithComCleanup(this Microsoft.Office.Interop.Word.ConnectorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ConnectorFormat, Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFillFormat WithComCleanup(this Microsoft.Office.Interop.Word.FillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FillFormat, Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFreeformBuilder WithComCleanup(this Microsoft.Office.Interop.Word.FreeformBuilder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FreeformBuilder, Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineFormat WithComCleanup(this Microsoft.Office.Interop.Word.LineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LineFormat, Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureFormat WithComCleanup(this Microsoft.Office.Interop.Word.PictureFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PictureFormat, Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShadowFormat WithComCleanup(this Microsoft.Office.Interop.Word.ShadowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShadowFormat, Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNode WithComCleanup(this Microsoft.Office.Interop.Word.ShapeNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShapeNode, Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNodes WithComCleanup(this Microsoft.Office.Interop.Word.ShapeNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShapeNodes, Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextEffectFormat WithComCleanup(this Microsoft.Office.Interop.Word.TextEffectFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextEffectFormat, Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThreeDFormat WithComCleanup(this Microsoft.Office.Interop.Word.ThreeDFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ThreeDFormat, Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents, Interfaces.IApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlobal WithComCleanup(this Microsoft.Office.Interop.Word.Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Global, Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents_Event, Interfaces.IApplicationEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents2_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents2_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents2_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents2_Event, Interfaces.IApplicationEvents2_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents3_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents3_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents3_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents3_Event, Interfaces.IApplicationEvents3_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents4_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents4_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents4_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents4_Event, Interfaces.IApplicationEvents4_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplication WithComCleanup(this Microsoft.Office.Interop.Word.Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Application, Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents, Interfaces.IDocumentEvents>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents_Event WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents_Event, Interfaces.IDocumentEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents2_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents2_Event WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents2_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents2_Event, Interfaces.IDocumentEvents2_Event>();
		}

		/// <summary>
		/// Wrapper interface for Document which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocument WithComCleanup(this Microsoft.Office.Interop.Word.Document resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Document, Interfaces.IDocument>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFont WithComCleanup(this Microsoft.Office.Interop.Word.Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Font, Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphFormat WithComCleanup(this Microsoft.Office.Interop.Word.ParagraphFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ParagraphFormat, Interfaces.IParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for OCXEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXEvents WithComCleanup(this Microsoft.Office.Interop.Word.OCXEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OCXEvents, Interfaces.IOCXEvents>();
		}

		/// <summary>
		/// Wrapper interface for OCXEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXEvents_Event WithComCleanup(this Microsoft.Office.Interop.Word.OCXEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OCXEvents_Event, Interfaces.IOCXEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEControl WithComCleanup(this Microsoft.Office.Interop.Word.OLEControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OLEControl, Interfaces.IOLEControl>();
		}

		/// <summary>
		/// Wrapper interface for LetterContent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILetterContent WithComCleanup(this Microsoft.Office.Interop.Word.LetterContent resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LetterContent, Interfaces.ILetterContent>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents, Interfaces.IIApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents2 WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents2, Interfaces.IIApplicationEvents2>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents2 WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents2, Interfaces.IApplicationEvents2>();
		}

		/// <summary>
		/// Wrapper interface for EmailAuthor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailAuthor WithComCleanup(this Microsoft.Office.Interop.Word.EmailAuthor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailAuthor, Interfaces.IEmailAuthor>();
		}

		/// <summary>
		/// Wrapper interface for EmailOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailOptions WithComCleanup(this Microsoft.Office.Interop.Word.EmailOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailOptions, Interfaces.IEmailOptions>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignature which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailSignature WithComCleanup(this Microsoft.Office.Interop.Word.EmailSignature resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailSignature, Interfaces.IEmailSignature>();
		}

		/// <summary>
		/// Wrapper interface for Email which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmail WithComCleanup(this Microsoft.Office.Interop.Word.Email resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Email, Interfaces.IEmail>();
		}

		/// <summary>
		/// Wrapper interface for HorizontalLineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHorizontalLineFormat WithComCleanup(this Microsoft.Office.Interop.Word.HorizontalLineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HorizontalLineFormat, Interfaces.IHorizontalLineFormat>();
		}

		/// <summary>
		/// Wrapper interface for Frameset which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFrameset WithComCleanup(this Microsoft.Office.Interop.Word.Frameset resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Frameset, Interfaces.IFrameset>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDefaultWebOptions WithComCleanup(this Microsoft.Office.Interop.Word.DefaultWebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DefaultWebOptions, Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebOptions WithComCleanup(this Microsoft.Office.Interop.Word.WebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.WebOptions, Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for OtherCorrectionsExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOtherCorrectionsExceptions WithComCleanup(this Microsoft.Office.Interop.Word.OtherCorrectionsExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OtherCorrectionsExceptions, Interfaces.IOtherCorrectionsExceptions>();
		}

		/// <summary>
		/// Wrapper interface for OtherCorrectionsException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOtherCorrectionsException WithComCleanup(this Microsoft.Office.Interop.Word.OtherCorrectionsException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OtherCorrectionsException, Interfaces.IOtherCorrectionsException>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignatureEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailSignatureEntries WithComCleanup(this Microsoft.Office.Interop.Word.EmailSignatureEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailSignatureEntries, Interfaces.IEmailSignatureEntries>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignatureEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailSignatureEntry WithComCleanup(this Microsoft.Office.Interop.Word.EmailSignatureEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailSignatureEntry, Interfaces.IEmailSignatureEntry>();
		}

		/// <summary>
		/// Wrapper interface for HTMLDivision which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLDivision WithComCleanup(this Microsoft.Office.Interop.Word.HTMLDivision resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HTMLDivision, Interfaces.IHTMLDivision>();
		}

		/// <summary>
		/// Wrapper interface for HTMLDivisions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLDivisions WithComCleanup(this Microsoft.Office.Interop.Word.HTMLDivisions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HTMLDivisions, Interfaces.IHTMLDivisions>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNode WithComCleanup(this Microsoft.Office.Interop.Word.DiagramNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DiagramNode, Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodeChildren WithComCleanup(this Microsoft.Office.Interop.Word.DiagramNodeChildren resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DiagramNodeChildren, Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodes WithComCleanup(this Microsoft.Office.Interop.Word.DiagramNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DiagramNodes, Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagram WithComCleanup(this Microsoft.Office.Interop.Word.Diagram resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Diagram, Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomProperty WithComCleanup(this Microsoft.Office.Interop.Word.CustomProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomProperty, Interfaces.ICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomProperties WithComCleanup(this Microsoft.Office.Interop.Word.CustomProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomProperties, Interfaces.ICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for SmartTag which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTag WithComCleanup(this Microsoft.Office.Interop.Word.SmartTag resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTag, Interfaces.ISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for SmartTags which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTags WithComCleanup(this Microsoft.Office.Interop.Word.SmartTags resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTags, Interfaces.ISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for StyleSheet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyleSheet WithComCleanup(this Microsoft.Office.Interop.Word.StyleSheet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.StyleSheet, Interfaces.IStyleSheet>();
		}

		/// <summary>
		/// Wrapper interface for StyleSheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyleSheets WithComCleanup(this Microsoft.Office.Interop.Word.StyleSheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.StyleSheets, Interfaces.IStyleSheets>();
		}

		/// <summary>
		/// Wrapper interface for MappedDataField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMappedDataField WithComCleanup(this Microsoft.Office.Interop.Word.MappedDataField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MappedDataField, Interfaces.IMappedDataField>();
		}

		/// <summary>
		/// Wrapper interface for MappedDataFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMappedDataFields WithComCleanup(this Microsoft.Office.Interop.Word.MappedDataFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MappedDataFields, Interfaces.IMappedDataFields>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICanvasShapes WithComCleanup(this Microsoft.Office.Interop.Word.CanvasShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CanvasShapes, Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyle WithComCleanup(this Microsoft.Office.Interop.Word.TableStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableStyle, Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for ConditionalStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConditionalStyle WithComCleanup(this Microsoft.Office.Interop.Word.ConditionalStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ConditionalStyle, Interfaces.IConditionalStyle>();
		}

		/// <summary>
		/// Wrapper interface for FootnoteOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFootnoteOptions WithComCleanup(this Microsoft.Office.Interop.Word.FootnoteOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FootnoteOptions, Interfaces.IFootnoteOptions>();
		}

		/// <summary>
		/// Wrapper interface for EndnoteOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEndnoteOptions WithComCleanup(this Microsoft.Office.Interop.Word.EndnoteOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EndnoteOptions, Interfaces.IEndnoteOptions>();
		}

		/// <summary>
		/// Wrapper interface for Reviewers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReviewers WithComCleanup(this Microsoft.Office.Interop.Word.Reviewers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Reviewers, Interfaces.IReviewers>();
		}

		/// <summary>
		/// Wrapper interface for Reviewer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReviewer WithComCleanup(this Microsoft.Office.Interop.Word.Reviewer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Reviewer, Interfaces.IReviewer>();
		}

		/// <summary>
		/// Wrapper interface for TaskPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskPane WithComCleanup(this Microsoft.Office.Interop.Word.TaskPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TaskPane, Interfaces.ITaskPane>();
		}

		/// <summary>
		/// Wrapper interface for TaskPanes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskPanes WithComCleanup(this Microsoft.Office.Interop.Word.TaskPanes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TaskPanes, Interfaces.ITaskPanes>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents3 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents3 WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents3 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents3, Interfaces.IIApplicationEvents3>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents3 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents3 WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents3 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents3, Interfaces.IApplicationEvents3>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagAction WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagAction, Interfaces.ISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagActions WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagActions, Interfaces.ISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagRecognizer WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagRecognizer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagRecognizer, Interfaces.ISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagRecognizers WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagRecognizers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagRecognizers, Interfaces.ISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagType which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagType WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagType resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagType, Interfaces.ISmartTagType>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagTypes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagTypes WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagTypes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagTypes, Interfaces.ISmartTagTypes>();
		}

		/// <summary>
		/// Wrapper interface for Line which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILine WithComCleanup(this Microsoft.Office.Interop.Word.Line resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Line, Interfaces.ILine>();
		}

		/// <summary>
		/// Wrapper interface for Lines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILines WithComCleanup(this Microsoft.Office.Interop.Word.Lines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Lines, Interfaces.ILines>();
		}

		/// <summary>
		/// Wrapper interface for Rectangle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangle WithComCleanup(this Microsoft.Office.Interop.Word.Rectangle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Rectangle, Interfaces.IRectangle>();
		}

		/// <summary>
		/// Wrapper interface for Rectangles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangles WithComCleanup(this Microsoft.Office.Interop.Word.Rectangles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Rectangles, Interfaces.IRectangles>();
		}

		/// <summary>
		/// Wrapper interface for Break which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBreak WithComCleanup(this Microsoft.Office.Interop.Word.Break resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Break, Interfaces.IBreak>();
		}

		/// <summary>
		/// Wrapper interface for Breaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBreaks WithComCleanup(this Microsoft.Office.Interop.Word.Breaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Breaks, Interfaces.IBreaks>();
		}

		/// <summary>
		/// Wrapper interface for Page which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPage WithComCleanup(this Microsoft.Office.Interop.Word.Page resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Page, Interfaces.IPage>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPages WithComCleanup(this Microsoft.Office.Interop.Word.Pages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Pages, Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for XMLNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNode WithComCleanup(this Microsoft.Office.Interop.Word.XMLNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNode, Interfaces.IXMLNode>();
		}

		/// <summary>
		/// Wrapper interface for XMLNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNodes WithComCleanup(this Microsoft.Office.Interop.Word.XMLNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNodes, Interfaces.IXMLNodes>();
		}

		/// <summary>
		/// Wrapper interface for XMLSchemaReference which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLSchemaReference WithComCleanup(this Microsoft.Office.Interop.Word.XMLSchemaReference resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLSchemaReference, Interfaces.IXMLSchemaReference>();
		}

		/// <summary>
		/// Wrapper interface for XMLSchemaReferences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLSchemaReferences WithComCleanup(this Microsoft.Office.Interop.Word.XMLSchemaReferences resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLSchemaReferences, Interfaces.IXMLSchemaReferences>();
		}

		/// <summary>
		/// Wrapper interface for XMLChildNodeSuggestion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLChildNodeSuggestion WithComCleanup(this Microsoft.Office.Interop.Word.XMLChildNodeSuggestion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLChildNodeSuggestion, Interfaces.IXMLChildNodeSuggestion>();
		}

		/// <summary>
		/// Wrapper interface for XMLChildNodeSuggestions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLChildNodeSuggestions WithComCleanup(this Microsoft.Office.Interop.Word.XMLChildNodeSuggestions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLChildNodeSuggestions, Interfaces.IXMLChildNodeSuggestions>();
		}

		/// <summary>
		/// Wrapper interface for XMLNamespace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNamespace WithComCleanup(this Microsoft.Office.Interop.Word.XMLNamespace resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNamespace, Interfaces.IXMLNamespace>();
		}

		/// <summary>
		/// Wrapper interface for XMLNamespaces which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNamespaces WithComCleanup(this Microsoft.Office.Interop.Word.XMLNamespaces resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNamespaces, Interfaces.IXMLNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for XSLTransform which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXSLTransform WithComCleanup(this Microsoft.Office.Interop.Word.XSLTransform resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XSLTransform, Interfaces.IXSLTransform>();
		}

		/// <summary>
		/// Wrapper interface for XSLTransforms which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXSLTransforms WithComCleanup(this Microsoft.Office.Interop.Word.XSLTransforms resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XSLTransforms, Interfaces.IXSLTransforms>();
		}

		/// <summary>
		/// Wrapper interface for Editors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEditors WithComCleanup(this Microsoft.Office.Interop.Word.Editors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Editors, Interfaces.IEditors>();
		}

		/// <summary>
		/// Wrapper interface for Editor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEditor WithComCleanup(this Microsoft.Office.Interop.Word.Editor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Editor, Interfaces.IEditor>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents4 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents4 WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents4 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents4, Interfaces.IIApplicationEvents4>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents4 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents4 WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents4 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents4, Interfaces.IApplicationEvents4>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents2 WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents2, Interfaces.IDocumentEvents2>();
		}

		/// <summary>
		/// Wrapper interface for Source which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISource WithComCleanup(this Microsoft.Office.Interop.Word.Source resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Source, Interfaces.ISource>();
		}

		/// <summary>
		/// Wrapper interface for Sources which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISources WithComCleanup(this Microsoft.Office.Interop.Word.Sources resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Sources, Interfaces.ISources>();
		}

		/// <summary>
		/// Wrapper interface for Bibliography which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBibliography WithComCleanup(this Microsoft.Office.Interop.Word.Bibliography resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Bibliography, Interfaces.IBibliography>();
		}

		/// <summary>
		/// Wrapper interface for OMaths which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMaths WithComCleanup(this Microsoft.Office.Interop.Word.OMaths resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMaths, Interfaces.IOMaths>();
		}

		/// <summary>
		/// Wrapper interface for OMath which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMath WithComCleanup(this Microsoft.Office.Interop.Word.OMath resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMath, Interfaces.IOMath>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunctions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFunctions WithComCleanup(this Microsoft.Office.Interop.Word.OMathFunctions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFunctions, Interfaces.IOMathFunctions>();
		}

		/// <summary>
		/// Wrapper interface for OMathArgs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathArgs WithComCleanup(this Microsoft.Office.Interop.Word.OMathArgs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathArgs, Interfaces.IOMathArgs>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFunction WithComCleanup(this Microsoft.Office.Interop.Word.OMathFunction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFunction, Interfaces.IOMathFunction>();
		}

		/// <summary>
		/// Wrapper interface for OMathAcc which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAcc WithComCleanup(this Microsoft.Office.Interop.Word.OMathAcc resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAcc, Interfaces.IOMathAcc>();
		}

		/// <summary>
		/// Wrapper interface for OMathBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBar WithComCleanup(this Microsoft.Office.Interop.Word.OMathBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBar, Interfaces.IOMathBar>();
		}

		/// <summary>
		/// Wrapper interface for OMathBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBox WithComCleanup(this Microsoft.Office.Interop.Word.OMathBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBox, Interfaces.IOMathBox>();
		}

		/// <summary>
		/// Wrapper interface for OMathBorderBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBorderBox WithComCleanup(this Microsoft.Office.Interop.Word.OMathBorderBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBorderBox, Interfaces.IOMathBorderBox>();
		}

		/// <summary>
		/// Wrapper interface for OMathDelim which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathDelim WithComCleanup(this Microsoft.Office.Interop.Word.OMathDelim resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathDelim, Interfaces.IOMathDelim>();
		}

		/// <summary>
		/// Wrapper interface for OMathEqArray which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathEqArray WithComCleanup(this Microsoft.Office.Interop.Word.OMathEqArray resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathEqArray, Interfaces.IOMathEqArray>();
		}

		/// <summary>
		/// Wrapper interface for OMathFrac which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFrac WithComCleanup(this Microsoft.Office.Interop.Word.OMathFrac resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFrac, Interfaces.IOMathFrac>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunc which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFunc WithComCleanup(this Microsoft.Office.Interop.Word.OMathFunc resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFunc, Interfaces.IOMathFunc>();
		}

		/// <summary>
		/// Wrapper interface for OMathGroupChar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathGroupChar WithComCleanup(this Microsoft.Office.Interop.Word.OMathGroupChar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathGroupChar, Interfaces.IOMathGroupChar>();
		}

		/// <summary>
		/// Wrapper interface for OMathLimLow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathLimLow WithComCleanup(this Microsoft.Office.Interop.Word.OMathLimLow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathLimLow, Interfaces.IOMathLimLow>();
		}

		/// <summary>
		/// Wrapper interface for OMathLimUpp which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathLimUpp WithComCleanup(this Microsoft.Office.Interop.Word.OMathLimUpp resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathLimUpp, Interfaces.IOMathLimUpp>();
		}

		/// <summary>
		/// Wrapper interface for OMathMat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMat WithComCleanup(this Microsoft.Office.Interop.Word.OMathMat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMat, Interfaces.IOMathMat>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatRows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatRows WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatRows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatRows, Interfaces.IOMathMatRows>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatCols which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatCols WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatCols resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatCols, Interfaces.IOMathMatCols>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatRow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatRow WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatRow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatRow, Interfaces.IOMathMatRow>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatCol which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatCol WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatCol resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatCol, Interfaces.IOMathMatCol>();
		}

		/// <summary>
		/// Wrapper interface for OMathNary which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathNary WithComCleanup(this Microsoft.Office.Interop.Word.OMathNary resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathNary, Interfaces.IOMathNary>();
		}

		/// <summary>
		/// Wrapper interface for OMathPhantom which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathPhantom WithComCleanup(this Microsoft.Office.Interop.Word.OMathPhantom resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathPhantom, Interfaces.IOMathPhantom>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrPre which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrPre WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrPre resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrPre, Interfaces.IOMathScrPre>();
		}

		/// <summary>
		/// Wrapper interface for OMathRad which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathRad WithComCleanup(this Microsoft.Office.Interop.Word.OMathRad resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathRad, Interfaces.IOMathRad>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSub which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrSub WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrSub resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrSub, Interfaces.IOMathScrSub>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSubSup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrSubSup WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrSubSup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrSubSup, Interfaces.IOMathScrSubSup>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrSup WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrSup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrSup, Interfaces.IOMathScrSup>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAutoCorrect WithComCleanup(this Microsoft.Office.Interop.Word.OMathAutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAutoCorrect, Interfaces.IOMathAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrectEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAutoCorrectEntries WithComCleanup(this Microsoft.Office.Interop.Word.OMathAutoCorrectEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAutoCorrectEntries, Interfaces.IOMathAutoCorrectEntries>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrectEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAutoCorrectEntry WithComCleanup(this Microsoft.Office.Interop.Word.OMathAutoCorrectEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAutoCorrectEntry, Interfaces.IOMathAutoCorrectEntry>();
		}

		/// <summary>
		/// Wrapper interface for OMathRecognizedFunctions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathRecognizedFunctions WithComCleanup(this Microsoft.Office.Interop.Word.OMathRecognizedFunctions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathRecognizedFunctions, Interfaces.IOMathRecognizedFunctions>();
		}

		/// <summary>
		/// Wrapper interface for OMathRecognizedFunction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathRecognizedFunction WithComCleanup(this Microsoft.Office.Interop.Word.OMathRecognizedFunction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathRecognizedFunction, Interfaces.IOMathRecognizedFunction>();
		}

		/// <summary>
		/// Wrapper interface for ContentControls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControls WithComCleanup(this Microsoft.Office.Interop.Word.ContentControls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControls, Interfaces.IContentControls>();
		}

		/// <summary>
		/// Wrapper interface for ContentControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControl WithComCleanup(this Microsoft.Office.Interop.Word.ContentControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControl, Interfaces.IContentControl>();
		}

		/// <summary>
		/// Wrapper interface for XMLMapping which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLMapping WithComCleanup(this Microsoft.Office.Interop.Word.XMLMapping resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLMapping, Interfaces.IXMLMapping>();
		}

		/// <summary>
		/// Wrapper interface for ContentControlListEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControlListEntries WithComCleanup(this Microsoft.Office.Interop.Word.ContentControlListEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControlListEntries, Interfaces.IContentControlListEntries>();
		}

		/// <summary>
		/// Wrapper interface for ContentControlListEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControlListEntry WithComCleanup(this Microsoft.Office.Interop.Word.ContentControlListEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControlListEntry, Interfaces.IContentControlListEntry>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockTypes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlockTypes WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlockTypes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlockTypes, Interfaces.IBuildingBlockTypes>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockType which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlockType WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlockType resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlockType, Interfaces.IBuildingBlockType>();
		}

		/// <summary>
		/// Wrapper interface for Categories which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategories WithComCleanup(this Microsoft.Office.Interop.Word.Categories resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Categories, Interfaces.ICategories>();
		}

		/// <summary>
		/// Wrapper interface for Category which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategory WithComCleanup(this Microsoft.Office.Interop.Word.Category resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Category, Interfaces.ICategory>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlocks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlocks WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlocks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlocks, Interfaces.IBuildingBlocks>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlock which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlock WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlock resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlock, Interfaces.IBuildingBlock>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlockEntries WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlockEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlockEntries, Interfaces.IBuildingBlockEntries>();
		}

		/// <summary>
		/// Wrapper interface for OMathBreaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBreaks WithComCleanup(this Microsoft.Office.Interop.Word.OMathBreaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBreaks, Interfaces.IOMathBreaks>();
		}

		/// <summary>
		/// Wrapper interface for OMathBreak which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBreak WithComCleanup(this Microsoft.Office.Interop.Word.OMathBreak resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBreak, Interfaces.IOMathBreak>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResearch WithComCleanup(this Microsoft.Office.Interop.Word.Research resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Research, Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoftEdgeFormat WithComCleanup(this Microsoft.Office.Interop.Word.SoftEdgeFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SoftEdgeFormat, Interfaces.ISoftEdgeFormat>();
		}

		/// <summary>
		/// Wrapper interface for GlowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlowFormat WithComCleanup(this Microsoft.Office.Interop.Word.GlowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.GlowFormat, Interfaces.IGlowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReflectionFormat WithComCleanup(this Microsoft.Office.Interop.Word.ReflectionFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ReflectionFormat, Interfaces.IReflectionFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartData WithComCleanup(this Microsoft.Office.Interop.Word.ChartData resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartData, Interfaces.IChartData>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChart WithComCleanup(this Microsoft.Office.Interop.Word.Chart resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Chart, Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICorners WithComCleanup(this Microsoft.Office.Interop.Word.Corners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Corners, Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegend WithComCleanup(this Microsoft.Office.Interop.Word.Legend resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Legend, Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for ChartBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartBorder WithComCleanup(this Microsoft.Office.Interop.Word.ChartBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartBorder, Interfaces.IChartBorder>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWalls WithComCleanup(this Microsoft.Office.Interop.Word.Walls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Walls, Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFloor WithComCleanup(this Microsoft.Office.Interop.Word.Floor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Floor, Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlotArea WithComCleanup(this Microsoft.Office.Interop.Word.PlotArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PlotArea, Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartArea WithComCleanup(this Microsoft.Office.Interop.Word.ChartArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartArea, Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesLines WithComCleanup(this Microsoft.Office.Interop.Word.SeriesLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SeriesLines, Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILeaderLines WithComCleanup(this Microsoft.Office.Interop.Word.LeaderLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LeaderLines, Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGridlines WithComCleanup(this Microsoft.Office.Interop.Word.Gridlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Gridlines, Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUpBars WithComCleanup(this Microsoft.Office.Interop.Word.UpBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.UpBars, Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDownBars WithComCleanup(this Microsoft.Office.Interop.Word.DownBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DownBars, Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInterior WithComCleanup(this Microsoft.Office.Interop.Word.Interior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Interior, Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFillFormat WithComCleanup(this Microsoft.Office.Interop.Word.ChartFillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartFillFormat, Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntries WithComCleanup(this Microsoft.Office.Interop.Word.LegendEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LegendEntries, Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFont WithComCleanup(this Microsoft.Office.Interop.Word.ChartFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartFont, Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartColorFormat WithComCleanup(this Microsoft.Office.Interop.Word.ChartColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartColorFormat, Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntry WithComCleanup(this Microsoft.Office.Interop.Word.LegendEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LegendEntry, Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendKey WithComCleanup(this Microsoft.Office.Interop.Word.LegendKey resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LegendKey, Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesCollection WithComCleanup(this Microsoft.Office.Interop.Word.SeriesCollection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SeriesCollection, Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeries WithComCleanup(this Microsoft.Office.Interop.Word.Series resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Series, Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrorBars WithComCleanup(this Microsoft.Office.Interop.Word.ErrorBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ErrorBars, Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendline WithComCleanup(this Microsoft.Office.Interop.Word.Trendline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Trendline, Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendlines WithComCleanup(this Microsoft.Office.Interop.Word.Trendlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Trendlines, Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabels WithComCleanup(this Microsoft.Office.Interop.Word.DataLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DataLabels, Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabel WithComCleanup(this Microsoft.Office.Interop.Word.DataLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DataLabel, Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoints WithComCleanup(this Microsoft.Office.Interop.Word.Points resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Points, Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoint WithComCleanup(this Microsoft.Office.Interop.Word.Point resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Point, Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxes WithComCleanup(this Microsoft.Office.Interop.Word.Axes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Axes, Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxis WithComCleanup(this Microsoft.Office.Interop.Word.Axis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Axis, Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataTable WithComCleanup(this Microsoft.Office.Interop.Word.DataTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DataTable, Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartTitle WithComCleanup(this Microsoft.Office.Interop.Word.ChartTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartTitle, Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxisTitle WithComCleanup(this Microsoft.Office.Interop.Word.AxisTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AxisTitle, Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDisplayUnitLabel WithComCleanup(this Microsoft.Office.Interop.Word.DisplayUnitLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DisplayUnitLabel, Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITickLabels WithComCleanup(this Microsoft.Office.Interop.Word.TickLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TickLabels, Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropLines WithComCleanup(this Microsoft.Office.Interop.Word.DropLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DropLines, Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHiLoLines WithComCleanup(this Microsoft.Office.Interop.Word.HiLoLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HiLoLines, Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroup WithComCleanup(this Microsoft.Office.Interop.Word.ChartGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartGroup, Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroups WithComCleanup(this Microsoft.Office.Interop.Word.ChartGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartGroups, Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for ChartCharacters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartCharacters WithComCleanup(this Microsoft.Office.Interop.Word.ChartCharacters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartCharacters, Interfaces.IChartCharacters>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFormat WithComCleanup(this Microsoft.Office.Interop.Word.ChartFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartFormat, Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for UndoRecord which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUndoRecord WithComCleanup(this Microsoft.Office.Interop.Word.UndoRecord resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.UndoRecord, Interfaces.IUndoRecord>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthLock which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthLock WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthLock resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthLock, Interfaces.ICoAuthLock>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthLocks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthLocks WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthLocks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthLocks, Interfaces.ICoAuthLocks>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthUpdate which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthUpdate WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthUpdate resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthUpdate, Interfaces.ICoAuthUpdate>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthUpdates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthUpdates WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthUpdates resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthUpdates, Interfaces.ICoAuthUpdates>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthor WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthor, Interfaces.ICoAuthor>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthors WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthors, Interfaces.ICoAuthors>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthoring which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthoring WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthoring resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthoring, Interfaces.ICoAuthoring>();
		}

		/// <summary>
		/// Wrapper interface for Conflicts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConflicts WithComCleanup(this Microsoft.Office.Interop.Word.Conflicts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Conflicts, Interfaces.IConflicts>();
		}

		/// <summary>
		/// Wrapper interface for Conflict which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConflict WithComCleanup(this Microsoft.Office.Interop.Word.Conflict resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Conflict, Interfaces.IConflict>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindows WithComCleanup(this Microsoft.Office.Interop.Word.ProtectedViewWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ProtectedViewWindows, Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindow WithComCleanup(this Microsoft.Office.Interop.Word.ProtectedViewWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ProtectedViewWindow, Interfaces.IProtectedViewWindow>();
		}

	}
}