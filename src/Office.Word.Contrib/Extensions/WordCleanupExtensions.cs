//Microsoft.Office.Interop.Word, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace Office.Contrib.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_Application WithComCleanup(this Microsoft.Office.Interop.Word._Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Application, Word.Contrib.Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_Global WithComCleanup(this Microsoft.Office.Interop.Word._Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Global, Word.Contrib.Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for FontNames which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFontNames WithComCleanup(this Microsoft.Office.Interop.Word.FontNames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FontNames, Word.Contrib.Interfaces.IFontNames>();
		}

		/// <summary>
		/// Wrapper interface for Languages which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILanguages WithComCleanup(this Microsoft.Office.Interop.Word.Languages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Languages, Word.Contrib.Interfaces.ILanguages>();
		}

		/// <summary>
		/// Wrapper interface for Language which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILanguage WithComCleanup(this Microsoft.Office.Interop.Word.Language resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Language, Word.Contrib.Interfaces.ILanguage>();
		}

		/// <summary>
		/// Wrapper interface for Documents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDocuments WithComCleanup(this Microsoft.Office.Interop.Word.Documents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Documents, Word.Contrib.Interfaces.IDocuments>();
		}

		/// <summary>
		/// Wrapper interface for _Document which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_Document WithComCleanup(this Microsoft.Office.Interop.Word._Document resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Document, Word.Contrib.Interfaces.I_Document>();
		}

		/// <summary>
		/// Wrapper interface for Template which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITemplate WithComCleanup(this Microsoft.Office.Interop.Word.Template resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Template, Word.Contrib.Interfaces.ITemplate>();
		}

		/// <summary>
		/// Wrapper interface for Templates which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITemplates WithComCleanup(this Microsoft.Office.Interop.Word.Templates resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Templates, Word.Contrib.Interfaces.ITemplates>();
		}

		/// <summary>
		/// Wrapper interface for RoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRoutingSlip WithComCleanup(this Microsoft.Office.Interop.Word.RoutingSlip resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.RoutingSlip, Word.Contrib.Interfaces.IRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for Bookmark which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBookmark WithComCleanup(this Microsoft.Office.Interop.Word.Bookmark resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Bookmark, Word.Contrib.Interfaces.IBookmark>();
		}

		/// <summary>
		/// Wrapper interface for Bookmarks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBookmarks WithComCleanup(this Microsoft.Office.Interop.Word.Bookmarks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Bookmarks, Word.Contrib.Interfaces.IBookmarks>();
		}

		/// <summary>
		/// Wrapper interface for Variable which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IVariable WithComCleanup(this Microsoft.Office.Interop.Word.Variable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Variable, Word.Contrib.Interfaces.IVariable>();
		}

		/// <summary>
		/// Wrapper interface for Variables which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IVariables WithComCleanup(this Microsoft.Office.Interop.Word.Variables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Variables, Word.Contrib.Interfaces.IVariables>();
		}

		/// <summary>
		/// Wrapper interface for RecentFile which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRecentFile WithComCleanup(this Microsoft.Office.Interop.Word.RecentFile resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.RecentFile, Word.Contrib.Interfaces.IRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for RecentFiles which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRecentFiles WithComCleanup(this Microsoft.Office.Interop.Word.RecentFiles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.RecentFiles, Word.Contrib.Interfaces.IRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for Window which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IWindow WithComCleanup(this Microsoft.Office.Interop.Word.Window resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Window, Word.Contrib.Interfaces.IWindow>();
		}

		/// <summary>
		/// Wrapper interface for Windows which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IWindows WithComCleanup(this Microsoft.Office.Interop.Word.Windows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Windows, Word.Contrib.Interfaces.IWindows>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPane WithComCleanup(this Microsoft.Office.Interop.Word.Pane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Pane, Word.Contrib.Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPanes WithComCleanup(this Microsoft.Office.Interop.Word.Panes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Panes, Word.Contrib.Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Range which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRange WithComCleanup(this Microsoft.Office.Interop.Word.Range resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Range, Word.Contrib.Interfaces.IRange>();
		}

		/// <summary>
		/// Wrapper interface for ListFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListFormat WithComCleanup(this Microsoft.Office.Interop.Word.ListFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListFormat, Word.Contrib.Interfaces.IListFormat>();
		}

		/// <summary>
		/// Wrapper interface for Find which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFind WithComCleanup(this Microsoft.Office.Interop.Word.Find resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Find, Word.Contrib.Interfaces.IFind>();
		}

		/// <summary>
		/// Wrapper interface for Replacement which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IReplacement WithComCleanup(this Microsoft.Office.Interop.Word.Replacement resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Replacement, Word.Contrib.Interfaces.IReplacement>();
		}

		/// <summary>
		/// Wrapper interface for Characters which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICharacters WithComCleanup(this Microsoft.Office.Interop.Word.Characters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Characters, Word.Contrib.Interfaces.ICharacters>();
		}

		/// <summary>
		/// Wrapper interface for Words which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IWords WithComCleanup(this Microsoft.Office.Interop.Word.Words resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Words, Word.Contrib.Interfaces.IWords>();
		}

		/// <summary>
		/// Wrapper interface for Sentences which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISentences WithComCleanup(this Microsoft.Office.Interop.Word.Sentences resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Sentences, Word.Contrib.Interfaces.ISentences>();
		}

		/// <summary>
		/// Wrapper interface for Sections which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISections WithComCleanup(this Microsoft.Office.Interop.Word.Sections resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Sections, Word.Contrib.Interfaces.ISections>();
		}

		/// <summary>
		/// Wrapper interface for Section which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISection WithComCleanup(this Microsoft.Office.Interop.Word.Section resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Section, Word.Contrib.Interfaces.ISection>();
		}

		/// <summary>
		/// Wrapper interface for Paragraphs which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IParagraphs WithComCleanup(this Microsoft.Office.Interop.Word.Paragraphs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Paragraphs, Word.Contrib.Interfaces.IParagraphs>();
		}

		/// <summary>
		/// Wrapper interface for Paragraph which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IParagraph WithComCleanup(this Microsoft.Office.Interop.Word.Paragraph resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Paragraph, Word.Contrib.Interfaces.IParagraph>();
		}

		/// <summary>
		/// Wrapper interface for DropCap which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDropCap WithComCleanup(this Microsoft.Office.Interop.Word.DropCap resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DropCap, Word.Contrib.Interfaces.IDropCap>();
		}

		/// <summary>
		/// Wrapper interface for TabStops which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITabStops WithComCleanup(this Microsoft.Office.Interop.Word.TabStops resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TabStops, Word.Contrib.Interfaces.ITabStops>();
		}

		/// <summary>
		/// Wrapper interface for TabStop which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITabStop WithComCleanup(this Microsoft.Office.Interop.Word.TabStop resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TabStop, Word.Contrib.Interfaces.ITabStop>();
		}

		/// <summary>
		/// Wrapper interface for _ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_ParagraphFormat WithComCleanup(this Microsoft.Office.Interop.Word._ParagraphFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._ParagraphFormat, Word.Contrib.Interfaces.I_ParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for _Font which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_Font WithComCleanup(this Microsoft.Office.Interop.Word._Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._Font, Word.Contrib.Interfaces.I_Font>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITable WithComCleanup(this Microsoft.Office.Interop.Word.Table resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Table, Word.Contrib.Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRow WithComCleanup(this Microsoft.Office.Interop.Word.Row resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Row, Word.Contrib.Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IColumn WithComCleanup(this Microsoft.Office.Interop.Word.Column resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Column, Word.Contrib.Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for Cell which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICell WithComCleanup(this Microsoft.Office.Interop.Word.Cell resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Cell, Word.Contrib.Interfaces.ICell>();
		}

		/// <summary>
		/// Wrapper interface for Tables which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITables WithComCleanup(this Microsoft.Office.Interop.Word.Tables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Tables, Word.Contrib.Interfaces.ITables>();
		}

		/// <summary>
		/// Wrapper interface for Rows which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRows WithComCleanup(this Microsoft.Office.Interop.Word.Rows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Rows, Word.Contrib.Interfaces.IRows>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IColumns WithComCleanup(this Microsoft.Office.Interop.Word.Columns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Columns, Word.Contrib.Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for Cells which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICells WithComCleanup(this Microsoft.Office.Interop.Word.Cells resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Cells, Word.Contrib.Interfaces.ICells>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoCorrect WithComCleanup(this Microsoft.Office.Interop.Word.AutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCorrect, Word.Contrib.Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrectEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoCorrectEntries WithComCleanup(this Microsoft.Office.Interop.Word.AutoCorrectEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCorrectEntries, Word.Contrib.Interfaces.IAutoCorrectEntries>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrectEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoCorrectEntry WithComCleanup(this Microsoft.Office.Interop.Word.AutoCorrectEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCorrectEntry, Word.Contrib.Interfaces.IAutoCorrectEntry>();
		}

		/// <summary>
		/// Wrapper interface for FirstLetterExceptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFirstLetterExceptions WithComCleanup(this Microsoft.Office.Interop.Word.FirstLetterExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FirstLetterExceptions, Word.Contrib.Interfaces.IFirstLetterExceptions>();
		}

		/// <summary>
		/// Wrapper interface for FirstLetterException which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFirstLetterException WithComCleanup(this Microsoft.Office.Interop.Word.FirstLetterException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FirstLetterException, Word.Contrib.Interfaces.IFirstLetterException>();
		}

		/// <summary>
		/// Wrapper interface for TwoInitialCapsExceptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITwoInitialCapsExceptions WithComCleanup(this Microsoft.Office.Interop.Word.TwoInitialCapsExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TwoInitialCapsExceptions, Word.Contrib.Interfaces.ITwoInitialCapsExceptions>();
		}

		/// <summary>
		/// Wrapper interface for TwoInitialCapsException which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITwoInitialCapsException WithComCleanup(this Microsoft.Office.Interop.Word.TwoInitialCapsException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TwoInitialCapsException, Word.Contrib.Interfaces.ITwoInitialCapsException>();
		}

		/// <summary>
		/// Wrapper interface for Footnotes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFootnotes WithComCleanup(this Microsoft.Office.Interop.Word.Footnotes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Footnotes, Word.Contrib.Interfaces.IFootnotes>();
		}

		/// <summary>
		/// Wrapper interface for Endnotes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEndnotes WithComCleanup(this Microsoft.Office.Interop.Word.Endnotes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Endnotes, Word.Contrib.Interfaces.IEndnotes>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IComments WithComCleanup(this Microsoft.Office.Interop.Word.Comments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Comments, Word.Contrib.Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Footnote which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFootnote WithComCleanup(this Microsoft.Office.Interop.Word.Footnote resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Footnote, Word.Contrib.Interfaces.IFootnote>();
		}

		/// <summary>
		/// Wrapper interface for Endnote which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEndnote WithComCleanup(this Microsoft.Office.Interop.Word.Endnote resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Endnote, Word.Contrib.Interfaces.IEndnote>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IComment WithComCleanup(this Microsoft.Office.Interop.Word.Comment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Comment, Word.Contrib.Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBorders WithComCleanup(this Microsoft.Office.Interop.Word.Borders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Borders, Word.Contrib.Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for Border which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBorder WithComCleanup(this Microsoft.Office.Interop.Word.Border resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Border, Word.Contrib.Interfaces.IBorder>();
		}

		/// <summary>
		/// Wrapper interface for Shading which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShading WithComCleanup(this Microsoft.Office.Interop.Word.Shading resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Shading, Word.Contrib.Interfaces.IShading>();
		}

		/// <summary>
		/// Wrapper interface for TextRetrievalMode which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITextRetrievalMode WithComCleanup(this Microsoft.Office.Interop.Word.TextRetrievalMode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextRetrievalMode, Word.Contrib.Interfaces.ITextRetrievalMode>();
		}

		/// <summary>
		/// Wrapper interface for AutoTextEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoTextEntries WithComCleanup(this Microsoft.Office.Interop.Word.AutoTextEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoTextEntries, Word.Contrib.Interfaces.IAutoTextEntries>();
		}

		/// <summary>
		/// Wrapper interface for AutoTextEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoTextEntry WithComCleanup(this Microsoft.Office.Interop.Word.AutoTextEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoTextEntry, Word.Contrib.Interfaces.IAutoTextEntry>();
		}

		/// <summary>
		/// Wrapper interface for System which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISystem WithComCleanup(this Microsoft.Office.Interop.Word.System resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.System, Word.Contrib.Interfaces.ISystem>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOLEFormat WithComCleanup(this Microsoft.Office.Interop.Word.OLEFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OLEFormat, Word.Contrib.Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILinkFormat WithComCleanup(this Microsoft.Office.Interop.Word.LinkFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LinkFormat, Word.Contrib.Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for _OLEControl which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_OLEControl WithComCleanup(this Microsoft.Office.Interop.Word._OLEControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._OLEControl, Word.Contrib.Interfaces.I_OLEControl>();
		}

		/// <summary>
		/// Wrapper interface for Fields which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFields WithComCleanup(this Microsoft.Office.Interop.Word.Fields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Fields, Word.Contrib.Interfaces.IFields>();
		}

		/// <summary>
		/// Wrapper interface for Field which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IField WithComCleanup(this Microsoft.Office.Interop.Word.Field resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Field, Word.Contrib.Interfaces.IField>();
		}

		/// <summary>
		/// Wrapper interface for Browser which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBrowser WithComCleanup(this Microsoft.Office.Interop.Word.Browser resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Browser, Word.Contrib.Interfaces.IBrowser>();
		}

		/// <summary>
		/// Wrapper interface for Styles which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IStyles WithComCleanup(this Microsoft.Office.Interop.Word.Styles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Styles, Word.Contrib.Interfaces.IStyles>();
		}

		/// <summary>
		/// Wrapper interface for Style which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IStyle WithComCleanup(this Microsoft.Office.Interop.Word.Style resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Style, Word.Contrib.Interfaces.IStyle>();
		}

		/// <summary>
		/// Wrapper interface for Frames which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFrames WithComCleanup(this Microsoft.Office.Interop.Word.Frames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Frames, Word.Contrib.Interfaces.IFrames>();
		}

		/// <summary>
		/// Wrapper interface for Frame which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFrame WithComCleanup(this Microsoft.Office.Interop.Word.Frame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Frame, Word.Contrib.Interfaces.IFrame>();
		}

		/// <summary>
		/// Wrapper interface for FormFields which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFormFields WithComCleanup(this Microsoft.Office.Interop.Word.FormFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FormFields, Word.Contrib.Interfaces.IFormFields>();
		}

		/// <summary>
		/// Wrapper interface for FormField which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFormField WithComCleanup(this Microsoft.Office.Interop.Word.FormField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FormField, Word.Contrib.Interfaces.IFormField>();
		}

		/// <summary>
		/// Wrapper interface for TextInput which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITextInput WithComCleanup(this Microsoft.Office.Interop.Word.TextInput resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextInput, Word.Contrib.Interfaces.ITextInput>();
		}

		/// <summary>
		/// Wrapper interface for CheckBox which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICheckBox WithComCleanup(this Microsoft.Office.Interop.Word.CheckBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CheckBox, Word.Contrib.Interfaces.ICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for DropDown which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDropDown WithComCleanup(this Microsoft.Office.Interop.Word.DropDown resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DropDown, Word.Contrib.Interfaces.IDropDown>();
		}

		/// <summary>
		/// Wrapper interface for ListEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListEntries WithComCleanup(this Microsoft.Office.Interop.Word.ListEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListEntries, Word.Contrib.Interfaces.IListEntries>();
		}

		/// <summary>
		/// Wrapper interface for ListEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListEntry WithComCleanup(this Microsoft.Office.Interop.Word.ListEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListEntry, Word.Contrib.Interfaces.IListEntry>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfFigures which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITablesOfFigures WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfFigures resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfFigures, Word.Contrib.Interfaces.ITablesOfFigures>();
		}

		/// <summary>
		/// Wrapper interface for TableOfFigures which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITableOfFigures WithComCleanup(this Microsoft.Office.Interop.Word.TableOfFigures resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfFigures, Word.Contrib.Interfaces.ITableOfFigures>();
		}

		/// <summary>
		/// Wrapper interface for MailMerge which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMerge WithComCleanup(this Microsoft.Office.Interop.Word.MailMerge resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMerge, Word.Contrib.Interfaces.IMailMerge>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFields which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeFields WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeFields, Word.Contrib.Interfaces.IMailMergeFields>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeField which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeField WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeField, Word.Contrib.Interfaces.IMailMergeField>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataSource which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeDataSource WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeDataSource resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeDataSource, Word.Contrib.Interfaces.IMailMergeDataSource>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFieldNames which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeFieldNames WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeFieldNames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeFieldNames, Word.Contrib.Interfaces.IMailMergeFieldNames>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFieldName which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeFieldName WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeFieldName resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeFieldName, Word.Contrib.Interfaces.IMailMergeFieldName>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataFields which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeDataFields WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeDataFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeDataFields, Word.Contrib.Interfaces.IMailMergeDataFields>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataField which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMergeDataField WithComCleanup(this Microsoft.Office.Interop.Word.MailMergeDataField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMergeDataField, Word.Contrib.Interfaces.IMailMergeDataField>();
		}

		/// <summary>
		/// Wrapper interface for Envelope which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEnvelope WithComCleanup(this Microsoft.Office.Interop.Word.Envelope resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Envelope, Word.Contrib.Interfaces.IEnvelope>();
		}

		/// <summary>
		/// Wrapper interface for MailingLabel which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailingLabel WithComCleanup(this Microsoft.Office.Interop.Word.MailingLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailingLabel, Word.Contrib.Interfaces.IMailingLabel>();
		}

		/// <summary>
		/// Wrapper interface for CustomLabels which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICustomLabels WithComCleanup(this Microsoft.Office.Interop.Word.CustomLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomLabels, Word.Contrib.Interfaces.ICustomLabels>();
		}

		/// <summary>
		/// Wrapper interface for CustomLabel which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICustomLabel WithComCleanup(this Microsoft.Office.Interop.Word.CustomLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomLabel, Word.Contrib.Interfaces.ICustomLabel>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfContents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITablesOfContents WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfContents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfContents, Word.Contrib.Interfaces.ITablesOfContents>();
		}

		/// <summary>
		/// Wrapper interface for TableOfContents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITableOfContents WithComCleanup(this Microsoft.Office.Interop.Word.TableOfContents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfContents, Word.Contrib.Interfaces.ITableOfContents>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfAuthorities which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITablesOfAuthorities WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfAuthorities resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfAuthorities, Word.Contrib.Interfaces.ITablesOfAuthorities>();
		}

		/// <summary>
		/// Wrapper interface for TableOfAuthorities which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITableOfAuthorities WithComCleanup(this Microsoft.Office.Interop.Word.TableOfAuthorities resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfAuthorities, Word.Contrib.Interfaces.ITableOfAuthorities>();
		}

		/// <summary>
		/// Wrapper interface for Dialogs which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDialogs WithComCleanup(this Microsoft.Office.Interop.Word.Dialogs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dialogs, Word.Contrib.Interfaces.IDialogs>();
		}

		/// <summary>
		/// Wrapper interface for Dialog which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDialog WithComCleanup(this Microsoft.Office.Interop.Word.Dialog resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dialog, Word.Contrib.Interfaces.IDialog>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPageSetup WithComCleanup(this Microsoft.Office.Interop.Word.PageSetup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PageSetup, Word.Contrib.Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for LineNumbering which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILineNumbering WithComCleanup(this Microsoft.Office.Interop.Word.LineNumbering resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LineNumbering, Word.Contrib.Interfaces.ILineNumbering>();
		}

		/// <summary>
		/// Wrapper interface for TextColumns which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITextColumns WithComCleanup(this Microsoft.Office.Interop.Word.TextColumns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextColumns, Word.Contrib.Interfaces.ITextColumns>();
		}

		/// <summary>
		/// Wrapper interface for TextColumn which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITextColumn WithComCleanup(this Microsoft.Office.Interop.Word.TextColumn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextColumn, Word.Contrib.Interfaces.ITextColumn>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISelection WithComCleanup(this Microsoft.Office.Interop.Word.Selection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Selection, Word.Contrib.Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfAuthoritiesCategories which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITablesOfAuthoritiesCategories WithComCleanup(this Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories, Word.Contrib.Interfaces.ITablesOfAuthoritiesCategories>();
		}

		/// <summary>
		/// Wrapper interface for TableOfAuthoritiesCategory which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITableOfAuthoritiesCategory WithComCleanup(this Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory, Word.Contrib.Interfaces.ITableOfAuthoritiesCategory>();
		}

		/// <summary>
		/// Wrapper interface for CaptionLabels which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICaptionLabels WithComCleanup(this Microsoft.Office.Interop.Word.CaptionLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CaptionLabels, Word.Contrib.Interfaces.ICaptionLabels>();
		}

		/// <summary>
		/// Wrapper interface for CaptionLabel which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICaptionLabel WithComCleanup(this Microsoft.Office.Interop.Word.CaptionLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CaptionLabel, Word.Contrib.Interfaces.ICaptionLabel>();
		}

		/// <summary>
		/// Wrapper interface for AutoCaptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoCaptions WithComCleanup(this Microsoft.Office.Interop.Word.AutoCaptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCaptions, Word.Contrib.Interfaces.IAutoCaptions>();
		}

		/// <summary>
		/// Wrapper interface for AutoCaption which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAutoCaption WithComCleanup(this Microsoft.Office.Interop.Word.AutoCaption resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AutoCaption, Word.Contrib.Interfaces.IAutoCaption>();
		}

		/// <summary>
		/// Wrapper interface for Indexes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IIndexes WithComCleanup(this Microsoft.Office.Interop.Word.Indexes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Indexes, Word.Contrib.Interfaces.IIndexes>();
		}

		/// <summary>
		/// Wrapper interface for Index which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IIndex WithComCleanup(this Microsoft.Office.Interop.Word.Index resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Index, Word.Contrib.Interfaces.IIndex>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAddIn WithComCleanup(this Microsoft.Office.Interop.Word.AddIn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AddIn, Word.Contrib.Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAddIns WithComCleanup(this Microsoft.Office.Interop.Word.AddIns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AddIns, Word.Contrib.Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for Revisions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRevisions WithComCleanup(this Microsoft.Office.Interop.Word.Revisions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Revisions, Word.Contrib.Interfaces.IRevisions>();
		}

		/// <summary>
		/// Wrapper interface for Revision which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRevision WithComCleanup(this Microsoft.Office.Interop.Word.Revision resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Revision, Word.Contrib.Interfaces.IRevision>();
		}

		/// <summary>
		/// Wrapper interface for Task which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITask WithComCleanup(this Microsoft.Office.Interop.Word.Task resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Task, Word.Contrib.Interfaces.ITask>();
		}

		/// <summary>
		/// Wrapper interface for Tasks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITasks WithComCleanup(this Microsoft.Office.Interop.Word.Tasks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Tasks, Word.Contrib.Interfaces.ITasks>();
		}

		/// <summary>
		/// Wrapper interface for HeadersFooters which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHeadersFooters WithComCleanup(this Microsoft.Office.Interop.Word.HeadersFooters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeadersFooters, Word.Contrib.Interfaces.IHeadersFooters>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHeaderFooter WithComCleanup(this Microsoft.Office.Interop.Word.HeaderFooter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeaderFooter, Word.Contrib.Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for PageNumbers which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPageNumbers WithComCleanup(this Microsoft.Office.Interop.Word.PageNumbers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PageNumbers, Word.Contrib.Interfaces.IPageNumbers>();
		}

		/// <summary>
		/// Wrapper interface for PageNumber which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPageNumber WithComCleanup(this Microsoft.Office.Interop.Word.PageNumber resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PageNumber, Word.Contrib.Interfaces.IPageNumber>();
		}

		/// <summary>
		/// Wrapper interface for Subdocuments which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISubdocuments WithComCleanup(this Microsoft.Office.Interop.Word.Subdocuments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Subdocuments, Word.Contrib.Interfaces.ISubdocuments>();
		}

		/// <summary>
		/// Wrapper interface for Subdocument which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISubdocument WithComCleanup(this Microsoft.Office.Interop.Word.Subdocument resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Subdocument, Word.Contrib.Interfaces.ISubdocument>();
		}

		/// <summary>
		/// Wrapper interface for HeadingStyles which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHeadingStyles WithComCleanup(this Microsoft.Office.Interop.Word.HeadingStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeadingStyles, Word.Contrib.Interfaces.IHeadingStyles>();
		}

		/// <summary>
		/// Wrapper interface for HeadingStyle which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHeadingStyle WithComCleanup(this Microsoft.Office.Interop.Word.HeadingStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HeadingStyle, Word.Contrib.Interfaces.IHeadingStyle>();
		}

		/// <summary>
		/// Wrapper interface for StoryRanges which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IStoryRanges WithComCleanup(this Microsoft.Office.Interop.Word.StoryRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.StoryRanges, Word.Contrib.Interfaces.IStoryRanges>();
		}

		/// <summary>
		/// Wrapper interface for ListLevel which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListLevel WithComCleanup(this Microsoft.Office.Interop.Word.ListLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListLevel, Word.Contrib.Interfaces.IListLevel>();
		}

		/// <summary>
		/// Wrapper interface for ListLevels which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListLevels WithComCleanup(this Microsoft.Office.Interop.Word.ListLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListLevels, Word.Contrib.Interfaces.IListLevels>();
		}

		/// <summary>
		/// Wrapper interface for ListTemplate which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListTemplate WithComCleanup(this Microsoft.Office.Interop.Word.ListTemplate resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListTemplate, Word.Contrib.Interfaces.IListTemplate>();
		}

		/// <summary>
		/// Wrapper interface for ListTemplates which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListTemplates WithComCleanup(this Microsoft.Office.Interop.Word.ListTemplates resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListTemplates, Word.Contrib.Interfaces.IListTemplates>();
		}

		/// <summary>
		/// Wrapper interface for ListParagraphs which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListParagraphs WithComCleanup(this Microsoft.Office.Interop.Word.ListParagraphs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListParagraphs, Word.Contrib.Interfaces.IListParagraphs>();
		}

		/// <summary>
		/// Wrapper interface for List which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IList WithComCleanup(this Microsoft.Office.Interop.Word.List resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.List, Word.Contrib.Interfaces.IList>();
		}

		/// <summary>
		/// Wrapper interface for Lists which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILists WithComCleanup(this Microsoft.Office.Interop.Word.Lists resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Lists, Word.Contrib.Interfaces.ILists>();
		}

		/// <summary>
		/// Wrapper interface for ListGallery which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListGallery WithComCleanup(this Microsoft.Office.Interop.Word.ListGallery resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListGallery, Word.Contrib.Interfaces.IListGallery>();
		}

		/// <summary>
		/// Wrapper interface for ListGalleries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IListGalleries WithComCleanup(this Microsoft.Office.Interop.Word.ListGalleries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ListGalleries, Word.Contrib.Interfaces.IListGalleries>();
		}

		/// <summary>
		/// Wrapper interface for KeyBindings which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IKeyBindings WithComCleanup(this Microsoft.Office.Interop.Word.KeyBindings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.KeyBindings, Word.Contrib.Interfaces.IKeyBindings>();
		}

		/// <summary>
		/// Wrapper interface for KeysBoundTo which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IKeysBoundTo WithComCleanup(this Microsoft.Office.Interop.Word.KeysBoundTo resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.KeysBoundTo, Word.Contrib.Interfaces.IKeysBoundTo>();
		}

		/// <summary>
		/// Wrapper interface for KeyBinding which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IKeyBinding WithComCleanup(this Microsoft.Office.Interop.Word.KeyBinding resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.KeyBinding, Word.Contrib.Interfaces.IKeyBinding>();
		}

		/// <summary>
		/// Wrapper interface for FileConverter which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFileConverter WithComCleanup(this Microsoft.Office.Interop.Word.FileConverter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FileConverter, Word.Contrib.Interfaces.IFileConverter>();
		}

		/// <summary>
		/// Wrapper interface for FileConverters which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFileConverters WithComCleanup(this Microsoft.Office.Interop.Word.FileConverters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FileConverters, Word.Contrib.Interfaces.IFileConverters>();
		}

		/// <summary>
		/// Wrapper interface for SynonymInfo which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISynonymInfo WithComCleanup(this Microsoft.Office.Interop.Word.SynonymInfo resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SynonymInfo, Word.Contrib.Interfaces.ISynonymInfo>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHyperlinks WithComCleanup(this Microsoft.Office.Interop.Word.Hyperlinks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Hyperlinks, Word.Contrib.Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHyperlink WithComCleanup(this Microsoft.Office.Interop.Word.Hyperlink resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Hyperlink, Word.Contrib.Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShapes WithComCleanup(this Microsoft.Office.Interop.Word.Shapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Shapes, Word.Contrib.Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShapeRange WithComCleanup(this Microsoft.Office.Interop.Word.ShapeRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShapeRange, Word.Contrib.Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IGroupShapes WithComCleanup(this Microsoft.Office.Interop.Word.GroupShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.GroupShapes, Word.Contrib.Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShape WithComCleanup(this Microsoft.Office.Interop.Word.Shape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Shape, Word.Contrib.Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITextFrame WithComCleanup(this Microsoft.Office.Interop.Word.TextFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextFrame, Word.Contrib.Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for _LetterContent which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.I_LetterContent WithComCleanup(this Microsoft.Office.Interop.Word._LetterContent resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word._LetterContent, Word.Contrib.Interfaces.I_LetterContent>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IView WithComCleanup(this Microsoft.Office.Interop.Word.View resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.View, Word.Contrib.Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for Zoom which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IZoom WithComCleanup(this Microsoft.Office.Interop.Word.Zoom resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Zoom, Word.Contrib.Interfaces.IZoom>();
		}

		/// <summary>
		/// Wrapper interface for Zooms which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IZooms WithComCleanup(this Microsoft.Office.Interop.Word.Zooms resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Zooms, Word.Contrib.Interfaces.IZooms>();
		}

		/// <summary>
		/// Wrapper interface for InlineShape which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IInlineShape WithComCleanup(this Microsoft.Office.Interop.Word.InlineShape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.InlineShape, Word.Contrib.Interfaces.IInlineShape>();
		}

		/// <summary>
		/// Wrapper interface for InlineShapes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IInlineShapes WithComCleanup(this Microsoft.Office.Interop.Word.InlineShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.InlineShapes, Word.Contrib.Interfaces.IInlineShapes>();
		}

		/// <summary>
		/// Wrapper interface for SpellingSuggestions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISpellingSuggestions WithComCleanup(this Microsoft.Office.Interop.Word.SpellingSuggestions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SpellingSuggestions, Word.Contrib.Interfaces.ISpellingSuggestions>();
		}

		/// <summary>
		/// Wrapper interface for SpellingSuggestion which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISpellingSuggestion WithComCleanup(this Microsoft.Office.Interop.Word.SpellingSuggestion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SpellingSuggestion, Word.Contrib.Interfaces.ISpellingSuggestion>();
		}

		/// <summary>
		/// Wrapper interface for Dictionaries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDictionaries WithComCleanup(this Microsoft.Office.Interop.Word.Dictionaries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dictionaries, Word.Contrib.Interfaces.IDictionaries>();
		}

		/// <summary>
		/// Wrapper interface for HangulHanjaConversionDictionaries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHangulHanjaConversionDictionaries WithComCleanup(this Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries, Word.Contrib.Interfaces.IHangulHanjaConversionDictionaries>();
		}

		/// <summary>
		/// Wrapper interface for Dictionary which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDictionary WithComCleanup(this Microsoft.Office.Interop.Word.Dictionary resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Dictionary, Word.Contrib.Interfaces.IDictionary>();
		}

		/// <summary>
		/// Wrapper interface for ReadabilityStatistics which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IReadabilityStatistics WithComCleanup(this Microsoft.Office.Interop.Word.ReadabilityStatistics resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ReadabilityStatistics, Word.Contrib.Interfaces.IReadabilityStatistics>();
		}

		/// <summary>
		/// Wrapper interface for ReadabilityStatistic which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IReadabilityStatistic WithComCleanup(this Microsoft.Office.Interop.Word.ReadabilityStatistic resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ReadabilityStatistic, Word.Contrib.Interfaces.IReadabilityStatistic>();
		}

		/// <summary>
		/// Wrapper interface for Versions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IVersions WithComCleanup(this Microsoft.Office.Interop.Word.Versions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Versions, Word.Contrib.Interfaces.IVersions>();
		}

		/// <summary>
		/// Wrapper interface for Version which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IVersion WithComCleanup(this Microsoft.Office.Interop.Word.Version resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Version, Word.Contrib.Interfaces.IVersion>();
		}

		/// <summary>
		/// Wrapper interface for Options which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOptions WithComCleanup(this Microsoft.Office.Interop.Word.Options resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Options, Word.Contrib.Interfaces.IOptions>();
		}

		/// <summary>
		/// Wrapper interface for MailMessage which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailMessage WithComCleanup(this Microsoft.Office.Interop.Word.MailMessage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MailMessage, Word.Contrib.Interfaces.IMailMessage>();
		}

		/// <summary>
		/// Wrapper interface for ProofreadingErrors which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IProofreadingErrors WithComCleanup(this Microsoft.Office.Interop.Word.ProofreadingErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ProofreadingErrors, Word.Contrib.Interfaces.IProofreadingErrors>();
		}

		/// <summary>
		/// Wrapper interface for Mailer which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMailer WithComCleanup(this Microsoft.Office.Interop.Word.Mailer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Mailer, Word.Contrib.Interfaces.IMailer>();
		}

		/// <summary>
		/// Wrapper interface for WrapFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IWrapFormat WithComCleanup(this Microsoft.Office.Interop.Word.WrapFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.WrapFormat, Word.Contrib.Interfaces.IWrapFormat>();
		}

		/// <summary>
		/// Wrapper interface for HangulAndAlphabetExceptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHangulAndAlphabetExceptions WithComCleanup(this Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions, Word.Contrib.Interfaces.IHangulAndAlphabetExceptions>();
		}

		/// <summary>
		/// Wrapper interface for HangulAndAlphabetException which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHangulAndAlphabetException WithComCleanup(this Microsoft.Office.Interop.Word.HangulAndAlphabetException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HangulAndAlphabetException, Word.Contrib.Interfaces.IHangulAndAlphabetException>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAdjustments WithComCleanup(this Microsoft.Office.Interop.Word.Adjustments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Adjustments, Word.Contrib.Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICalloutFormat WithComCleanup(this Microsoft.Office.Interop.Word.CalloutFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CalloutFormat, Word.Contrib.Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IColorFormat WithComCleanup(this Microsoft.Office.Interop.Word.ColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ColorFormat, Word.Contrib.Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IConnectorFormat WithComCleanup(this Microsoft.Office.Interop.Word.ConnectorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ConnectorFormat, Word.Contrib.Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFillFormat WithComCleanup(this Microsoft.Office.Interop.Word.FillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FillFormat, Word.Contrib.Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFreeformBuilder WithComCleanup(this Microsoft.Office.Interop.Word.FreeformBuilder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FreeformBuilder, Word.Contrib.Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILineFormat WithComCleanup(this Microsoft.Office.Interop.Word.LineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LineFormat, Word.Contrib.Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPictureFormat WithComCleanup(this Microsoft.Office.Interop.Word.PictureFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PictureFormat, Word.Contrib.Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShadowFormat WithComCleanup(this Microsoft.Office.Interop.Word.ShadowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShadowFormat, Word.Contrib.Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShapeNode WithComCleanup(this Microsoft.Office.Interop.Word.ShapeNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShapeNode, Word.Contrib.Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IShapeNodes WithComCleanup(this Microsoft.Office.Interop.Word.ShapeNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ShapeNodes, Word.Contrib.Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITextEffectFormat WithComCleanup(this Microsoft.Office.Interop.Word.TextEffectFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TextEffectFormat, Word.Contrib.Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IThreeDFormat WithComCleanup(this Microsoft.Office.Interop.Word.ThreeDFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ThreeDFormat, Word.Contrib.Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents, Word.Contrib.Interfaces.IApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IGlobal WithComCleanup(this Microsoft.Office.Interop.Word.Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Global, Word.Contrib.Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents_Event, Word.Contrib.Interfaces.IApplicationEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents2_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents2_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents2_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents2_Event, Word.Contrib.Interfaces.IApplicationEvents2_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents3_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents3_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents3_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents3_Event, Word.Contrib.Interfaces.IApplicationEvents3_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents4_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents4_Event WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents4_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents4_Event, Word.Contrib.Interfaces.IApplicationEvents4_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplication WithComCleanup(this Microsoft.Office.Interop.Word.Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Application, Word.Contrib.Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDocumentEvents WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents, Word.Contrib.Interfaces.IDocumentEvents>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDocumentEvents_Event WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents_Event, Word.Contrib.Interfaces.IDocumentEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents2_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDocumentEvents2_Event WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents2_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents2_Event, Word.Contrib.Interfaces.IDocumentEvents2_Event>();
		}

		/// <summary>
		/// Wrapper interface for Document which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDocument WithComCleanup(this Microsoft.Office.Interop.Word.Document resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Document, Word.Contrib.Interfaces.IDocument>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFont WithComCleanup(this Microsoft.Office.Interop.Word.Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Font, Word.Contrib.Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IParagraphFormat WithComCleanup(this Microsoft.Office.Interop.Word.ParagraphFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ParagraphFormat, Word.Contrib.Interfaces.IParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for OCXEvents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOCXEvents WithComCleanup(this Microsoft.Office.Interop.Word.OCXEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OCXEvents, Word.Contrib.Interfaces.IOCXEvents>();
		}

		/// <summary>
		/// Wrapper interface for OCXEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOCXEvents_Event WithComCleanup(this Microsoft.Office.Interop.Word.OCXEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OCXEvents_Event, Word.Contrib.Interfaces.IOCXEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEControl which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOLEControl WithComCleanup(this Microsoft.Office.Interop.Word.OLEControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OLEControl, Word.Contrib.Interfaces.IOLEControl>();
		}

		/// <summary>
		/// Wrapper interface for LetterContent which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILetterContent WithComCleanup(this Microsoft.Office.Interop.Word.LetterContent resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LetterContent, Word.Contrib.Interfaces.ILetterContent>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IIApplicationEvents WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents, Word.Contrib.Interfaces.IIApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents2 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IIApplicationEvents2 WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents2, Word.Contrib.Interfaces.IIApplicationEvents2>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents2 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents2 WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents2, Word.Contrib.Interfaces.IApplicationEvents2>();
		}

		/// <summary>
		/// Wrapper interface for EmailAuthor which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEmailAuthor WithComCleanup(this Microsoft.Office.Interop.Word.EmailAuthor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailAuthor, Word.Contrib.Interfaces.IEmailAuthor>();
		}

		/// <summary>
		/// Wrapper interface for EmailOptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEmailOptions WithComCleanup(this Microsoft.Office.Interop.Word.EmailOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailOptions, Word.Contrib.Interfaces.IEmailOptions>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignature which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEmailSignature WithComCleanup(this Microsoft.Office.Interop.Word.EmailSignature resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailSignature, Word.Contrib.Interfaces.IEmailSignature>();
		}

		/// <summary>
		/// Wrapper interface for Email which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEmail WithComCleanup(this Microsoft.Office.Interop.Word.Email resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Email, Word.Contrib.Interfaces.IEmail>();
		}

		/// <summary>
		/// Wrapper interface for HorizontalLineFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHorizontalLineFormat WithComCleanup(this Microsoft.Office.Interop.Word.HorizontalLineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HorizontalLineFormat, Word.Contrib.Interfaces.IHorizontalLineFormat>();
		}

		/// <summary>
		/// Wrapper interface for Frameset which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFrameset WithComCleanup(this Microsoft.Office.Interop.Word.Frameset resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Frameset, Word.Contrib.Interfaces.IFrameset>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDefaultWebOptions WithComCleanup(this Microsoft.Office.Interop.Word.DefaultWebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DefaultWebOptions, Word.Contrib.Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IWebOptions WithComCleanup(this Microsoft.Office.Interop.Word.WebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.WebOptions, Word.Contrib.Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for OtherCorrectionsExceptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOtherCorrectionsExceptions WithComCleanup(this Microsoft.Office.Interop.Word.OtherCorrectionsExceptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OtherCorrectionsExceptions, Word.Contrib.Interfaces.IOtherCorrectionsExceptions>();
		}

		/// <summary>
		/// Wrapper interface for OtherCorrectionsException which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOtherCorrectionsException WithComCleanup(this Microsoft.Office.Interop.Word.OtherCorrectionsException resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OtherCorrectionsException, Word.Contrib.Interfaces.IOtherCorrectionsException>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignatureEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEmailSignatureEntries WithComCleanup(this Microsoft.Office.Interop.Word.EmailSignatureEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailSignatureEntries, Word.Contrib.Interfaces.IEmailSignatureEntries>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignatureEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEmailSignatureEntry WithComCleanup(this Microsoft.Office.Interop.Word.EmailSignatureEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EmailSignatureEntry, Word.Contrib.Interfaces.IEmailSignatureEntry>();
		}

		/// <summary>
		/// Wrapper interface for HTMLDivision which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHTMLDivision WithComCleanup(this Microsoft.Office.Interop.Word.HTMLDivision resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HTMLDivision, Word.Contrib.Interfaces.IHTMLDivision>();
		}

		/// <summary>
		/// Wrapper interface for HTMLDivisions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHTMLDivisions WithComCleanup(this Microsoft.Office.Interop.Word.HTMLDivisions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HTMLDivisions, Word.Contrib.Interfaces.IHTMLDivisions>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDiagramNode WithComCleanup(this Microsoft.Office.Interop.Word.DiagramNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DiagramNode, Word.Contrib.Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDiagramNodeChildren WithComCleanup(this Microsoft.Office.Interop.Word.DiagramNodeChildren resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DiagramNodeChildren, Word.Contrib.Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDiagramNodes WithComCleanup(this Microsoft.Office.Interop.Word.DiagramNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DiagramNodes, Word.Contrib.Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDiagram WithComCleanup(this Microsoft.Office.Interop.Word.Diagram resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Diagram, Word.Contrib.Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperty which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICustomProperty WithComCleanup(this Microsoft.Office.Interop.Word.CustomProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomProperty, Word.Contrib.Interfaces.ICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperties which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICustomProperties WithComCleanup(this Microsoft.Office.Interop.Word.CustomProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CustomProperties, Word.Contrib.Interfaces.ICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for SmartTag which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTag WithComCleanup(this Microsoft.Office.Interop.Word.SmartTag resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTag, Word.Contrib.Interfaces.ISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for SmartTags which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTags WithComCleanup(this Microsoft.Office.Interop.Word.SmartTags resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTags, Word.Contrib.Interfaces.ISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for StyleSheet which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IStyleSheet WithComCleanup(this Microsoft.Office.Interop.Word.StyleSheet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.StyleSheet, Word.Contrib.Interfaces.IStyleSheet>();
		}

		/// <summary>
		/// Wrapper interface for StyleSheets which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IStyleSheets WithComCleanup(this Microsoft.Office.Interop.Word.StyleSheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.StyleSheets, Word.Contrib.Interfaces.IStyleSheets>();
		}

		/// <summary>
		/// Wrapper interface for MappedDataField which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMappedDataField WithComCleanup(this Microsoft.Office.Interop.Word.MappedDataField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MappedDataField, Word.Contrib.Interfaces.IMappedDataField>();
		}

		/// <summary>
		/// Wrapper interface for MappedDataFields which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IMappedDataFields WithComCleanup(this Microsoft.Office.Interop.Word.MappedDataFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.MappedDataFields, Word.Contrib.Interfaces.IMappedDataFields>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICanvasShapes WithComCleanup(this Microsoft.Office.Interop.Word.CanvasShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CanvasShapes, Word.Contrib.Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITableStyle WithComCleanup(this Microsoft.Office.Interop.Word.TableStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TableStyle, Word.Contrib.Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for ConditionalStyle which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IConditionalStyle WithComCleanup(this Microsoft.Office.Interop.Word.ConditionalStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ConditionalStyle, Word.Contrib.Interfaces.IConditionalStyle>();
		}

		/// <summary>
		/// Wrapper interface for FootnoteOptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFootnoteOptions WithComCleanup(this Microsoft.Office.Interop.Word.FootnoteOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.FootnoteOptions, Word.Contrib.Interfaces.IFootnoteOptions>();
		}

		/// <summary>
		/// Wrapper interface for EndnoteOptions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEndnoteOptions WithComCleanup(this Microsoft.Office.Interop.Word.EndnoteOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.EndnoteOptions, Word.Contrib.Interfaces.IEndnoteOptions>();
		}

		/// <summary>
		/// Wrapper interface for Reviewers which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IReviewers WithComCleanup(this Microsoft.Office.Interop.Word.Reviewers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Reviewers, Word.Contrib.Interfaces.IReviewers>();
		}

		/// <summary>
		/// Wrapper interface for Reviewer which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IReviewer WithComCleanup(this Microsoft.Office.Interop.Word.Reviewer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Reviewer, Word.Contrib.Interfaces.IReviewer>();
		}

		/// <summary>
		/// Wrapper interface for TaskPane which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITaskPane WithComCleanup(this Microsoft.Office.Interop.Word.TaskPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TaskPane, Word.Contrib.Interfaces.ITaskPane>();
		}

		/// <summary>
		/// Wrapper interface for TaskPanes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITaskPanes WithComCleanup(this Microsoft.Office.Interop.Word.TaskPanes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TaskPanes, Word.Contrib.Interfaces.ITaskPanes>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents3 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IIApplicationEvents3 WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents3 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents3, Word.Contrib.Interfaces.IIApplicationEvents3>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents3 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents3 WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents3 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents3, Word.Contrib.Interfaces.IApplicationEvents3>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTagAction WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagAction, Word.Contrib.Interfaces.ISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTagActions WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagActions, Word.Contrib.Interfaces.ISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTagRecognizer WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagRecognizer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagRecognizer, Word.Contrib.Interfaces.ISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTagRecognizers WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagRecognizers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagRecognizers, Word.Contrib.Interfaces.ISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagType which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTagType WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagType resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagType, Word.Contrib.Interfaces.ISmartTagType>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagTypes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISmartTagTypes WithComCleanup(this Microsoft.Office.Interop.Word.SmartTagTypes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SmartTagTypes, Word.Contrib.Interfaces.ISmartTagTypes>();
		}

		/// <summary>
		/// Wrapper interface for Line which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILine WithComCleanup(this Microsoft.Office.Interop.Word.Line resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Line, Word.Contrib.Interfaces.ILine>();
		}

		/// <summary>
		/// Wrapper interface for Lines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILines WithComCleanup(this Microsoft.Office.Interop.Word.Lines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Lines, Word.Contrib.Interfaces.ILines>();
		}

		/// <summary>
		/// Wrapper interface for Rectangle which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRectangle WithComCleanup(this Microsoft.Office.Interop.Word.Rectangle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Rectangle, Word.Contrib.Interfaces.IRectangle>();
		}

		/// <summary>
		/// Wrapper interface for Rectangles which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IRectangles WithComCleanup(this Microsoft.Office.Interop.Word.Rectangles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Rectangles, Word.Contrib.Interfaces.IRectangles>();
		}

		/// <summary>
		/// Wrapper interface for Break which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBreak WithComCleanup(this Microsoft.Office.Interop.Word.Break resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Break, Word.Contrib.Interfaces.IBreak>();
		}

		/// <summary>
		/// Wrapper interface for Breaks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBreaks WithComCleanup(this Microsoft.Office.Interop.Word.Breaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Breaks, Word.Contrib.Interfaces.IBreaks>();
		}

		/// <summary>
		/// Wrapper interface for Page which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPage WithComCleanup(this Microsoft.Office.Interop.Word.Page resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Page, Word.Contrib.Interfaces.IPage>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPages WithComCleanup(this Microsoft.Office.Interop.Word.Pages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Pages, Word.Contrib.Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for XMLNode which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLNode WithComCleanup(this Microsoft.Office.Interop.Word.XMLNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNode, Word.Contrib.Interfaces.IXMLNode>();
		}

		/// <summary>
		/// Wrapper interface for XMLNodes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLNodes WithComCleanup(this Microsoft.Office.Interop.Word.XMLNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNodes, Word.Contrib.Interfaces.IXMLNodes>();
		}

		/// <summary>
		/// Wrapper interface for XMLSchemaReference which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLSchemaReference WithComCleanup(this Microsoft.Office.Interop.Word.XMLSchemaReference resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLSchemaReference, Word.Contrib.Interfaces.IXMLSchemaReference>();
		}

		/// <summary>
		/// Wrapper interface for XMLSchemaReferences which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLSchemaReferences WithComCleanup(this Microsoft.Office.Interop.Word.XMLSchemaReferences resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLSchemaReferences, Word.Contrib.Interfaces.IXMLSchemaReferences>();
		}

		/// <summary>
		/// Wrapper interface for XMLChildNodeSuggestion which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLChildNodeSuggestion WithComCleanup(this Microsoft.Office.Interop.Word.XMLChildNodeSuggestion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLChildNodeSuggestion, Word.Contrib.Interfaces.IXMLChildNodeSuggestion>();
		}

		/// <summary>
		/// Wrapper interface for XMLChildNodeSuggestions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLChildNodeSuggestions WithComCleanup(this Microsoft.Office.Interop.Word.XMLChildNodeSuggestions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLChildNodeSuggestions, Word.Contrib.Interfaces.IXMLChildNodeSuggestions>();
		}

		/// <summary>
		/// Wrapper interface for XMLNamespace which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLNamespace WithComCleanup(this Microsoft.Office.Interop.Word.XMLNamespace resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNamespace, Word.Contrib.Interfaces.IXMLNamespace>();
		}

		/// <summary>
		/// Wrapper interface for XMLNamespaces which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLNamespaces WithComCleanup(this Microsoft.Office.Interop.Word.XMLNamespaces resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLNamespaces, Word.Contrib.Interfaces.IXMLNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for XSLTransform which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXSLTransform WithComCleanup(this Microsoft.Office.Interop.Word.XSLTransform resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XSLTransform, Word.Contrib.Interfaces.IXSLTransform>();
		}

		/// <summary>
		/// Wrapper interface for XSLTransforms which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXSLTransforms WithComCleanup(this Microsoft.Office.Interop.Word.XSLTransforms resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XSLTransforms, Word.Contrib.Interfaces.IXSLTransforms>();
		}

		/// <summary>
		/// Wrapper interface for Editors which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEditors WithComCleanup(this Microsoft.Office.Interop.Word.Editors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Editors, Word.Contrib.Interfaces.IEditors>();
		}

		/// <summary>
		/// Wrapper interface for Editor which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IEditor WithComCleanup(this Microsoft.Office.Interop.Word.Editor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Editor, Word.Contrib.Interfaces.IEditor>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents4 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IIApplicationEvents4 WithComCleanup(this Microsoft.Office.Interop.Word.IApplicationEvents4 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.IApplicationEvents4, Word.Contrib.Interfaces.IIApplicationEvents4>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents4 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IApplicationEvents4 WithComCleanup(this Microsoft.Office.Interop.Word.ApplicationEvents4 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ApplicationEvents4, Word.Contrib.Interfaces.IApplicationEvents4>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents2 which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDocumentEvents2 WithComCleanup(this Microsoft.Office.Interop.Word.DocumentEvents2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DocumentEvents2, Word.Contrib.Interfaces.IDocumentEvents2>();
		}

		/// <summary>
		/// Wrapper interface for Source which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISource WithComCleanup(this Microsoft.Office.Interop.Word.Source resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Source, Word.Contrib.Interfaces.ISource>();
		}

		/// <summary>
		/// Wrapper interface for Sources which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISources WithComCleanup(this Microsoft.Office.Interop.Word.Sources resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Sources, Word.Contrib.Interfaces.ISources>();
		}

		/// <summary>
		/// Wrapper interface for Bibliography which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBibliography WithComCleanup(this Microsoft.Office.Interop.Word.Bibliography resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Bibliography, Word.Contrib.Interfaces.IBibliography>();
		}

		/// <summary>
		/// Wrapper interface for OMaths which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMaths WithComCleanup(this Microsoft.Office.Interop.Word.OMaths resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMaths, Word.Contrib.Interfaces.IOMaths>();
		}

		/// <summary>
		/// Wrapper interface for OMath which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMath WithComCleanup(this Microsoft.Office.Interop.Word.OMath resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMath, Word.Contrib.Interfaces.IOMath>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunctions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathFunctions WithComCleanup(this Microsoft.Office.Interop.Word.OMathFunctions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFunctions, Word.Contrib.Interfaces.IOMathFunctions>();
		}

		/// <summary>
		/// Wrapper interface for OMathArgs which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathArgs WithComCleanup(this Microsoft.Office.Interop.Word.OMathArgs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathArgs, Word.Contrib.Interfaces.IOMathArgs>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunction which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathFunction WithComCleanup(this Microsoft.Office.Interop.Word.OMathFunction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFunction, Word.Contrib.Interfaces.IOMathFunction>();
		}

		/// <summary>
		/// Wrapper interface for OMathAcc which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathAcc WithComCleanup(this Microsoft.Office.Interop.Word.OMathAcc resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAcc, Word.Contrib.Interfaces.IOMathAcc>();
		}

		/// <summary>
		/// Wrapper interface for OMathBar which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathBar WithComCleanup(this Microsoft.Office.Interop.Word.OMathBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBar, Word.Contrib.Interfaces.IOMathBar>();
		}

		/// <summary>
		/// Wrapper interface for OMathBox which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathBox WithComCleanup(this Microsoft.Office.Interop.Word.OMathBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBox, Word.Contrib.Interfaces.IOMathBox>();
		}

		/// <summary>
		/// Wrapper interface for OMathBorderBox which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathBorderBox WithComCleanup(this Microsoft.Office.Interop.Word.OMathBorderBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBorderBox, Word.Contrib.Interfaces.IOMathBorderBox>();
		}

		/// <summary>
		/// Wrapper interface for OMathDelim which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathDelim WithComCleanup(this Microsoft.Office.Interop.Word.OMathDelim resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathDelim, Word.Contrib.Interfaces.IOMathDelim>();
		}

		/// <summary>
		/// Wrapper interface for OMathEqArray which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathEqArray WithComCleanup(this Microsoft.Office.Interop.Word.OMathEqArray resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathEqArray, Word.Contrib.Interfaces.IOMathEqArray>();
		}

		/// <summary>
		/// Wrapper interface for OMathFrac which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathFrac WithComCleanup(this Microsoft.Office.Interop.Word.OMathFrac resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFrac, Word.Contrib.Interfaces.IOMathFrac>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunc which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathFunc WithComCleanup(this Microsoft.Office.Interop.Word.OMathFunc resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathFunc, Word.Contrib.Interfaces.IOMathFunc>();
		}

		/// <summary>
		/// Wrapper interface for OMathGroupChar which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathGroupChar WithComCleanup(this Microsoft.Office.Interop.Word.OMathGroupChar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathGroupChar, Word.Contrib.Interfaces.IOMathGroupChar>();
		}

		/// <summary>
		/// Wrapper interface for OMathLimLow which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathLimLow WithComCleanup(this Microsoft.Office.Interop.Word.OMathLimLow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathLimLow, Word.Contrib.Interfaces.IOMathLimLow>();
		}

		/// <summary>
		/// Wrapper interface for OMathLimUpp which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathLimUpp WithComCleanup(this Microsoft.Office.Interop.Word.OMathLimUpp resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathLimUpp, Word.Contrib.Interfaces.IOMathLimUpp>();
		}

		/// <summary>
		/// Wrapper interface for OMathMat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathMat WithComCleanup(this Microsoft.Office.Interop.Word.OMathMat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMat, Word.Contrib.Interfaces.IOMathMat>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatRows which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathMatRows WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatRows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatRows, Word.Contrib.Interfaces.IOMathMatRows>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatCols which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathMatCols WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatCols resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatCols, Word.Contrib.Interfaces.IOMathMatCols>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatRow which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathMatRow WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatRow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatRow, Word.Contrib.Interfaces.IOMathMatRow>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatCol which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathMatCol WithComCleanup(this Microsoft.Office.Interop.Word.OMathMatCol resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathMatCol, Word.Contrib.Interfaces.IOMathMatCol>();
		}

		/// <summary>
		/// Wrapper interface for OMathNary which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathNary WithComCleanup(this Microsoft.Office.Interop.Word.OMathNary resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathNary, Word.Contrib.Interfaces.IOMathNary>();
		}

		/// <summary>
		/// Wrapper interface for OMathPhantom which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathPhantom WithComCleanup(this Microsoft.Office.Interop.Word.OMathPhantom resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathPhantom, Word.Contrib.Interfaces.IOMathPhantom>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrPre which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathScrPre WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrPre resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrPre, Word.Contrib.Interfaces.IOMathScrPre>();
		}

		/// <summary>
		/// Wrapper interface for OMathRad which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathRad WithComCleanup(this Microsoft.Office.Interop.Word.OMathRad resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathRad, Word.Contrib.Interfaces.IOMathRad>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSub which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathScrSub WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrSub resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrSub, Word.Contrib.Interfaces.IOMathScrSub>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSubSup which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathScrSubSup WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrSubSup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrSubSup, Word.Contrib.Interfaces.IOMathScrSubSup>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSup which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathScrSup WithComCleanup(this Microsoft.Office.Interop.Word.OMathScrSup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathScrSup, Word.Contrib.Interfaces.IOMathScrSup>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathAutoCorrect WithComCleanup(this Microsoft.Office.Interop.Word.OMathAutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAutoCorrect, Word.Contrib.Interfaces.IOMathAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrectEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathAutoCorrectEntries WithComCleanup(this Microsoft.Office.Interop.Word.OMathAutoCorrectEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAutoCorrectEntries, Word.Contrib.Interfaces.IOMathAutoCorrectEntries>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrectEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathAutoCorrectEntry WithComCleanup(this Microsoft.Office.Interop.Word.OMathAutoCorrectEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathAutoCorrectEntry, Word.Contrib.Interfaces.IOMathAutoCorrectEntry>();
		}

		/// <summary>
		/// Wrapper interface for OMathRecognizedFunctions which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathRecognizedFunctions WithComCleanup(this Microsoft.Office.Interop.Word.OMathRecognizedFunctions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathRecognizedFunctions, Word.Contrib.Interfaces.IOMathRecognizedFunctions>();
		}

		/// <summary>
		/// Wrapper interface for OMathRecognizedFunction which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathRecognizedFunction WithComCleanup(this Microsoft.Office.Interop.Word.OMathRecognizedFunction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathRecognizedFunction, Word.Contrib.Interfaces.IOMathRecognizedFunction>();
		}

		/// <summary>
		/// Wrapper interface for ContentControls which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IContentControls WithComCleanup(this Microsoft.Office.Interop.Word.ContentControls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControls, Word.Contrib.Interfaces.IContentControls>();
		}

		/// <summary>
		/// Wrapper interface for ContentControl which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IContentControl WithComCleanup(this Microsoft.Office.Interop.Word.ContentControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControl, Word.Contrib.Interfaces.IContentControl>();
		}

		/// <summary>
		/// Wrapper interface for XMLMapping which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IXMLMapping WithComCleanup(this Microsoft.Office.Interop.Word.XMLMapping resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.XMLMapping, Word.Contrib.Interfaces.IXMLMapping>();
		}

		/// <summary>
		/// Wrapper interface for ContentControlListEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IContentControlListEntries WithComCleanup(this Microsoft.Office.Interop.Word.ContentControlListEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControlListEntries, Word.Contrib.Interfaces.IContentControlListEntries>();
		}

		/// <summary>
		/// Wrapper interface for ContentControlListEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IContentControlListEntry WithComCleanup(this Microsoft.Office.Interop.Word.ContentControlListEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ContentControlListEntry, Word.Contrib.Interfaces.IContentControlListEntry>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockTypes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBuildingBlockTypes WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlockTypes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlockTypes, Word.Contrib.Interfaces.IBuildingBlockTypes>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockType which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBuildingBlockType WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlockType resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlockType, Word.Contrib.Interfaces.IBuildingBlockType>();
		}

		/// <summary>
		/// Wrapper interface for Categories which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICategories WithComCleanup(this Microsoft.Office.Interop.Word.Categories resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Categories, Word.Contrib.Interfaces.ICategories>();
		}

		/// <summary>
		/// Wrapper interface for Category which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICategory WithComCleanup(this Microsoft.Office.Interop.Word.Category resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Category, Word.Contrib.Interfaces.ICategory>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlocks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBuildingBlocks WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlocks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlocks, Word.Contrib.Interfaces.IBuildingBlocks>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlock which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBuildingBlock WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlock resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlock, Word.Contrib.Interfaces.IBuildingBlock>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IBuildingBlockEntries WithComCleanup(this Microsoft.Office.Interop.Word.BuildingBlockEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.BuildingBlockEntries, Word.Contrib.Interfaces.IBuildingBlockEntries>();
		}

		/// <summary>
		/// Wrapper interface for OMathBreaks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathBreaks WithComCleanup(this Microsoft.Office.Interop.Word.OMathBreaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBreaks, Word.Contrib.Interfaces.IOMathBreaks>();
		}

		/// <summary>
		/// Wrapper interface for OMathBreak which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IOMathBreak WithComCleanup(this Microsoft.Office.Interop.Word.OMathBreak resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.OMathBreak, Word.Contrib.Interfaces.IOMathBreak>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IResearch WithComCleanup(this Microsoft.Office.Interop.Word.Research resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Research, Word.Contrib.Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISoftEdgeFormat WithComCleanup(this Microsoft.Office.Interop.Word.SoftEdgeFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SoftEdgeFormat, Word.Contrib.Interfaces.ISoftEdgeFormat>();
		}

		/// <summary>
		/// Wrapper interface for GlowFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IGlowFormat WithComCleanup(this Microsoft.Office.Interop.Word.GlowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.GlowFormat, Word.Contrib.Interfaces.IGlowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IReflectionFormat WithComCleanup(this Microsoft.Office.Interop.Word.ReflectionFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ReflectionFormat, Word.Contrib.Interfaces.IReflectionFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartData which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartData WithComCleanup(this Microsoft.Office.Interop.Word.ChartData resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartData, Word.Contrib.Interfaces.IChartData>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChart WithComCleanup(this Microsoft.Office.Interop.Word.Chart resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Chart, Word.Contrib.Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICorners WithComCleanup(this Microsoft.Office.Interop.Word.Corners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Corners, Word.Contrib.Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILegend WithComCleanup(this Microsoft.Office.Interop.Word.Legend resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Legend, Word.Contrib.Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for ChartBorder which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartBorder WithComCleanup(this Microsoft.Office.Interop.Word.ChartBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartBorder, Word.Contrib.Interfaces.IChartBorder>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IWalls WithComCleanup(this Microsoft.Office.Interop.Word.Walls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Walls, Word.Contrib.Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IFloor WithComCleanup(this Microsoft.Office.Interop.Word.Floor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Floor, Word.Contrib.Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPlotArea WithComCleanup(this Microsoft.Office.Interop.Word.PlotArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.PlotArea, Word.Contrib.Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartArea WithComCleanup(this Microsoft.Office.Interop.Word.ChartArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartArea, Word.Contrib.Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISeriesLines WithComCleanup(this Microsoft.Office.Interop.Word.SeriesLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SeriesLines, Word.Contrib.Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILeaderLines WithComCleanup(this Microsoft.Office.Interop.Word.LeaderLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LeaderLines, Word.Contrib.Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IGridlines WithComCleanup(this Microsoft.Office.Interop.Word.Gridlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Gridlines, Word.Contrib.Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IUpBars WithComCleanup(this Microsoft.Office.Interop.Word.UpBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.UpBars, Word.Contrib.Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDownBars WithComCleanup(this Microsoft.Office.Interop.Word.DownBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DownBars, Word.Contrib.Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IInterior WithComCleanup(this Microsoft.Office.Interop.Word.Interior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Interior, Word.Contrib.Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartFillFormat WithComCleanup(this Microsoft.Office.Interop.Word.ChartFillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartFillFormat, Word.Contrib.Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILegendEntries WithComCleanup(this Microsoft.Office.Interop.Word.LegendEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LegendEntries, Word.Contrib.Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartFont WithComCleanup(this Microsoft.Office.Interop.Word.ChartFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartFont, Word.Contrib.Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartColorFormat WithComCleanup(this Microsoft.Office.Interop.Word.ChartColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartColorFormat, Word.Contrib.Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILegendEntry WithComCleanup(this Microsoft.Office.Interop.Word.LegendEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LegendEntry, Word.Contrib.Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ILegendKey WithComCleanup(this Microsoft.Office.Interop.Word.LegendKey resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.LegendKey, Word.Contrib.Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISeriesCollection WithComCleanup(this Microsoft.Office.Interop.Word.SeriesCollection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.SeriesCollection, Word.Contrib.Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ISeries WithComCleanup(this Microsoft.Office.Interop.Word.Series resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Series, Word.Contrib.Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IErrorBars WithComCleanup(this Microsoft.Office.Interop.Word.ErrorBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ErrorBars, Word.Contrib.Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITrendline WithComCleanup(this Microsoft.Office.Interop.Word.Trendline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Trendline, Word.Contrib.Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITrendlines WithComCleanup(this Microsoft.Office.Interop.Word.Trendlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Trendlines, Word.Contrib.Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDataLabels WithComCleanup(this Microsoft.Office.Interop.Word.DataLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DataLabels, Word.Contrib.Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDataLabel WithComCleanup(this Microsoft.Office.Interop.Word.DataLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DataLabel, Word.Contrib.Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPoints WithComCleanup(this Microsoft.Office.Interop.Word.Points resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Points, Word.Contrib.Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IPoint WithComCleanup(this Microsoft.Office.Interop.Word.Point resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Point, Word.Contrib.Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAxes WithComCleanup(this Microsoft.Office.Interop.Word.Axes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Axes, Word.Contrib.Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAxis WithComCleanup(this Microsoft.Office.Interop.Word.Axis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Axis, Word.Contrib.Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDataTable WithComCleanup(this Microsoft.Office.Interop.Word.DataTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DataTable, Word.Contrib.Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartTitle WithComCleanup(this Microsoft.Office.Interop.Word.ChartTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartTitle, Word.Contrib.Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IAxisTitle WithComCleanup(this Microsoft.Office.Interop.Word.AxisTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.AxisTitle, Word.Contrib.Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDisplayUnitLabel WithComCleanup(this Microsoft.Office.Interop.Word.DisplayUnitLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DisplayUnitLabel, Word.Contrib.Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ITickLabels WithComCleanup(this Microsoft.Office.Interop.Word.TickLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.TickLabels, Word.Contrib.Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IDropLines WithComCleanup(this Microsoft.Office.Interop.Word.DropLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.DropLines, Word.Contrib.Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IHiLoLines WithComCleanup(this Microsoft.Office.Interop.Word.HiLoLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.HiLoLines, Word.Contrib.Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartGroup WithComCleanup(this Microsoft.Office.Interop.Word.ChartGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartGroup, Word.Contrib.Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartGroups WithComCleanup(this Microsoft.Office.Interop.Word.ChartGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartGroups, Word.Contrib.Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for ChartCharacters which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartCharacters WithComCleanup(this Microsoft.Office.Interop.Word.ChartCharacters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartCharacters, Word.Contrib.Interfaces.IChartCharacters>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IChartFormat WithComCleanup(this Microsoft.Office.Interop.Word.ChartFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ChartFormat, Word.Contrib.Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for UndoRecord which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IUndoRecord WithComCleanup(this Microsoft.Office.Interop.Word.UndoRecord resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.UndoRecord, Word.Contrib.Interfaces.IUndoRecord>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthLock which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthLock WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthLock resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthLock, Word.Contrib.Interfaces.ICoAuthLock>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthLocks which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthLocks WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthLocks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthLocks, Word.Contrib.Interfaces.ICoAuthLocks>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthUpdate which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthUpdate WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthUpdate resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthUpdate, Word.Contrib.Interfaces.ICoAuthUpdate>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthUpdates which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthUpdates WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthUpdates resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthUpdates, Word.Contrib.Interfaces.ICoAuthUpdates>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthor which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthor WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthor, Word.Contrib.Interfaces.ICoAuthor>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthors which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthors WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthors, Word.Contrib.Interfaces.ICoAuthors>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthoring which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.ICoAuthoring WithComCleanup(this Microsoft.Office.Interop.Word.CoAuthoring resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.CoAuthoring, Word.Contrib.Interfaces.ICoAuthoring>();
		}

		/// <summary>
		/// Wrapper interface for Conflicts which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IConflicts WithComCleanup(this Microsoft.Office.Interop.Word.Conflicts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Conflicts, Word.Contrib.Interfaces.IConflicts>();
		}

		/// <summary>
		/// Wrapper interface for Conflict which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IConflict WithComCleanup(this Microsoft.Office.Interop.Word.Conflict resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.Conflict, Word.Contrib.Interfaces.IConflict>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IProtectedViewWindows WithComCleanup(this Microsoft.Office.Interop.Word.ProtectedViewWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ProtectedViewWindows, Word.Contrib.Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Word.Contrib.Interfaces.IProtectedViewWindow WithComCleanup(this Microsoft.Office.Interop.Word.ProtectedViewWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Word.ProtectedViewWindow, Word.Contrib.Interfaces.IProtectedViewWindow>();
		}

	}
}