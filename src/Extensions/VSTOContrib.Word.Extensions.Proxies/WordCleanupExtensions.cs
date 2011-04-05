using VSTOContrib.Extensions.Proxies;

//Microsoft.Office.Interop.Word, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.Word.Extensions.Proxies
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Application WithComCleanupProxy(this Microsoft.Office.Interop.Word._Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._Application,Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Global WithComCleanupProxy(this Microsoft.Office.Interop.Word._Global resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._Global,Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for FontNames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFontNames WithComCleanupProxy(this Microsoft.Office.Interop.Word.FontNames resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FontNames,Interfaces.IFontNames>();
		}

		/// <summary>
		/// Wrapper interface for Languages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILanguages WithComCleanupProxy(this Microsoft.Office.Interop.Word.Languages resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Languages,Interfaces.ILanguages>();
		}

		/// <summary>
		/// Wrapper interface for Language which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILanguage WithComCleanupProxy(this Microsoft.Office.Interop.Word.Language resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Language,Interfaces.ILanguage>();
		}

		/// <summary>
		/// Wrapper interface for Documents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocuments WithComCleanupProxy(this Microsoft.Office.Interop.Word.Documents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Documents,Interfaces.IDocuments>();
		}

		/// <summary>
		/// Wrapper interface for _Document which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Document WithComCleanupProxy(this Microsoft.Office.Interop.Word._Document resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._Document,Interfaces.I_Document>();
		}

		/// <summary>
		/// Wrapper interface for Template which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITemplate WithComCleanupProxy(this Microsoft.Office.Interop.Word.Template resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Template,Interfaces.ITemplate>();
		}

		/// <summary>
		/// Wrapper interface for Templates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITemplates WithComCleanupProxy(this Microsoft.Office.Interop.Word.Templates resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Templates,Interfaces.ITemplates>();
		}

		/// <summary>
		/// Wrapper interface for RoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRoutingSlip WithComCleanupProxy(this Microsoft.Office.Interop.Word.RoutingSlip resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.RoutingSlip,Interfaces.IRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for Bookmark which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBookmark WithComCleanupProxy(this Microsoft.Office.Interop.Word.Bookmark resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Bookmark,Interfaces.IBookmark>();
		}

		/// <summary>
		/// Wrapper interface for Bookmarks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBookmarks WithComCleanupProxy(this Microsoft.Office.Interop.Word.Bookmarks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Bookmarks,Interfaces.IBookmarks>();
		}

		/// <summary>
		/// Wrapper interface for Variable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVariable WithComCleanupProxy(this Microsoft.Office.Interop.Word.Variable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Variable,Interfaces.IVariable>();
		}

		/// <summary>
		/// Wrapper interface for Variables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVariables WithComCleanupProxy(this Microsoft.Office.Interop.Word.Variables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Variables,Interfaces.IVariables>();
		}

		/// <summary>
		/// Wrapper interface for RecentFile which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecentFile WithComCleanupProxy(this Microsoft.Office.Interop.Word.RecentFile resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.RecentFile,Interfaces.IRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for RecentFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecentFiles WithComCleanupProxy(this Microsoft.Office.Interop.Word.RecentFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.RecentFiles,Interfaces.IRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for Window which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWindow WithComCleanupProxy(this Microsoft.Office.Interop.Word.Window resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Window,Interfaces.IWindow>();
		}

		/// <summary>
		/// Wrapper interface for Windows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWindows WithComCleanupProxy(this Microsoft.Office.Interop.Word.Windows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Windows,Interfaces.IWindows>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPane WithComCleanupProxy(this Microsoft.Office.Interop.Word.Pane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Pane,Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPanes WithComCleanupProxy(this Microsoft.Office.Interop.Word.Panes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Panes,Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Range which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRange WithComCleanupProxy(this Microsoft.Office.Interop.Word.Range resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Range,Interfaces.IRange>();
		}

		/// <summary>
		/// Wrapper interface for ListFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListFormat,Interfaces.IListFormat>();
		}

		/// <summary>
		/// Wrapper interface for Find which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFind WithComCleanupProxy(this Microsoft.Office.Interop.Word.Find resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Find,Interfaces.IFind>();
		}

		/// <summary>
		/// Wrapper interface for Replacement which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReplacement WithComCleanupProxy(this Microsoft.Office.Interop.Word.Replacement resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Replacement,Interfaces.IReplacement>();
		}

		/// <summary>
		/// Wrapper interface for Characters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICharacters WithComCleanupProxy(this Microsoft.Office.Interop.Word.Characters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Characters,Interfaces.ICharacters>();
		}

		/// <summary>
		/// Wrapper interface for Words which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWords WithComCleanupProxy(this Microsoft.Office.Interop.Word.Words resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Words,Interfaces.IWords>();
		}

		/// <summary>
		/// Wrapper interface for Sentences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISentences WithComCleanupProxy(this Microsoft.Office.Interop.Word.Sentences resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Sentences,Interfaces.ISentences>();
		}

		/// <summary>
		/// Wrapper interface for Sections which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISections WithComCleanupProxy(this Microsoft.Office.Interop.Word.Sections resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Sections,Interfaces.ISections>();
		}

		/// <summary>
		/// Wrapper interface for Section which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISection WithComCleanupProxy(this Microsoft.Office.Interop.Word.Section resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Section,Interfaces.ISection>();
		}

		/// <summary>
		/// Wrapper interface for Paragraphs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphs WithComCleanupProxy(this Microsoft.Office.Interop.Word.Paragraphs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Paragraphs,Interfaces.IParagraphs>();
		}

		/// <summary>
		/// Wrapper interface for Paragraph which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraph WithComCleanupProxy(this Microsoft.Office.Interop.Word.Paragraph resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Paragraph,Interfaces.IParagraph>();
		}

		/// <summary>
		/// Wrapper interface for DropCap which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropCap WithComCleanupProxy(this Microsoft.Office.Interop.Word.DropCap resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DropCap,Interfaces.IDropCap>();
		}

		/// <summary>
		/// Wrapper interface for TabStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStops WithComCleanupProxy(this Microsoft.Office.Interop.Word.TabStops resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TabStops,Interfaces.ITabStops>();
		}

		/// <summary>
		/// Wrapper interface for TabStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStop WithComCleanupProxy(this Microsoft.Office.Interop.Word.TabStop resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TabStop,Interfaces.ITabStop>();
		}

		/// <summary>
		/// Wrapper interface for _ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_ParagraphFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word._ParagraphFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._ParagraphFormat,Interfaces.I_ParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for _Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Font WithComCleanupProxy(this Microsoft.Office.Interop.Word._Font resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._Font,Interfaces.I_Font>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITable WithComCleanupProxy(this Microsoft.Office.Interop.Word.Table resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Table,Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRow WithComCleanupProxy(this Microsoft.Office.Interop.Word.Row resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Row,Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumn WithComCleanupProxy(this Microsoft.Office.Interop.Word.Column resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Column,Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for Cell which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICell WithComCleanupProxy(this Microsoft.Office.Interop.Word.Cell resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Cell,Interfaces.ICell>();
		}

		/// <summary>
		/// Wrapper interface for Tables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITables WithComCleanupProxy(this Microsoft.Office.Interop.Word.Tables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Tables,Interfaces.ITables>();
		}

		/// <summary>
		/// Wrapper interface for Rows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRows WithComCleanupProxy(this Microsoft.Office.Interop.Word.Rows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Rows,Interfaces.IRows>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumns WithComCleanupProxy(this Microsoft.Office.Interop.Word.Columns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Columns,Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for Cells which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICells WithComCleanupProxy(this Microsoft.Office.Interop.Word.Cells resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Cells,Interfaces.ICells>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrect WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoCorrect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoCorrect,Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrectEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrectEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoCorrectEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoCorrectEntries,Interfaces.IAutoCorrectEntries>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrectEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrectEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoCorrectEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoCorrectEntry,Interfaces.IAutoCorrectEntry>();
		}

		/// <summary>
		/// Wrapper interface for FirstLetterExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFirstLetterExceptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.FirstLetterExceptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FirstLetterExceptions,Interfaces.IFirstLetterExceptions>();
		}

		/// <summary>
		/// Wrapper interface for FirstLetterException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFirstLetterException WithComCleanupProxy(this Microsoft.Office.Interop.Word.FirstLetterException resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FirstLetterException,Interfaces.IFirstLetterException>();
		}

		/// <summary>
		/// Wrapper interface for TwoInitialCapsExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITwoInitialCapsExceptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.TwoInitialCapsExceptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TwoInitialCapsExceptions,Interfaces.ITwoInitialCapsExceptions>();
		}

		/// <summary>
		/// Wrapper interface for TwoInitialCapsException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITwoInitialCapsException WithComCleanupProxy(this Microsoft.Office.Interop.Word.TwoInitialCapsException resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TwoInitialCapsException,Interfaces.ITwoInitialCapsException>();
		}

		/// <summary>
		/// Wrapper interface for Footnotes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFootnotes WithComCleanupProxy(this Microsoft.Office.Interop.Word.Footnotes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Footnotes,Interfaces.IFootnotes>();
		}

		/// <summary>
		/// Wrapper interface for Endnotes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEndnotes WithComCleanupProxy(this Microsoft.Office.Interop.Word.Endnotes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Endnotes,Interfaces.IEndnotes>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComments WithComCleanupProxy(this Microsoft.Office.Interop.Word.Comments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Comments,Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Footnote which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFootnote WithComCleanupProxy(this Microsoft.Office.Interop.Word.Footnote resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Footnote,Interfaces.IFootnote>();
		}

		/// <summary>
		/// Wrapper interface for Endnote which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEndnote WithComCleanupProxy(this Microsoft.Office.Interop.Word.Endnote resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Endnote,Interfaces.IEndnote>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComment WithComCleanupProxy(this Microsoft.Office.Interop.Word.Comment resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Comment,Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorders WithComCleanupProxy(this Microsoft.Office.Interop.Word.Borders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Borders,Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for Border which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorder WithComCleanupProxy(this Microsoft.Office.Interop.Word.Border resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Border,Interfaces.IBorder>();
		}

		/// <summary>
		/// Wrapper interface for Shading which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShading WithComCleanupProxy(this Microsoft.Office.Interop.Word.Shading resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Shading,Interfaces.IShading>();
		}

		/// <summary>
		/// Wrapper interface for TextRetrievalMode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextRetrievalMode WithComCleanupProxy(this Microsoft.Office.Interop.Word.TextRetrievalMode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TextRetrievalMode,Interfaces.ITextRetrievalMode>();
		}

		/// <summary>
		/// Wrapper interface for AutoTextEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoTextEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoTextEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoTextEntries,Interfaces.IAutoTextEntries>();
		}

		/// <summary>
		/// Wrapper interface for AutoTextEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoTextEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoTextEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoTextEntry,Interfaces.IAutoTextEntry>();
		}

		/// <summary>
		/// Wrapper interface for System which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISystem WithComCleanupProxy(this Microsoft.Office.Interop.Word.System resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.System,Interfaces.ISystem>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.OLEFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OLEFormat,Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinkFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.LinkFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LinkFormat,Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for _OLEControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OLEControl WithComCleanupProxy(this Microsoft.Office.Interop.Word._OLEControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._OLEControl,Interfaces.I_OLEControl>();
		}

		/// <summary>
		/// Wrapper interface for Fields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFields WithComCleanupProxy(this Microsoft.Office.Interop.Word.Fields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Fields,Interfaces.IFields>();
		}

		/// <summary>
		/// Wrapper interface for Field which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IField WithComCleanupProxy(this Microsoft.Office.Interop.Word.Field resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Field,Interfaces.IField>();
		}

		/// <summary>
		/// Wrapper interface for Browser which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBrowser WithComCleanupProxy(this Microsoft.Office.Interop.Word.Browser resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Browser,Interfaces.IBrowser>();
		}

		/// <summary>
		/// Wrapper interface for Styles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyles WithComCleanupProxy(this Microsoft.Office.Interop.Word.Styles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Styles,Interfaces.IStyles>();
		}

		/// <summary>
		/// Wrapper interface for Style which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyle WithComCleanupProxy(this Microsoft.Office.Interop.Word.Style resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Style,Interfaces.IStyle>();
		}

		/// <summary>
		/// Wrapper interface for Frames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFrames WithComCleanupProxy(this Microsoft.Office.Interop.Word.Frames resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Frames,Interfaces.IFrames>();
		}

		/// <summary>
		/// Wrapper interface for Frame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFrame WithComCleanupProxy(this Microsoft.Office.Interop.Word.Frame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Frame,Interfaces.IFrame>();
		}

		/// <summary>
		/// Wrapper interface for FormFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormFields WithComCleanupProxy(this Microsoft.Office.Interop.Word.FormFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FormFields,Interfaces.IFormFields>();
		}

		/// <summary>
		/// Wrapper interface for FormField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormField WithComCleanupProxy(this Microsoft.Office.Interop.Word.FormField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FormField,Interfaces.IFormField>();
		}

		/// <summary>
		/// Wrapper interface for TextInput which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextInput WithComCleanupProxy(this Microsoft.Office.Interop.Word.TextInput resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TextInput,Interfaces.ITextInput>();
		}

		/// <summary>
		/// Wrapper interface for CheckBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICheckBox WithComCleanupProxy(this Microsoft.Office.Interop.Word.CheckBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CheckBox,Interfaces.ICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for DropDown which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropDown WithComCleanupProxy(this Microsoft.Office.Interop.Word.DropDown resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DropDown,Interfaces.IDropDown>();
		}

		/// <summary>
		/// Wrapper interface for ListEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListEntries,Interfaces.IListEntries>();
		}

		/// <summary>
		/// Wrapper interface for ListEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListEntry,Interfaces.IListEntry>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfFigures which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfFigures WithComCleanupProxy(this Microsoft.Office.Interop.Word.TablesOfFigures resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TablesOfFigures,Interfaces.ITablesOfFigures>();
		}

		/// <summary>
		/// Wrapper interface for TableOfFigures which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfFigures WithComCleanupProxy(this Microsoft.Office.Interop.Word.TableOfFigures resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TableOfFigures,Interfaces.ITableOfFigures>();
		}

		/// <summary>
		/// Wrapper interface for MailMerge which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMerge WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMerge resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMerge,Interfaces.IMailMerge>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeFields WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeFields,Interfaces.IMailMergeFields>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeField WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeField,Interfaces.IMailMergeField>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataSource which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeDataSource WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeDataSource resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeDataSource,Interfaces.IMailMergeDataSource>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFieldNames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeFieldNames WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeFieldNames resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeFieldNames,Interfaces.IMailMergeFieldNames>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeFieldName which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeFieldName WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeFieldName resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeFieldName,Interfaces.IMailMergeFieldName>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeDataFields WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeDataFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeDataFields,Interfaces.IMailMergeDataFields>();
		}

		/// <summary>
		/// Wrapper interface for MailMergeDataField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMergeDataField WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMergeDataField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMergeDataField,Interfaces.IMailMergeDataField>();
		}

		/// <summary>
		/// Wrapper interface for Envelope which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEnvelope WithComCleanupProxy(this Microsoft.Office.Interop.Word.Envelope resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Envelope,Interfaces.IEnvelope>();
		}

		/// <summary>
		/// Wrapper interface for MailingLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailingLabel WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailingLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailingLabel,Interfaces.IMailingLabel>();
		}

		/// <summary>
		/// Wrapper interface for CustomLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLabels WithComCleanupProxy(this Microsoft.Office.Interop.Word.CustomLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CustomLabels,Interfaces.ICustomLabels>();
		}

		/// <summary>
		/// Wrapper interface for CustomLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLabel WithComCleanupProxy(this Microsoft.Office.Interop.Word.CustomLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CustomLabel,Interfaces.ICustomLabel>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfContents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfContents WithComCleanupProxy(this Microsoft.Office.Interop.Word.TablesOfContents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TablesOfContents,Interfaces.ITablesOfContents>();
		}

		/// <summary>
		/// Wrapper interface for TableOfContents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfContents WithComCleanupProxy(this Microsoft.Office.Interop.Word.TableOfContents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TableOfContents,Interfaces.ITableOfContents>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfAuthorities which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfAuthorities WithComCleanupProxy(this Microsoft.Office.Interop.Word.TablesOfAuthorities resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TablesOfAuthorities,Interfaces.ITablesOfAuthorities>();
		}

		/// <summary>
		/// Wrapper interface for TableOfAuthorities which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfAuthorities WithComCleanupProxy(this Microsoft.Office.Interop.Word.TableOfAuthorities resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TableOfAuthorities,Interfaces.ITableOfAuthorities>();
		}

		/// <summary>
		/// Wrapper interface for Dialogs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogs WithComCleanupProxy(this Microsoft.Office.Interop.Word.Dialogs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Dialogs,Interfaces.IDialogs>();
		}

		/// <summary>
		/// Wrapper interface for Dialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialog WithComCleanupProxy(this Microsoft.Office.Interop.Word.Dialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Dialog,Interfaces.IDialog>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageSetup WithComCleanupProxy(this Microsoft.Office.Interop.Word.PageSetup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.PageSetup,Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for LineNumbering which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineNumbering WithComCleanupProxy(this Microsoft.Office.Interop.Word.LineNumbering resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LineNumbering,Interfaces.ILineNumbering>();
		}

		/// <summary>
		/// Wrapper interface for TextColumns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextColumns WithComCleanupProxy(this Microsoft.Office.Interop.Word.TextColumns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TextColumns,Interfaces.ITextColumns>();
		}

		/// <summary>
		/// Wrapper interface for TextColumn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextColumn WithComCleanupProxy(this Microsoft.Office.Interop.Word.TextColumn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TextColumn,Interfaces.ITextColumn>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISelection WithComCleanupProxy(this Microsoft.Office.Interop.Word.Selection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Selection,Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for TablesOfAuthoritiesCategories which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITablesOfAuthoritiesCategories WithComCleanupProxy(this Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories,Interfaces.ITablesOfAuthoritiesCategories>();
		}

		/// <summary>
		/// Wrapper interface for TableOfAuthoritiesCategory which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableOfAuthoritiesCategory WithComCleanupProxy(this Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory,Interfaces.ITableOfAuthoritiesCategory>();
		}

		/// <summary>
		/// Wrapper interface for CaptionLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICaptionLabels WithComCleanupProxy(this Microsoft.Office.Interop.Word.CaptionLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CaptionLabels,Interfaces.ICaptionLabels>();
		}

		/// <summary>
		/// Wrapper interface for CaptionLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICaptionLabel WithComCleanupProxy(this Microsoft.Office.Interop.Word.CaptionLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CaptionLabel,Interfaces.ICaptionLabel>();
		}

		/// <summary>
		/// Wrapper interface for AutoCaptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCaptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoCaptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoCaptions,Interfaces.IAutoCaptions>();
		}

		/// <summary>
		/// Wrapper interface for AutoCaption which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCaption WithComCleanupProxy(this Microsoft.Office.Interop.Word.AutoCaption resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AutoCaption,Interfaces.IAutoCaption>();
		}

		/// <summary>
		/// Wrapper interface for Indexes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIndexes WithComCleanupProxy(this Microsoft.Office.Interop.Word.Indexes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Indexes,Interfaces.IIndexes>();
		}

		/// <summary>
		/// Wrapper interface for Index which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIndex WithComCleanupProxy(this Microsoft.Office.Interop.Word.Index resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Index,Interfaces.IIndex>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIn WithComCleanupProxy(this Microsoft.Office.Interop.Word.AddIn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AddIn,Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIns WithComCleanupProxy(this Microsoft.Office.Interop.Word.AddIns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AddIns,Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for Revisions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRevisions WithComCleanupProxy(this Microsoft.Office.Interop.Word.Revisions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Revisions,Interfaces.IRevisions>();
		}

		/// <summary>
		/// Wrapper interface for Revision which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRevision WithComCleanupProxy(this Microsoft.Office.Interop.Word.Revision resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Revision,Interfaces.IRevision>();
		}

		/// <summary>
		/// Wrapper interface for Task which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITask WithComCleanupProxy(this Microsoft.Office.Interop.Word.Task resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Task,Interfaces.ITask>();
		}

		/// <summary>
		/// Wrapper interface for Tasks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITasks WithComCleanupProxy(this Microsoft.Office.Interop.Word.Tasks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Tasks,Interfaces.ITasks>();
		}

		/// <summary>
		/// Wrapper interface for HeadersFooters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadersFooters WithComCleanupProxy(this Microsoft.Office.Interop.Word.HeadersFooters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HeadersFooters,Interfaces.IHeadersFooters>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeaderFooter WithComCleanupProxy(this Microsoft.Office.Interop.Word.HeaderFooter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HeaderFooter,Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for PageNumbers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageNumbers WithComCleanupProxy(this Microsoft.Office.Interop.Word.PageNumbers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.PageNumbers,Interfaces.IPageNumbers>();
		}

		/// <summary>
		/// Wrapper interface for PageNumber which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageNumber WithComCleanupProxy(this Microsoft.Office.Interop.Word.PageNumber resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.PageNumber,Interfaces.IPageNumber>();
		}

		/// <summary>
		/// Wrapper interface for Subdocuments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISubdocuments WithComCleanupProxy(this Microsoft.Office.Interop.Word.Subdocuments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Subdocuments,Interfaces.ISubdocuments>();
		}

		/// <summary>
		/// Wrapper interface for Subdocument which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISubdocument WithComCleanupProxy(this Microsoft.Office.Interop.Word.Subdocument resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Subdocument,Interfaces.ISubdocument>();
		}

		/// <summary>
		/// Wrapper interface for HeadingStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadingStyles WithComCleanupProxy(this Microsoft.Office.Interop.Word.HeadingStyles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HeadingStyles,Interfaces.IHeadingStyles>();
		}

		/// <summary>
		/// Wrapper interface for HeadingStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadingStyle WithComCleanupProxy(this Microsoft.Office.Interop.Word.HeadingStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HeadingStyle,Interfaces.IHeadingStyle>();
		}

		/// <summary>
		/// Wrapper interface for StoryRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStoryRanges WithComCleanupProxy(this Microsoft.Office.Interop.Word.StoryRanges resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.StoryRanges,Interfaces.IStoryRanges>();
		}

		/// <summary>
		/// Wrapper interface for ListLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListLevel WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListLevel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListLevel,Interfaces.IListLevel>();
		}

		/// <summary>
		/// Wrapper interface for ListLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListLevels WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListLevels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListLevels,Interfaces.IListLevels>();
		}

		/// <summary>
		/// Wrapper interface for ListTemplate which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListTemplate WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListTemplate resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListTemplate,Interfaces.IListTemplate>();
		}

		/// <summary>
		/// Wrapper interface for ListTemplates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListTemplates WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListTemplates resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListTemplates,Interfaces.IListTemplates>();
		}

		/// <summary>
		/// Wrapper interface for ListParagraphs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListParagraphs WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListParagraphs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListParagraphs,Interfaces.IListParagraphs>();
		}

		/// <summary>
		/// Wrapper interface for List which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IList WithComCleanupProxy(this Microsoft.Office.Interop.Word.List resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.List,Interfaces.IList>();
		}

		/// <summary>
		/// Wrapper interface for Lists which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILists WithComCleanupProxy(this Microsoft.Office.Interop.Word.Lists resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Lists,Interfaces.ILists>();
		}

		/// <summary>
		/// Wrapper interface for ListGallery which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListGallery WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListGallery resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListGallery,Interfaces.IListGallery>();
		}

		/// <summary>
		/// Wrapper interface for ListGalleries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListGalleries WithComCleanupProxy(this Microsoft.Office.Interop.Word.ListGalleries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ListGalleries,Interfaces.IListGalleries>();
		}

		/// <summary>
		/// Wrapper interface for KeyBindings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IKeyBindings WithComCleanupProxy(this Microsoft.Office.Interop.Word.KeyBindings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.KeyBindings,Interfaces.IKeyBindings>();
		}

		/// <summary>
		/// Wrapper interface for KeysBoundTo which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IKeysBoundTo WithComCleanupProxy(this Microsoft.Office.Interop.Word.KeysBoundTo resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.KeysBoundTo,Interfaces.IKeysBoundTo>();
		}

		/// <summary>
		/// Wrapper interface for KeyBinding which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IKeyBinding WithComCleanupProxy(this Microsoft.Office.Interop.Word.KeyBinding resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.KeyBinding,Interfaces.IKeyBinding>();
		}

		/// <summary>
		/// Wrapper interface for FileConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverter WithComCleanupProxy(this Microsoft.Office.Interop.Word.FileConverter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FileConverter,Interfaces.IFileConverter>();
		}

		/// <summary>
		/// Wrapper interface for FileConverters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverters WithComCleanupProxy(this Microsoft.Office.Interop.Word.FileConverters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FileConverters,Interfaces.IFileConverters>();
		}

		/// <summary>
		/// Wrapper interface for SynonymInfo which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISynonymInfo WithComCleanupProxy(this Microsoft.Office.Interop.Word.SynonymInfo resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SynonymInfo,Interfaces.ISynonymInfo>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlinks WithComCleanupProxy(this Microsoft.Office.Interop.Word.Hyperlinks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Hyperlinks,Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlink WithComCleanupProxy(this Microsoft.Office.Interop.Word.Hyperlink resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Hyperlink,Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapes WithComCleanupProxy(this Microsoft.Office.Interop.Word.Shapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Shapes,Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeRange WithComCleanupProxy(this Microsoft.Office.Interop.Word.ShapeRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ShapeRange,Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupShapes WithComCleanupProxy(this Microsoft.Office.Interop.Word.GroupShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.GroupShapes,Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShape WithComCleanupProxy(this Microsoft.Office.Interop.Word.Shape resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Shape,Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame WithComCleanupProxy(this Microsoft.Office.Interop.Word.TextFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TextFrame,Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for _LetterContent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_LetterContent WithComCleanupProxy(this Microsoft.Office.Interop.Word._LetterContent resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word._LetterContent,Interfaces.I_LetterContent>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IView WithComCleanupProxy(this Microsoft.Office.Interop.Word.View resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.View,Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for Zoom which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IZoom WithComCleanupProxy(this Microsoft.Office.Interop.Word.Zoom resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Zoom,Interfaces.IZoom>();
		}

		/// <summary>
		/// Wrapper interface for Zooms which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IZooms WithComCleanupProxy(this Microsoft.Office.Interop.Word.Zooms resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Zooms,Interfaces.IZooms>();
		}

		/// <summary>
		/// Wrapper interface for InlineShape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInlineShape WithComCleanupProxy(this Microsoft.Office.Interop.Word.InlineShape resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.InlineShape,Interfaces.IInlineShape>();
		}

		/// <summary>
		/// Wrapper interface for InlineShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInlineShapes WithComCleanupProxy(this Microsoft.Office.Interop.Word.InlineShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.InlineShapes,Interfaces.IInlineShapes>();
		}

		/// <summary>
		/// Wrapper interface for SpellingSuggestions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpellingSuggestions WithComCleanupProxy(this Microsoft.Office.Interop.Word.SpellingSuggestions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SpellingSuggestions,Interfaces.ISpellingSuggestions>();
		}

		/// <summary>
		/// Wrapper interface for SpellingSuggestion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpellingSuggestion WithComCleanupProxy(this Microsoft.Office.Interop.Word.SpellingSuggestion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SpellingSuggestion,Interfaces.ISpellingSuggestion>();
		}

		/// <summary>
		/// Wrapper interface for Dictionaries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDictionaries WithComCleanupProxy(this Microsoft.Office.Interop.Word.Dictionaries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Dictionaries,Interfaces.IDictionaries>();
		}

		/// <summary>
		/// Wrapper interface for HangulHanjaConversionDictionaries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHangulHanjaConversionDictionaries WithComCleanupProxy(this Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries,Interfaces.IHangulHanjaConversionDictionaries>();
		}

		/// <summary>
		/// Wrapper interface for Dictionary which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDictionary WithComCleanupProxy(this Microsoft.Office.Interop.Word.Dictionary resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Dictionary,Interfaces.IDictionary>();
		}

		/// <summary>
		/// Wrapper interface for ReadabilityStatistics which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReadabilityStatistics WithComCleanupProxy(this Microsoft.Office.Interop.Word.ReadabilityStatistics resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ReadabilityStatistics,Interfaces.IReadabilityStatistics>();
		}

		/// <summary>
		/// Wrapper interface for ReadabilityStatistic which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReadabilityStatistic WithComCleanupProxy(this Microsoft.Office.Interop.Word.ReadabilityStatistic resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ReadabilityStatistic,Interfaces.IReadabilityStatistic>();
		}

		/// <summary>
		/// Wrapper interface for Versions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVersions WithComCleanupProxy(this Microsoft.Office.Interop.Word.Versions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Versions,Interfaces.IVersions>();
		}

		/// <summary>
		/// Wrapper interface for Version which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVersion WithComCleanupProxy(this Microsoft.Office.Interop.Word.Version resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Version,Interfaces.IVersion>();
		}

		/// <summary>
		/// Wrapper interface for Options which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.Options resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Options,Interfaces.IOptions>();
		}

		/// <summary>
		/// Wrapper interface for MailMessage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailMessage WithComCleanupProxy(this Microsoft.Office.Interop.Word.MailMessage resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MailMessage,Interfaces.IMailMessage>();
		}

		/// <summary>
		/// Wrapper interface for ProofreadingErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProofreadingErrors WithComCleanupProxy(this Microsoft.Office.Interop.Word.ProofreadingErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ProofreadingErrors,Interfaces.IProofreadingErrors>();
		}

		/// <summary>
		/// Wrapper interface for Mailer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailer WithComCleanupProxy(this Microsoft.Office.Interop.Word.Mailer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Mailer,Interfaces.IMailer>();
		}

		/// <summary>
		/// Wrapper interface for WrapFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWrapFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.WrapFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.WrapFormat,Interfaces.IWrapFormat>();
		}

		/// <summary>
		/// Wrapper interface for HangulAndAlphabetExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHangulAndAlphabetExceptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions,Interfaces.IHangulAndAlphabetExceptions>();
		}

		/// <summary>
		/// Wrapper interface for HangulAndAlphabetException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHangulAndAlphabetException WithComCleanupProxy(this Microsoft.Office.Interop.Word.HangulAndAlphabetException resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HangulAndAlphabetException,Interfaces.IHangulAndAlphabetException>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAdjustments WithComCleanupProxy(this Microsoft.Office.Interop.Word.Adjustments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Adjustments,Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalloutFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.CalloutFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CalloutFormat,Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ColorFormat,Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnectorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ConnectorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ConnectorFormat,Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.FillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FillFormat,Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFreeformBuilder WithComCleanupProxy(this Microsoft.Office.Interop.Word.FreeformBuilder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FreeformBuilder,Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.LineFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LineFormat,Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.PictureFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.PictureFormat,Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShadowFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ShadowFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ShadowFormat,Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNode WithComCleanupProxy(this Microsoft.Office.Interop.Word.ShapeNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ShapeNode,Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNodes WithComCleanupProxy(this Microsoft.Office.Interop.Word.ShapeNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ShapeNodes,Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextEffectFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.TextEffectFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TextEffectFormat,Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThreeDFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ThreeDFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ThreeDFormat,Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents,Interfaces.IApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlobal WithComCleanupProxy(this Microsoft.Office.Interop.Word.Global resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Global,Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents_Event,Interfaces.IApplicationEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents2_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents2_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents2_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents2_Event,Interfaces.IApplicationEvents2_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents3_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents3_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents3_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents3_Event,Interfaces.IApplicationEvents3_Event>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents4_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents4_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents4_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents4_Event,Interfaces.IApplicationEvents4_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplication WithComCleanupProxy(this Microsoft.Office.Interop.Word.Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Application,Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents WithComCleanupProxy(this Microsoft.Office.Interop.Word.DocumentEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DocumentEvents,Interfaces.IDocumentEvents>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.DocumentEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DocumentEvents_Event,Interfaces.IDocumentEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents2_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents2_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.DocumentEvents2_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DocumentEvents2_Event,Interfaces.IDocumentEvents2_Event>();
		}

		/// <summary>
		/// Wrapper interface for Document which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocument WithComCleanupProxy(this Microsoft.Office.Interop.Word.Document resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Document,Interfaces.IDocument>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFont WithComCleanupProxy(this Microsoft.Office.Interop.Word.Font resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Font,Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ParagraphFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ParagraphFormat,Interfaces.IParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for OCXEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXEvents WithComCleanupProxy(this Microsoft.Office.Interop.Word.OCXEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OCXEvents,Interfaces.IOCXEvents>();
		}

		/// <summary>
		/// Wrapper interface for OCXEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Word.OCXEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OCXEvents_Event,Interfaces.IOCXEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEControl WithComCleanupProxy(this Microsoft.Office.Interop.Word.OLEControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OLEControl,Interfaces.IOLEControl>();
		}

		/// <summary>
		/// Wrapper interface for LetterContent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILetterContent WithComCleanupProxy(this Microsoft.Office.Interop.Word.LetterContent resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LetterContent,Interfaces.ILetterContent>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents WithComCleanupProxy(this Microsoft.Office.Interop.Word.IApplicationEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.IApplicationEvents,Interfaces.IIApplicationEvents>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents2 WithComCleanupProxy(this Microsoft.Office.Interop.Word.IApplicationEvents2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.IApplicationEvents2,Interfaces.IIApplicationEvents2>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents2 WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents2,Interfaces.IApplicationEvents2>();
		}

		/// <summary>
		/// Wrapper interface for EmailAuthor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailAuthor WithComCleanupProxy(this Microsoft.Office.Interop.Word.EmailAuthor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.EmailAuthor,Interfaces.IEmailAuthor>();
		}

		/// <summary>
		/// Wrapper interface for EmailOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailOptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.EmailOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.EmailOptions,Interfaces.IEmailOptions>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignature which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailSignature WithComCleanupProxy(this Microsoft.Office.Interop.Word.EmailSignature resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.EmailSignature,Interfaces.IEmailSignature>();
		}

		/// <summary>
		/// Wrapper interface for Email which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmail WithComCleanupProxy(this Microsoft.Office.Interop.Word.Email resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Email,Interfaces.IEmail>();
		}

		/// <summary>
		/// Wrapper interface for HorizontalLineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHorizontalLineFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.HorizontalLineFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HorizontalLineFormat,Interfaces.IHorizontalLineFormat>();
		}

		/// <summary>
		/// Wrapper interface for Frameset which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFrameset WithComCleanupProxy(this Microsoft.Office.Interop.Word.Frameset resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Frameset,Interfaces.IFrameset>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDefaultWebOptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.DefaultWebOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DefaultWebOptions,Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebOptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.WebOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.WebOptions,Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for OtherCorrectionsExceptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOtherCorrectionsExceptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.OtherCorrectionsExceptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OtherCorrectionsExceptions,Interfaces.IOtherCorrectionsExceptions>();
		}

		/// <summary>
		/// Wrapper interface for OtherCorrectionsException which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOtherCorrectionsException WithComCleanupProxy(this Microsoft.Office.Interop.Word.OtherCorrectionsException resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OtherCorrectionsException,Interfaces.IOtherCorrectionsException>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignatureEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailSignatureEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.EmailSignatureEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.EmailSignatureEntries,Interfaces.IEmailSignatureEntries>();
		}

		/// <summary>
		/// Wrapper interface for EmailSignatureEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEmailSignatureEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.EmailSignatureEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.EmailSignatureEntry,Interfaces.IEmailSignatureEntry>();
		}

		/// <summary>
		/// Wrapper interface for HTMLDivision which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLDivision WithComCleanupProxy(this Microsoft.Office.Interop.Word.HTMLDivision resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HTMLDivision,Interfaces.IHTMLDivision>();
		}

		/// <summary>
		/// Wrapper interface for HTMLDivisions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHTMLDivisions WithComCleanupProxy(this Microsoft.Office.Interop.Word.HTMLDivisions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HTMLDivisions,Interfaces.IHTMLDivisions>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNode WithComCleanupProxy(this Microsoft.Office.Interop.Word.DiagramNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DiagramNode,Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodeChildren WithComCleanupProxy(this Microsoft.Office.Interop.Word.DiagramNodeChildren resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DiagramNodeChildren,Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodes WithComCleanupProxy(this Microsoft.Office.Interop.Word.DiagramNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DiagramNodes,Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagram WithComCleanupProxy(this Microsoft.Office.Interop.Word.Diagram resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Diagram,Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomProperty WithComCleanupProxy(this Microsoft.Office.Interop.Word.CustomProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CustomProperty,Interfaces.ICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomProperties WithComCleanupProxy(this Microsoft.Office.Interop.Word.CustomProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CustomProperties,Interfaces.ICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for SmartTag which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTag WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTag resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTag,Interfaces.ISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for SmartTags which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTags WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTags resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTags,Interfaces.ISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for StyleSheet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyleSheet WithComCleanupProxy(this Microsoft.Office.Interop.Word.StyleSheet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.StyleSheet,Interfaces.IStyleSheet>();
		}

		/// <summary>
		/// Wrapper interface for StyleSheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyleSheets WithComCleanupProxy(this Microsoft.Office.Interop.Word.StyleSheets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.StyleSheets,Interfaces.IStyleSheets>();
		}

		/// <summary>
		/// Wrapper interface for MappedDataField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMappedDataField WithComCleanupProxy(this Microsoft.Office.Interop.Word.MappedDataField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MappedDataField,Interfaces.IMappedDataField>();
		}

		/// <summary>
		/// Wrapper interface for MappedDataFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMappedDataFields WithComCleanupProxy(this Microsoft.Office.Interop.Word.MappedDataFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.MappedDataFields,Interfaces.IMappedDataFields>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICanvasShapes WithComCleanupProxy(this Microsoft.Office.Interop.Word.CanvasShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CanvasShapes,Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyle WithComCleanupProxy(this Microsoft.Office.Interop.Word.TableStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TableStyle,Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for ConditionalStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConditionalStyle WithComCleanupProxy(this Microsoft.Office.Interop.Word.ConditionalStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ConditionalStyle,Interfaces.IConditionalStyle>();
		}

		/// <summary>
		/// Wrapper interface for FootnoteOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFootnoteOptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.FootnoteOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.FootnoteOptions,Interfaces.IFootnoteOptions>();
		}

		/// <summary>
		/// Wrapper interface for EndnoteOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEndnoteOptions WithComCleanupProxy(this Microsoft.Office.Interop.Word.EndnoteOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.EndnoteOptions,Interfaces.IEndnoteOptions>();
		}

		/// <summary>
		/// Wrapper interface for Reviewers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReviewers WithComCleanupProxy(this Microsoft.Office.Interop.Word.Reviewers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Reviewers,Interfaces.IReviewers>();
		}

		/// <summary>
		/// Wrapper interface for Reviewer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReviewer WithComCleanupProxy(this Microsoft.Office.Interop.Word.Reviewer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Reviewer,Interfaces.IReviewer>();
		}

		/// <summary>
		/// Wrapper interface for TaskPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskPane WithComCleanupProxy(this Microsoft.Office.Interop.Word.TaskPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TaskPane,Interfaces.ITaskPane>();
		}

		/// <summary>
		/// Wrapper interface for TaskPanes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITaskPanes WithComCleanupProxy(this Microsoft.Office.Interop.Word.TaskPanes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TaskPanes,Interfaces.ITaskPanes>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents3 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents3 WithComCleanupProxy(this Microsoft.Office.Interop.Word.IApplicationEvents3 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.IApplicationEvents3,Interfaces.IIApplicationEvents3>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents3 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents3 WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents3 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents3,Interfaces.IApplicationEvents3>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagAction WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTagAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTagAction,Interfaces.ISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagActions WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTagActions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTagActions,Interfaces.ISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagRecognizer WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTagRecognizer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTagRecognizer,Interfaces.ISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagRecognizers WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTagRecognizers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTagRecognizers,Interfaces.ISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagType which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagType WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTagType resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTagType,Interfaces.ISmartTagType>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagTypes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagTypes WithComCleanupProxy(this Microsoft.Office.Interop.Word.SmartTagTypes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SmartTagTypes,Interfaces.ISmartTagTypes>();
		}

		/// <summary>
		/// Wrapper interface for Line which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILine WithComCleanupProxy(this Microsoft.Office.Interop.Word.Line resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Line,Interfaces.ILine>();
		}

		/// <summary>
		/// Wrapper interface for Lines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILines WithComCleanupProxy(this Microsoft.Office.Interop.Word.Lines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Lines,Interfaces.ILines>();
		}

		/// <summary>
		/// Wrapper interface for Rectangle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangle WithComCleanupProxy(this Microsoft.Office.Interop.Word.Rectangle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Rectangle,Interfaces.IRectangle>();
		}

		/// <summary>
		/// Wrapper interface for Rectangles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangles WithComCleanupProxy(this Microsoft.Office.Interop.Word.Rectangles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Rectangles,Interfaces.IRectangles>();
		}

		/// <summary>
		/// Wrapper interface for Break which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBreak WithComCleanupProxy(this Microsoft.Office.Interop.Word.Break resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Break,Interfaces.IBreak>();
		}

		/// <summary>
		/// Wrapper interface for Breaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBreaks WithComCleanupProxy(this Microsoft.Office.Interop.Word.Breaks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Breaks,Interfaces.IBreaks>();
		}

		/// <summary>
		/// Wrapper interface for Page which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPage WithComCleanupProxy(this Microsoft.Office.Interop.Word.Page resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Page,Interfaces.IPage>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPages WithComCleanupProxy(this Microsoft.Office.Interop.Word.Pages resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Pages,Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for XMLNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNode WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLNode,Interfaces.IXMLNode>();
		}

		/// <summary>
		/// Wrapper interface for XMLNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNodes WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLNodes,Interfaces.IXMLNodes>();
		}

		/// <summary>
		/// Wrapper interface for XMLSchemaReference which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLSchemaReference WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLSchemaReference resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLSchemaReference,Interfaces.IXMLSchemaReference>();
		}

		/// <summary>
		/// Wrapper interface for XMLSchemaReferences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLSchemaReferences WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLSchemaReferences resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLSchemaReferences,Interfaces.IXMLSchemaReferences>();
		}

		/// <summary>
		/// Wrapper interface for XMLChildNodeSuggestion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLChildNodeSuggestion WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLChildNodeSuggestion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLChildNodeSuggestion,Interfaces.IXMLChildNodeSuggestion>();
		}

		/// <summary>
		/// Wrapper interface for XMLChildNodeSuggestions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLChildNodeSuggestions WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLChildNodeSuggestions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLChildNodeSuggestions,Interfaces.IXMLChildNodeSuggestions>();
		}

		/// <summary>
		/// Wrapper interface for XMLNamespace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNamespace WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLNamespace resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLNamespace,Interfaces.IXMLNamespace>();
		}

		/// <summary>
		/// Wrapper interface for XMLNamespaces which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLNamespaces WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLNamespaces resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLNamespaces,Interfaces.IXMLNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for XSLTransform which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXSLTransform WithComCleanupProxy(this Microsoft.Office.Interop.Word.XSLTransform resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XSLTransform,Interfaces.IXSLTransform>();
		}

		/// <summary>
		/// Wrapper interface for XSLTransforms which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXSLTransforms WithComCleanupProxy(this Microsoft.Office.Interop.Word.XSLTransforms resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XSLTransforms,Interfaces.IXSLTransforms>();
		}

		/// <summary>
		/// Wrapper interface for Editors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEditors WithComCleanupProxy(this Microsoft.Office.Interop.Word.Editors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Editors,Interfaces.IEditors>();
		}

		/// <summary>
		/// Wrapper interface for Editor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEditor WithComCleanupProxy(this Microsoft.Office.Interop.Word.Editor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Editor,Interfaces.IEditor>();
		}

		/// <summary>
		/// Wrapper interface for IApplicationEvents4 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIApplicationEvents4 WithComCleanupProxy(this Microsoft.Office.Interop.Word.IApplicationEvents4 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.IApplicationEvents4,Interfaces.IIApplicationEvents4>();
		}

		/// <summary>
		/// Wrapper interface for ApplicationEvents4 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplicationEvents4 WithComCleanupProxy(this Microsoft.Office.Interop.Word.ApplicationEvents4 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ApplicationEvents4,Interfaces.IApplicationEvents4>();
		}

		/// <summary>
		/// Wrapper interface for DocumentEvents2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentEvents2 WithComCleanupProxy(this Microsoft.Office.Interop.Word.DocumentEvents2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DocumentEvents2,Interfaces.IDocumentEvents2>();
		}

		/// <summary>
		/// Wrapper interface for Source which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISource WithComCleanupProxy(this Microsoft.Office.Interop.Word.Source resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Source,Interfaces.ISource>();
		}

		/// <summary>
		/// Wrapper interface for Sources which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISources WithComCleanupProxy(this Microsoft.Office.Interop.Word.Sources resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Sources,Interfaces.ISources>();
		}

		/// <summary>
		/// Wrapper interface for Bibliography which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBibliography WithComCleanupProxy(this Microsoft.Office.Interop.Word.Bibliography resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Bibliography,Interfaces.IBibliography>();
		}

		/// <summary>
		/// Wrapper interface for OMaths which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMaths WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMaths resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMaths,Interfaces.IOMaths>();
		}

		/// <summary>
		/// Wrapper interface for OMath which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMath WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMath resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMath,Interfaces.IOMath>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunctions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFunctions WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathFunctions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathFunctions,Interfaces.IOMathFunctions>();
		}

		/// <summary>
		/// Wrapper interface for OMathArgs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathArgs WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathArgs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathArgs,Interfaces.IOMathArgs>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFunction WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathFunction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathFunction,Interfaces.IOMathFunction>();
		}

		/// <summary>
		/// Wrapper interface for OMathAcc which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAcc WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathAcc resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathAcc,Interfaces.IOMathAcc>();
		}

		/// <summary>
		/// Wrapper interface for OMathBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBar WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathBar,Interfaces.IOMathBar>();
		}

		/// <summary>
		/// Wrapper interface for OMathBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBox WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathBox,Interfaces.IOMathBox>();
		}

		/// <summary>
		/// Wrapper interface for OMathBorderBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBorderBox WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathBorderBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathBorderBox,Interfaces.IOMathBorderBox>();
		}

		/// <summary>
		/// Wrapper interface for OMathDelim which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathDelim WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathDelim resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathDelim,Interfaces.IOMathDelim>();
		}

		/// <summary>
		/// Wrapper interface for OMathEqArray which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathEqArray WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathEqArray resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathEqArray,Interfaces.IOMathEqArray>();
		}

		/// <summary>
		/// Wrapper interface for OMathFrac which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFrac WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathFrac resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathFrac,Interfaces.IOMathFrac>();
		}

		/// <summary>
		/// Wrapper interface for OMathFunc which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathFunc WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathFunc resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathFunc,Interfaces.IOMathFunc>();
		}

		/// <summary>
		/// Wrapper interface for OMathGroupChar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathGroupChar WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathGroupChar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathGroupChar,Interfaces.IOMathGroupChar>();
		}

		/// <summary>
		/// Wrapper interface for OMathLimLow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathLimLow WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathLimLow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathLimLow,Interfaces.IOMathLimLow>();
		}

		/// <summary>
		/// Wrapper interface for OMathLimUpp which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathLimUpp WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathLimUpp resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathLimUpp,Interfaces.IOMathLimUpp>();
		}

		/// <summary>
		/// Wrapper interface for OMathMat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMat WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathMat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathMat,Interfaces.IOMathMat>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatRows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatRows WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathMatRows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathMatRows,Interfaces.IOMathMatRows>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatCols which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatCols WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathMatCols resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathMatCols,Interfaces.IOMathMatCols>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatRow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatRow WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathMatRow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathMatRow,Interfaces.IOMathMatRow>();
		}

		/// <summary>
		/// Wrapper interface for OMathMatCol which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathMatCol WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathMatCol resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathMatCol,Interfaces.IOMathMatCol>();
		}

		/// <summary>
		/// Wrapper interface for OMathNary which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathNary WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathNary resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathNary,Interfaces.IOMathNary>();
		}

		/// <summary>
		/// Wrapper interface for OMathPhantom which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathPhantom WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathPhantom resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathPhantom,Interfaces.IOMathPhantom>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrPre which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrPre WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathScrPre resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathScrPre,Interfaces.IOMathScrPre>();
		}

		/// <summary>
		/// Wrapper interface for OMathRad which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathRad WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathRad resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathRad,Interfaces.IOMathRad>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSub which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrSub WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathScrSub resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathScrSub,Interfaces.IOMathScrSub>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSubSup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrSubSup WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathScrSubSup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathScrSubSup,Interfaces.IOMathScrSubSup>();
		}

		/// <summary>
		/// Wrapper interface for OMathScrSup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathScrSup WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathScrSup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathScrSup,Interfaces.IOMathScrSup>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAutoCorrect WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathAutoCorrect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathAutoCorrect,Interfaces.IOMathAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrectEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAutoCorrectEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathAutoCorrectEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathAutoCorrectEntries,Interfaces.IOMathAutoCorrectEntries>();
		}

		/// <summary>
		/// Wrapper interface for OMathAutoCorrectEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathAutoCorrectEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathAutoCorrectEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathAutoCorrectEntry,Interfaces.IOMathAutoCorrectEntry>();
		}

		/// <summary>
		/// Wrapper interface for OMathRecognizedFunctions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathRecognizedFunctions WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathRecognizedFunctions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathRecognizedFunctions,Interfaces.IOMathRecognizedFunctions>();
		}

		/// <summary>
		/// Wrapper interface for OMathRecognizedFunction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathRecognizedFunction WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathRecognizedFunction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathRecognizedFunction,Interfaces.IOMathRecognizedFunction>();
		}

		/// <summary>
		/// Wrapper interface for ContentControls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControls WithComCleanupProxy(this Microsoft.Office.Interop.Word.ContentControls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ContentControls,Interfaces.IContentControls>();
		}

		/// <summary>
		/// Wrapper interface for ContentControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControl WithComCleanupProxy(this Microsoft.Office.Interop.Word.ContentControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ContentControl,Interfaces.IContentControl>();
		}

		/// <summary>
		/// Wrapper interface for XMLMapping which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXMLMapping WithComCleanupProxy(this Microsoft.Office.Interop.Word.XMLMapping resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.XMLMapping,Interfaces.IXMLMapping>();
		}

		/// <summary>
		/// Wrapper interface for ContentControlListEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControlListEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.ContentControlListEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ContentControlListEntries,Interfaces.IContentControlListEntries>();
		}

		/// <summary>
		/// Wrapper interface for ContentControlListEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IContentControlListEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.ContentControlListEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ContentControlListEntry,Interfaces.IContentControlListEntry>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockTypes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlockTypes WithComCleanupProxy(this Microsoft.Office.Interop.Word.BuildingBlockTypes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.BuildingBlockTypes,Interfaces.IBuildingBlockTypes>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockType which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlockType WithComCleanupProxy(this Microsoft.Office.Interop.Word.BuildingBlockType resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.BuildingBlockType,Interfaces.IBuildingBlockType>();
		}

		/// <summary>
		/// Wrapper interface for Categories which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategories WithComCleanupProxy(this Microsoft.Office.Interop.Word.Categories resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Categories,Interfaces.ICategories>();
		}

		/// <summary>
		/// Wrapper interface for Category which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICategory WithComCleanupProxy(this Microsoft.Office.Interop.Word.Category resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Category,Interfaces.ICategory>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlocks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlocks WithComCleanupProxy(this Microsoft.Office.Interop.Word.BuildingBlocks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.BuildingBlocks,Interfaces.IBuildingBlocks>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlock which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlock WithComCleanupProxy(this Microsoft.Office.Interop.Word.BuildingBlock resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.BuildingBlock,Interfaces.IBuildingBlock>();
		}

		/// <summary>
		/// Wrapper interface for BuildingBlockEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBuildingBlockEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.BuildingBlockEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.BuildingBlockEntries,Interfaces.IBuildingBlockEntries>();
		}

		/// <summary>
		/// Wrapper interface for OMathBreaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBreaks WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathBreaks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathBreaks,Interfaces.IOMathBreaks>();
		}

		/// <summary>
		/// Wrapper interface for OMathBreak which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOMathBreak WithComCleanupProxy(this Microsoft.Office.Interop.Word.OMathBreak resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.OMathBreak,Interfaces.IOMathBreak>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResearch WithComCleanupProxy(this Microsoft.Office.Interop.Word.Research resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Research,Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoftEdgeFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.SoftEdgeFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SoftEdgeFormat,Interfaces.ISoftEdgeFormat>();
		}

		/// <summary>
		/// Wrapper interface for GlowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlowFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.GlowFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.GlowFormat,Interfaces.IGlowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IReflectionFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ReflectionFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ReflectionFormat,Interfaces.IReflectionFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartData WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartData resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartData,Interfaces.IChartData>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChart WithComCleanupProxy(this Microsoft.Office.Interop.Word.Chart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Chart,Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICorners WithComCleanupProxy(this Microsoft.Office.Interop.Word.Corners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Corners,Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegend WithComCleanupProxy(this Microsoft.Office.Interop.Word.Legend resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Legend,Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for ChartBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartBorder WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartBorder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartBorder,Interfaces.IChartBorder>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWalls WithComCleanupProxy(this Microsoft.Office.Interop.Word.Walls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Walls,Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFloor WithComCleanupProxy(this Microsoft.Office.Interop.Word.Floor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Floor,Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlotArea WithComCleanupProxy(this Microsoft.Office.Interop.Word.PlotArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.PlotArea,Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartArea WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartArea,Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesLines WithComCleanupProxy(this Microsoft.Office.Interop.Word.SeriesLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SeriesLines,Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILeaderLines WithComCleanupProxy(this Microsoft.Office.Interop.Word.LeaderLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LeaderLines,Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGridlines WithComCleanupProxy(this Microsoft.Office.Interop.Word.Gridlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Gridlines,Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUpBars WithComCleanupProxy(this Microsoft.Office.Interop.Word.UpBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.UpBars,Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDownBars WithComCleanupProxy(this Microsoft.Office.Interop.Word.DownBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DownBars,Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInterior WithComCleanupProxy(this Microsoft.Office.Interop.Word.Interior resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Interior,Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartFillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartFillFormat,Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntries WithComCleanupProxy(this Microsoft.Office.Interop.Word.LegendEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LegendEntries,Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFont WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartFont,Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartColorFormat,Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntry WithComCleanupProxy(this Microsoft.Office.Interop.Word.LegendEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LegendEntry,Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendKey WithComCleanupProxy(this Microsoft.Office.Interop.Word.LegendKey resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.LegendKey,Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesCollection WithComCleanupProxy(this Microsoft.Office.Interop.Word.SeriesCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.SeriesCollection,Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeries WithComCleanupProxy(this Microsoft.Office.Interop.Word.Series resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Series,Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrorBars WithComCleanupProxy(this Microsoft.Office.Interop.Word.ErrorBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ErrorBars,Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendline WithComCleanupProxy(this Microsoft.Office.Interop.Word.Trendline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Trendline,Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendlines WithComCleanupProxy(this Microsoft.Office.Interop.Word.Trendlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Trendlines,Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabels WithComCleanupProxy(this Microsoft.Office.Interop.Word.DataLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DataLabels,Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabel WithComCleanupProxy(this Microsoft.Office.Interop.Word.DataLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DataLabel,Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoints WithComCleanupProxy(this Microsoft.Office.Interop.Word.Points resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Points,Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoint WithComCleanupProxy(this Microsoft.Office.Interop.Word.Point resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Point,Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxes WithComCleanupProxy(this Microsoft.Office.Interop.Word.Axes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Axes,Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxis WithComCleanupProxy(this Microsoft.Office.Interop.Word.Axis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Axis,Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataTable WithComCleanupProxy(this Microsoft.Office.Interop.Word.DataTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DataTable,Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartTitle WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartTitle,Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxisTitle WithComCleanupProxy(this Microsoft.Office.Interop.Word.AxisTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.AxisTitle,Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDisplayUnitLabel WithComCleanupProxy(this Microsoft.Office.Interop.Word.DisplayUnitLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DisplayUnitLabel,Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITickLabels WithComCleanupProxy(this Microsoft.Office.Interop.Word.TickLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.TickLabels,Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropLines WithComCleanupProxy(this Microsoft.Office.Interop.Word.DropLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.DropLines,Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHiLoLines WithComCleanupProxy(this Microsoft.Office.Interop.Word.HiLoLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.HiLoLines,Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroup WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartGroup,Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroups WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartGroups,Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for ChartCharacters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartCharacters WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartCharacters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartCharacters,Interfaces.IChartCharacters>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFormat WithComCleanupProxy(this Microsoft.Office.Interop.Word.ChartFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ChartFormat,Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for UndoRecord which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUndoRecord WithComCleanupProxy(this Microsoft.Office.Interop.Word.UndoRecord resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.UndoRecord,Interfaces.IUndoRecord>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthLock which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthLock WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthLock resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthLock,Interfaces.ICoAuthLock>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthLocks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthLocks WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthLocks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthLocks,Interfaces.ICoAuthLocks>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthUpdate which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthUpdate WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthUpdate resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthUpdate,Interfaces.ICoAuthUpdate>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthUpdates which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthUpdates WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthUpdates resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthUpdates,Interfaces.ICoAuthUpdates>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthor WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthor,Interfaces.ICoAuthor>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthors WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthors,Interfaces.ICoAuthors>();
		}

		/// <summary>
		/// Wrapper interface for CoAuthoring which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoAuthoring WithComCleanupProxy(this Microsoft.Office.Interop.Word.CoAuthoring resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.CoAuthoring,Interfaces.ICoAuthoring>();
		}

		/// <summary>
		/// Wrapper interface for Conflicts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConflicts WithComCleanupProxy(this Microsoft.Office.Interop.Word.Conflicts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Conflicts,Interfaces.IConflicts>();
		}

		/// <summary>
		/// Wrapper interface for Conflict which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConflict WithComCleanupProxy(this Microsoft.Office.Interop.Word.Conflict resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.Conflict,Interfaces.IConflict>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindows WithComCleanupProxy(this Microsoft.Office.Interop.Word.ProtectedViewWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ProtectedViewWindows,Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindow WithComCleanupProxy(this Microsoft.Office.Interop.Word.ProtectedViewWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Word.ProtectedViewWindow,Interfaces.IProtectedViewWindow>();
		}

	}
}