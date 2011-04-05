//Microsoft.Office.Interop.Word, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.Word.Extensions.Proxies.Interfaces
{
	/// <summary>
	/// Wrapper interface for _Application which adds IDispose to the interface
	/// </summary>
	public interface I_Application : Microsoft.Office.Interop.Word._Application, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Global which adds IDispose to the interface
	/// </summary>
	public interface I_Global : Microsoft.Office.Interop.Word._Global, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._Global Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FontNames which adds IDispose to the interface
	/// </summary>
	public interface IFontNames : Microsoft.Office.Interop.Word.FontNames, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FontNames Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Languages which adds IDispose to the interface
	/// </summary>
	public interface ILanguages : Microsoft.Office.Interop.Word.Languages, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Languages Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Language which adds IDispose to the interface
	/// </summary>
	public interface ILanguage : Microsoft.Office.Interop.Word.Language, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Language Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Documents which adds IDispose to the interface
	/// </summary>
	public interface IDocuments : Microsoft.Office.Interop.Word.Documents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Documents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Document which adds IDispose to the interface
	/// </summary>
	public interface I_Document : Microsoft.Office.Interop.Word._Document, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._Document Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Template which adds IDispose to the interface
	/// </summary>
	public interface ITemplate : Microsoft.Office.Interop.Word.Template, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Template Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Templates which adds IDispose to the interface
	/// </summary>
	public interface ITemplates : Microsoft.Office.Interop.Word.Templates, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Templates Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RoutingSlip which adds IDispose to the interface
	/// </summary>
	public interface IRoutingSlip : Microsoft.Office.Interop.Word.RoutingSlip, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.RoutingSlip Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Bookmark which adds IDispose to the interface
	/// </summary>
	public interface IBookmark : Microsoft.Office.Interop.Word.Bookmark, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Bookmark Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Bookmarks which adds IDispose to the interface
	/// </summary>
	public interface IBookmarks : Microsoft.Office.Interop.Word.Bookmarks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Bookmarks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Variable which adds IDispose to the interface
	/// </summary>
	public interface IVariable : Microsoft.Office.Interop.Word.Variable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Variable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Variables which adds IDispose to the interface
	/// </summary>
	public interface IVariables : Microsoft.Office.Interop.Word.Variables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Variables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RecentFile which adds IDispose to the interface
	/// </summary>
	public interface IRecentFile : Microsoft.Office.Interop.Word.RecentFile, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.RecentFile Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RecentFiles which adds IDispose to the interface
	/// </summary>
	public interface IRecentFiles : Microsoft.Office.Interop.Word.RecentFiles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.RecentFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Window which adds IDispose to the interface
	/// </summary>
	public interface IWindow : Microsoft.Office.Interop.Word.Window, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Window Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Windows which adds IDispose to the interface
	/// </summary>
	public interface IWindows : Microsoft.Office.Interop.Word.Windows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Windows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pane which adds IDispose to the interface
	/// </summary>
	public interface IPane : Microsoft.Office.Interop.Word.Pane, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Pane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Panes which adds IDispose to the interface
	/// </summary>
	public interface IPanes : Microsoft.Office.Interop.Word.Panes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Panes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Range which adds IDispose to the interface
	/// </summary>
	public interface IRange : Microsoft.Office.Interop.Word.Range, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Range Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListFormat which adds IDispose to the interface
	/// </summary>
	public interface IListFormat : Microsoft.Office.Interop.Word.ListFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Find which adds IDispose to the interface
	/// </summary>
	public interface IFind : Microsoft.Office.Interop.Word.Find, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Find Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Replacement which adds IDispose to the interface
	/// </summary>
	public interface IReplacement : Microsoft.Office.Interop.Word.Replacement, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Replacement Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Characters which adds IDispose to the interface
	/// </summary>
	public interface ICharacters : Microsoft.Office.Interop.Word.Characters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Characters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Words which adds IDispose to the interface
	/// </summary>
	public interface IWords : Microsoft.Office.Interop.Word.Words, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Words Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sentences which adds IDispose to the interface
	/// </summary>
	public interface ISentences : Microsoft.Office.Interop.Word.Sentences, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Sentences Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sections which adds IDispose to the interface
	/// </summary>
	public interface ISections : Microsoft.Office.Interop.Word.Sections, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Sections Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Section which adds IDispose to the interface
	/// </summary>
	public interface ISection : Microsoft.Office.Interop.Word.Section, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Section Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Paragraphs which adds IDispose to the interface
	/// </summary>
	public interface IParagraphs : Microsoft.Office.Interop.Word.Paragraphs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Paragraphs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Paragraph which adds IDispose to the interface
	/// </summary>
	public interface IParagraph : Microsoft.Office.Interop.Word.Paragraph, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Paragraph Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropCap which adds IDispose to the interface
	/// </summary>
	public interface IDropCap : Microsoft.Office.Interop.Word.DropCap, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DropCap Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TabStops which adds IDispose to the interface
	/// </summary>
	public interface ITabStops : Microsoft.Office.Interop.Word.TabStops, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TabStops Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TabStop which adds IDispose to the interface
	/// </summary>
	public interface ITabStop : Microsoft.Office.Interop.Word.TabStop, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TabStop Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _ParagraphFormat which adds IDispose to the interface
	/// </summary>
	public interface I_ParagraphFormat : Microsoft.Office.Interop.Word._ParagraphFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._ParagraphFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Font which adds IDispose to the interface
	/// </summary>
	public interface I_Font : Microsoft.Office.Interop.Word._Font, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._Font Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Table which adds IDispose to the interface
	/// </summary>
	public interface ITable : Microsoft.Office.Interop.Word.Table, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Table Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Row which adds IDispose to the interface
	/// </summary>
	public interface IRow : Microsoft.Office.Interop.Word.Row, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Row Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Column which adds IDispose to the interface
	/// </summary>
	public interface IColumn : Microsoft.Office.Interop.Word.Column, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Column Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Cell which adds IDispose to the interface
	/// </summary>
	public interface ICell : Microsoft.Office.Interop.Word.Cell, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Cell Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Tables which adds IDispose to the interface
	/// </summary>
	public interface ITables : Microsoft.Office.Interop.Word.Tables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Tables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rows which adds IDispose to the interface
	/// </summary>
	public interface IRows : Microsoft.Office.Interop.Word.Rows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Rows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Columns which adds IDispose to the interface
	/// </summary>
	public interface IColumns : Microsoft.Office.Interop.Word.Columns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Columns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Cells which adds IDispose to the interface
	/// </summary>
	public interface ICells : Microsoft.Office.Interop.Word.Cells, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Cells Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCorrect which adds IDispose to the interface
	/// </summary>
	public interface IAutoCorrect : Microsoft.Office.Interop.Word.AutoCorrect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoCorrect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCorrectEntries which adds IDispose to the interface
	/// </summary>
	public interface IAutoCorrectEntries : Microsoft.Office.Interop.Word.AutoCorrectEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoCorrectEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCorrectEntry which adds IDispose to the interface
	/// </summary>
	public interface IAutoCorrectEntry : Microsoft.Office.Interop.Word.AutoCorrectEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoCorrectEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FirstLetterExceptions which adds IDispose to the interface
	/// </summary>
	public interface IFirstLetterExceptions : Microsoft.Office.Interop.Word.FirstLetterExceptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FirstLetterExceptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FirstLetterException which adds IDispose to the interface
	/// </summary>
	public interface IFirstLetterException : Microsoft.Office.Interop.Word.FirstLetterException, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FirstLetterException Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TwoInitialCapsExceptions which adds IDispose to the interface
	/// </summary>
	public interface ITwoInitialCapsExceptions : Microsoft.Office.Interop.Word.TwoInitialCapsExceptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TwoInitialCapsExceptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TwoInitialCapsException which adds IDispose to the interface
	/// </summary>
	public interface ITwoInitialCapsException : Microsoft.Office.Interop.Word.TwoInitialCapsException, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TwoInitialCapsException Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Footnotes which adds IDispose to the interface
	/// </summary>
	public interface IFootnotes : Microsoft.Office.Interop.Word.Footnotes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Footnotes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Endnotes which adds IDispose to the interface
	/// </summary>
	public interface IEndnotes : Microsoft.Office.Interop.Word.Endnotes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Endnotes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Comments which adds IDispose to the interface
	/// </summary>
	public interface IComments : Microsoft.Office.Interop.Word.Comments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Comments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Footnote which adds IDispose to the interface
	/// </summary>
	public interface IFootnote : Microsoft.Office.Interop.Word.Footnote, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Footnote Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Endnote which adds IDispose to the interface
	/// </summary>
	public interface IEndnote : Microsoft.Office.Interop.Word.Endnote, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Endnote Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Comment which adds IDispose to the interface
	/// </summary>
	public interface IComment : Microsoft.Office.Interop.Word.Comment, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Comment Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Borders which adds IDispose to the interface
	/// </summary>
	public interface IBorders : Microsoft.Office.Interop.Word.Borders, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Borders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Border which adds IDispose to the interface
	/// </summary>
	public interface IBorder : Microsoft.Office.Interop.Word.Border, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Border Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shading which adds IDispose to the interface
	/// </summary>
	public interface IShading : Microsoft.Office.Interop.Word.Shading, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Shading Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextRetrievalMode which adds IDispose to the interface
	/// </summary>
	public interface ITextRetrievalMode : Microsoft.Office.Interop.Word.TextRetrievalMode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TextRetrievalMode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoTextEntries which adds IDispose to the interface
	/// </summary>
	public interface IAutoTextEntries : Microsoft.Office.Interop.Word.AutoTextEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoTextEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoTextEntry which adds IDispose to the interface
	/// </summary>
	public interface IAutoTextEntry : Microsoft.Office.Interop.Word.AutoTextEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoTextEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for System which adds IDispose to the interface
	/// </summary>
	public interface ISystem : Microsoft.Office.Interop.Word.System, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.System Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEFormat which adds IDispose to the interface
	/// </summary>
	public interface IOLEFormat : Microsoft.Office.Interop.Word.OLEFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OLEFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LinkFormat which adds IDispose to the interface
	/// </summary>
	public interface ILinkFormat : Microsoft.Office.Interop.Word.LinkFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LinkFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OLEControl which adds IDispose to the interface
	/// </summary>
	public interface I_OLEControl : Microsoft.Office.Interop.Word._OLEControl, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._OLEControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Fields which adds IDispose to the interface
	/// </summary>
	public interface IFields : Microsoft.Office.Interop.Word.Fields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Fields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Field which adds IDispose to the interface
	/// </summary>
	public interface IField : Microsoft.Office.Interop.Word.Field, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Field Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Browser which adds IDispose to the interface
	/// </summary>
	public interface IBrowser : Microsoft.Office.Interop.Word.Browser, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Browser Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Styles which adds IDispose to the interface
	/// </summary>
	public interface IStyles : Microsoft.Office.Interop.Word.Styles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Styles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Style which adds IDispose to the interface
	/// </summary>
	public interface IStyle : Microsoft.Office.Interop.Word.Style, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Style Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Frames which adds IDispose to the interface
	/// </summary>
	public interface IFrames : Microsoft.Office.Interop.Word.Frames, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Frames Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Frame which adds IDispose to the interface
	/// </summary>
	public interface IFrame : Microsoft.Office.Interop.Word.Frame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Frame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormFields which adds IDispose to the interface
	/// </summary>
	public interface IFormFields : Microsoft.Office.Interop.Word.FormFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FormFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormField which adds IDispose to the interface
	/// </summary>
	public interface IFormField : Microsoft.Office.Interop.Word.FormField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FormField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextInput which adds IDispose to the interface
	/// </summary>
	public interface ITextInput : Microsoft.Office.Interop.Word.TextInput, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TextInput Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CheckBox which adds IDispose to the interface
	/// </summary>
	public interface ICheckBox : Microsoft.Office.Interop.Word.CheckBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CheckBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropDown which adds IDispose to the interface
	/// </summary>
	public interface IDropDown : Microsoft.Office.Interop.Word.DropDown, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DropDown Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListEntries which adds IDispose to the interface
	/// </summary>
	public interface IListEntries : Microsoft.Office.Interop.Word.ListEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListEntry which adds IDispose to the interface
	/// </summary>
	public interface IListEntry : Microsoft.Office.Interop.Word.ListEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TablesOfFigures which adds IDispose to the interface
	/// </summary>
	public interface ITablesOfFigures : Microsoft.Office.Interop.Word.TablesOfFigures, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TablesOfFigures Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableOfFigures which adds IDispose to the interface
	/// </summary>
	public interface ITableOfFigures : Microsoft.Office.Interop.Word.TableOfFigures, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TableOfFigures Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMerge which adds IDispose to the interface
	/// </summary>
	public interface IMailMerge : Microsoft.Office.Interop.Word.MailMerge, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMerge Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeFields which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeFields : Microsoft.Office.Interop.Word.MailMergeFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeField which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeField : Microsoft.Office.Interop.Word.MailMergeField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeDataSource which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeDataSource : Microsoft.Office.Interop.Word.MailMergeDataSource, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeDataSource Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeFieldNames which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeFieldNames : Microsoft.Office.Interop.Word.MailMergeFieldNames, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeFieldNames Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeFieldName which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeFieldName : Microsoft.Office.Interop.Word.MailMergeFieldName, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeFieldName Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeDataFields which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeDataFields : Microsoft.Office.Interop.Word.MailMergeDataFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeDataFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMergeDataField which adds IDispose to the interface
	/// </summary>
	public interface IMailMergeDataField : Microsoft.Office.Interop.Word.MailMergeDataField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMergeDataField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Envelope which adds IDispose to the interface
	/// </summary>
	public interface IEnvelope : Microsoft.Office.Interop.Word.Envelope, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Envelope Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailingLabel which adds IDispose to the interface
	/// </summary>
	public interface IMailingLabel : Microsoft.Office.Interop.Word.MailingLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailingLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomLabels which adds IDispose to the interface
	/// </summary>
	public interface ICustomLabels : Microsoft.Office.Interop.Word.CustomLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CustomLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomLabel which adds IDispose to the interface
	/// </summary>
	public interface ICustomLabel : Microsoft.Office.Interop.Word.CustomLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CustomLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TablesOfContents which adds IDispose to the interface
	/// </summary>
	public interface ITablesOfContents : Microsoft.Office.Interop.Word.TablesOfContents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TablesOfContents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableOfContents which adds IDispose to the interface
	/// </summary>
	public interface ITableOfContents : Microsoft.Office.Interop.Word.TableOfContents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TableOfContents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TablesOfAuthorities which adds IDispose to the interface
	/// </summary>
	public interface ITablesOfAuthorities : Microsoft.Office.Interop.Word.TablesOfAuthorities, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TablesOfAuthorities Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableOfAuthorities which adds IDispose to the interface
	/// </summary>
	public interface ITableOfAuthorities : Microsoft.Office.Interop.Word.TableOfAuthorities, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TableOfAuthorities Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Dialogs which adds IDispose to the interface
	/// </summary>
	public interface IDialogs : Microsoft.Office.Interop.Word.Dialogs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Dialogs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Dialog which adds IDispose to the interface
	/// </summary>
	public interface IDialog : Microsoft.Office.Interop.Word.Dialog, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Dialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PageSetup which adds IDispose to the interface
	/// </summary>
	public interface IPageSetup : Microsoft.Office.Interop.Word.PageSetup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.PageSetup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LineNumbering which adds IDispose to the interface
	/// </summary>
	public interface ILineNumbering : Microsoft.Office.Interop.Word.LineNumbering, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LineNumbering Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextColumns which adds IDispose to the interface
	/// </summary>
	public interface ITextColumns : Microsoft.Office.Interop.Word.TextColumns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TextColumns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextColumn which adds IDispose to the interface
	/// </summary>
	public interface ITextColumn : Microsoft.Office.Interop.Word.TextColumn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TextColumn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Selection which adds IDispose to the interface
	/// </summary>
	public interface ISelection : Microsoft.Office.Interop.Word.Selection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Selection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TablesOfAuthoritiesCategories which adds IDispose to the interface
	/// </summary>
	public interface ITablesOfAuthoritiesCategories : Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TablesOfAuthoritiesCategories Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableOfAuthoritiesCategory which adds IDispose to the interface
	/// </summary>
	public interface ITableOfAuthoritiesCategory : Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TableOfAuthoritiesCategory Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CaptionLabels which adds IDispose to the interface
	/// </summary>
	public interface ICaptionLabels : Microsoft.Office.Interop.Word.CaptionLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CaptionLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CaptionLabel which adds IDispose to the interface
	/// </summary>
	public interface ICaptionLabel : Microsoft.Office.Interop.Word.CaptionLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CaptionLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCaptions which adds IDispose to the interface
	/// </summary>
	public interface IAutoCaptions : Microsoft.Office.Interop.Word.AutoCaptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoCaptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCaption which adds IDispose to the interface
	/// </summary>
	public interface IAutoCaption : Microsoft.Office.Interop.Word.AutoCaption, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AutoCaption Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Indexes which adds IDispose to the interface
	/// </summary>
	public interface IIndexes : Microsoft.Office.Interop.Word.Indexes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Indexes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Index which adds IDispose to the interface
	/// </summary>
	public interface IIndex : Microsoft.Office.Interop.Word.Index, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Index Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIn which adds IDispose to the interface
	/// </summary>
	public interface IAddIn : Microsoft.Office.Interop.Word.AddIn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AddIn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIns which adds IDispose to the interface
	/// </summary>
	public interface IAddIns : Microsoft.Office.Interop.Word.AddIns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AddIns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Revisions which adds IDispose to the interface
	/// </summary>
	public interface IRevisions : Microsoft.Office.Interop.Word.Revisions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Revisions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Revision which adds IDispose to the interface
	/// </summary>
	public interface IRevision : Microsoft.Office.Interop.Word.Revision, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Revision Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Task which adds IDispose to the interface
	/// </summary>
	public interface ITask : Microsoft.Office.Interop.Word.Task, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Task Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Tasks which adds IDispose to the interface
	/// </summary>
	public interface ITasks : Microsoft.Office.Interop.Word.Tasks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Tasks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeadersFooters which adds IDispose to the interface
	/// </summary>
	public interface IHeadersFooters : Microsoft.Office.Interop.Word.HeadersFooters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HeadersFooters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeaderFooter which adds IDispose to the interface
	/// </summary>
	public interface IHeaderFooter : Microsoft.Office.Interop.Word.HeaderFooter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HeaderFooter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PageNumbers which adds IDispose to the interface
	/// </summary>
	public interface IPageNumbers : Microsoft.Office.Interop.Word.PageNumbers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.PageNumbers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PageNumber which adds IDispose to the interface
	/// </summary>
	public interface IPageNumber : Microsoft.Office.Interop.Word.PageNumber, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.PageNumber Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Subdocuments which adds IDispose to the interface
	/// </summary>
	public interface ISubdocuments : Microsoft.Office.Interop.Word.Subdocuments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Subdocuments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Subdocument which adds IDispose to the interface
	/// </summary>
	public interface ISubdocument : Microsoft.Office.Interop.Word.Subdocument, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Subdocument Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeadingStyles which adds IDispose to the interface
	/// </summary>
	public interface IHeadingStyles : Microsoft.Office.Interop.Word.HeadingStyles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HeadingStyles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeadingStyle which adds IDispose to the interface
	/// </summary>
	public interface IHeadingStyle : Microsoft.Office.Interop.Word.HeadingStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HeadingStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for StoryRanges which adds IDispose to the interface
	/// </summary>
	public interface IStoryRanges : Microsoft.Office.Interop.Word.StoryRanges, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.StoryRanges Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListLevel which adds IDispose to the interface
	/// </summary>
	public interface IListLevel : Microsoft.Office.Interop.Word.ListLevel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListLevel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListLevels which adds IDispose to the interface
	/// </summary>
	public interface IListLevels : Microsoft.Office.Interop.Word.ListLevels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListLevels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListTemplate which adds IDispose to the interface
	/// </summary>
	public interface IListTemplate : Microsoft.Office.Interop.Word.ListTemplate, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListTemplate Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListTemplates which adds IDispose to the interface
	/// </summary>
	public interface IListTemplates : Microsoft.Office.Interop.Word.ListTemplates, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListTemplates Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListParagraphs which adds IDispose to the interface
	/// </summary>
	public interface IListParagraphs : Microsoft.Office.Interop.Word.ListParagraphs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListParagraphs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for List which adds IDispose to the interface
	/// </summary>
	public interface IList : Microsoft.Office.Interop.Word.List, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.List Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Lists which adds IDispose to the interface
	/// </summary>
	public interface ILists : Microsoft.Office.Interop.Word.Lists, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Lists Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListGallery which adds IDispose to the interface
	/// </summary>
	public interface IListGallery : Microsoft.Office.Interop.Word.ListGallery, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListGallery Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListGalleries which adds IDispose to the interface
	/// </summary>
	public interface IListGalleries : Microsoft.Office.Interop.Word.ListGalleries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ListGalleries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for KeyBindings which adds IDispose to the interface
	/// </summary>
	public interface IKeyBindings : Microsoft.Office.Interop.Word.KeyBindings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.KeyBindings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for KeysBoundTo which adds IDispose to the interface
	/// </summary>
	public interface IKeysBoundTo : Microsoft.Office.Interop.Word.KeysBoundTo, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.KeysBoundTo Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for KeyBinding which adds IDispose to the interface
	/// </summary>
	public interface IKeyBinding : Microsoft.Office.Interop.Word.KeyBinding, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.KeyBinding Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileConverter which adds IDispose to the interface
	/// </summary>
	public interface IFileConverter : Microsoft.Office.Interop.Word.FileConverter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FileConverter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileConverters which adds IDispose to the interface
	/// </summary>
	public interface IFileConverters : Microsoft.Office.Interop.Word.FileConverters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FileConverters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SynonymInfo which adds IDispose to the interface
	/// </summary>
	public interface ISynonymInfo : Microsoft.Office.Interop.Word.SynonymInfo, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SynonymInfo Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Hyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IHyperlinks : Microsoft.Office.Interop.Word.Hyperlinks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Hyperlinks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Hyperlink which adds IDispose to the interface
	/// </summary>
	public interface IHyperlink : Microsoft.Office.Interop.Word.Hyperlink, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Hyperlink Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shapes which adds IDispose to the interface
	/// </summary>
	public interface IShapes : Microsoft.Office.Interop.Word.Shapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Shapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IShapeRange : Microsoft.Office.Interop.Word.ShapeRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ShapeRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IGroupShapes : Microsoft.Office.Interop.Word.GroupShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.GroupShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shape which adds IDispose to the interface
	/// </summary>
	public interface IShape : Microsoft.Office.Interop.Word.Shape, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Shape Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame : Microsoft.Office.Interop.Word.TextFrame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TextFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _LetterContent which adds IDispose to the interface
	/// </summary>
	public interface I_LetterContent : Microsoft.Office.Interop.Word._LetterContent, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word._LetterContent Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for View which adds IDispose to the interface
	/// </summary>
	public interface IView : Microsoft.Office.Interop.Word.View, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.View Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Zoom which adds IDispose to the interface
	/// </summary>
	public interface IZoom : Microsoft.Office.Interop.Word.Zoom, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Zoom Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Zooms which adds IDispose to the interface
	/// </summary>
	public interface IZooms : Microsoft.Office.Interop.Word.Zooms, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Zooms Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InlineShape which adds IDispose to the interface
	/// </summary>
	public interface IInlineShape : Microsoft.Office.Interop.Word.InlineShape, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.InlineShape Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for InlineShapes which adds IDispose to the interface
	/// </summary>
	public interface IInlineShapes : Microsoft.Office.Interop.Word.InlineShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.InlineShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SpellingSuggestions which adds IDispose to the interface
	/// </summary>
	public interface ISpellingSuggestions : Microsoft.Office.Interop.Word.SpellingSuggestions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SpellingSuggestions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SpellingSuggestion which adds IDispose to the interface
	/// </summary>
	public interface ISpellingSuggestion : Microsoft.Office.Interop.Word.SpellingSuggestion, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SpellingSuggestion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Dictionaries which adds IDispose to the interface
	/// </summary>
	public interface IDictionaries : Microsoft.Office.Interop.Word.Dictionaries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Dictionaries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HangulHanjaConversionDictionaries which adds IDispose to the interface
	/// </summary>
	public interface IHangulHanjaConversionDictionaries : Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HangulHanjaConversionDictionaries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Dictionary which adds IDispose to the interface
	/// </summary>
	public interface IDictionary : Microsoft.Office.Interop.Word.Dictionary, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Dictionary Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReadabilityStatistics which adds IDispose to the interface
	/// </summary>
	public interface IReadabilityStatistics : Microsoft.Office.Interop.Word.ReadabilityStatistics, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ReadabilityStatistics Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReadabilityStatistic which adds IDispose to the interface
	/// </summary>
	public interface IReadabilityStatistic : Microsoft.Office.Interop.Word.ReadabilityStatistic, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ReadabilityStatistic Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Versions which adds IDispose to the interface
	/// </summary>
	public interface IVersions : Microsoft.Office.Interop.Word.Versions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Versions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Version which adds IDispose to the interface
	/// </summary>
	public interface IVersion : Microsoft.Office.Interop.Word.Version, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Version Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Options which adds IDispose to the interface
	/// </summary>
	public interface IOptions : Microsoft.Office.Interop.Word.Options, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Options Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MailMessage which adds IDispose to the interface
	/// </summary>
	public interface IMailMessage : Microsoft.Office.Interop.Word.MailMessage, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MailMessage Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProofreadingErrors which adds IDispose to the interface
	/// </summary>
	public interface IProofreadingErrors : Microsoft.Office.Interop.Word.ProofreadingErrors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ProofreadingErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Mailer which adds IDispose to the interface
	/// </summary>
	public interface IMailer : Microsoft.Office.Interop.Word.Mailer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Mailer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WrapFormat which adds IDispose to the interface
	/// </summary>
	public interface IWrapFormat : Microsoft.Office.Interop.Word.WrapFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.WrapFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HangulAndAlphabetExceptions which adds IDispose to the interface
	/// </summary>
	public interface IHangulAndAlphabetExceptions : Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HangulAndAlphabetExceptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HangulAndAlphabetException which adds IDispose to the interface
	/// </summary>
	public interface IHangulAndAlphabetException : Microsoft.Office.Interop.Word.HangulAndAlphabetException, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HangulAndAlphabetException Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Adjustments which adds IDispose to the interface
	/// </summary>
	public interface IAdjustments : Microsoft.Office.Interop.Word.Adjustments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Adjustments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalloutFormat which adds IDispose to the interface
	/// </summary>
	public interface ICalloutFormat : Microsoft.Office.Interop.Word.CalloutFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CalloutFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IColorFormat : Microsoft.Office.Interop.Word.ColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IConnectorFormat : Microsoft.Office.Interop.Word.ConnectorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ConnectorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FillFormat which adds IDispose to the interface
	/// </summary>
	public interface IFillFormat : Microsoft.Office.Interop.Word.FillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IFreeformBuilder : Microsoft.Office.Interop.Word.FreeformBuilder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FreeformBuilder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LineFormat which adds IDispose to the interface
	/// </summary>
	public interface ILineFormat : Microsoft.Office.Interop.Word.LineFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LineFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PictureFormat which adds IDispose to the interface
	/// </summary>
	public interface IPictureFormat : Microsoft.Office.Interop.Word.PictureFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.PictureFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShadowFormat which adds IDispose to the interface
	/// </summary>
	public interface IShadowFormat : Microsoft.Office.Interop.Word.ShadowFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ShadowFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNode which adds IDispose to the interface
	/// </summary>
	public interface IShapeNode : Microsoft.Office.Interop.Word.ShapeNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ShapeNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNodes which adds IDispose to the interface
	/// </summary>
	public interface IShapeNodes : Microsoft.Office.Interop.Word.ShapeNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ShapeNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
	/// </summary>
	public interface ITextEffectFormat : Microsoft.Office.Interop.Word.TextEffectFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TextEffectFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
	/// </summary>
	public interface IThreeDFormat : Microsoft.Office.Interop.Word.ThreeDFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ThreeDFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents : Microsoft.Office.Interop.Word.ApplicationEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Global which adds IDispose to the interface
	/// </summary>
	public interface IGlobal : Microsoft.Office.Interop.Word.Global, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Global Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents_Event : Microsoft.Office.Interop.Word.ApplicationEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents2_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents2_Event : Microsoft.Office.Interop.Word.ApplicationEvents2_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents2_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents3_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents3_Event : Microsoft.Office.Interop.Word.ApplicationEvents3_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents3_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents4_Event which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents4_Event : Microsoft.Office.Interop.Word.ApplicationEvents4_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents4_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Application which adds IDispose to the interface
	/// </summary>
	public interface IApplication : Microsoft.Office.Interop.Word.Application, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentEvents which adds IDispose to the interface
	/// </summary>
	public interface IDocumentEvents : Microsoft.Office.Interop.Word.DocumentEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DocumentEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IDocumentEvents_Event : Microsoft.Office.Interop.Word.DocumentEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DocumentEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentEvents2_Event which adds IDispose to the interface
	/// </summary>
	public interface IDocumentEvents2_Event : Microsoft.Office.Interop.Word.DocumentEvents2_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DocumentEvents2_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Document which adds IDispose to the interface
	/// </summary>
	public interface IDocument : Microsoft.Office.Interop.Word.Document, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Document Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Font which adds IDispose to the interface
	/// </summary>
	public interface IFont : Microsoft.Office.Interop.Word.Font, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Font Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
	/// </summary>
	public interface IParagraphFormat : Microsoft.Office.Interop.Word.ParagraphFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ParagraphFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OCXEvents which adds IDispose to the interface
	/// </summary>
	public interface IOCXEvents : Microsoft.Office.Interop.Word.OCXEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OCXEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OCXEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOCXEvents_Event : Microsoft.Office.Interop.Word.OCXEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OCXEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEControl which adds IDispose to the interface
	/// </summary>
	public interface IOLEControl : Microsoft.Office.Interop.Word.OLEControl, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OLEControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LetterContent which adds IDispose to the interface
	/// </summary>
	public interface ILetterContent : Microsoft.Office.Interop.Word.LetterContent, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LetterContent Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IApplicationEvents which adds IDispose to the interface
	/// </summary>
	public interface IIApplicationEvents : Microsoft.Office.Interop.Word.IApplicationEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.IApplicationEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IApplicationEvents2 which adds IDispose to the interface
	/// </summary>
	public interface IIApplicationEvents2 : Microsoft.Office.Interop.Word.IApplicationEvents2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.IApplicationEvents2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents2 which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents2 : Microsoft.Office.Interop.Word.ApplicationEvents2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EmailAuthor which adds IDispose to the interface
	/// </summary>
	public interface IEmailAuthor : Microsoft.Office.Interop.Word.EmailAuthor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.EmailAuthor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EmailOptions which adds IDispose to the interface
	/// </summary>
	public interface IEmailOptions : Microsoft.Office.Interop.Word.EmailOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.EmailOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EmailSignature which adds IDispose to the interface
	/// </summary>
	public interface IEmailSignature : Microsoft.Office.Interop.Word.EmailSignature, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.EmailSignature Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Email which adds IDispose to the interface
	/// </summary>
	public interface IEmail : Microsoft.Office.Interop.Word.Email, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Email Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HorizontalLineFormat which adds IDispose to the interface
	/// </summary>
	public interface IHorizontalLineFormat : Microsoft.Office.Interop.Word.HorizontalLineFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HorizontalLineFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Frameset which adds IDispose to the interface
	/// </summary>
	public interface IFrameset : Microsoft.Office.Interop.Word.Frameset, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Frameset Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
	/// </summary>
	public interface IDefaultWebOptions : Microsoft.Office.Interop.Word.DefaultWebOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DefaultWebOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebOptions which adds IDispose to the interface
	/// </summary>
	public interface IWebOptions : Microsoft.Office.Interop.Word.WebOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.WebOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OtherCorrectionsExceptions which adds IDispose to the interface
	/// </summary>
	public interface IOtherCorrectionsExceptions : Microsoft.Office.Interop.Word.OtherCorrectionsExceptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OtherCorrectionsExceptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OtherCorrectionsException which adds IDispose to the interface
	/// </summary>
	public interface IOtherCorrectionsException : Microsoft.Office.Interop.Word.OtherCorrectionsException, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OtherCorrectionsException Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EmailSignatureEntries which adds IDispose to the interface
	/// </summary>
	public interface IEmailSignatureEntries : Microsoft.Office.Interop.Word.EmailSignatureEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.EmailSignatureEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EmailSignatureEntry which adds IDispose to the interface
	/// </summary>
	public interface IEmailSignatureEntry : Microsoft.Office.Interop.Word.EmailSignatureEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.EmailSignatureEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HTMLDivision which adds IDispose to the interface
	/// </summary>
	public interface IHTMLDivision : Microsoft.Office.Interop.Word.HTMLDivision, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HTMLDivision Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HTMLDivisions which adds IDispose to the interface
	/// </summary>
	public interface IHTMLDivisions : Microsoft.Office.Interop.Word.HTMLDivisions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HTMLDivisions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNode which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNode : Microsoft.Office.Interop.Word.DiagramNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DiagramNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodeChildren : Microsoft.Office.Interop.Word.DiagramNodeChildren, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DiagramNodeChildren Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodes which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodes : Microsoft.Office.Interop.Word.DiagramNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DiagramNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Diagram which adds IDispose to the interface
	/// </summary>
	public interface IDiagram : Microsoft.Office.Interop.Word.Diagram, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Diagram Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomProperty which adds IDispose to the interface
	/// </summary>
	public interface ICustomProperty : Microsoft.Office.Interop.Word.CustomProperty, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CustomProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomProperties which adds IDispose to the interface
	/// </summary>
	public interface ICustomProperties : Microsoft.Office.Interop.Word.CustomProperties, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CustomProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTag which adds IDispose to the interface
	/// </summary>
	public interface ISmartTag : Microsoft.Office.Interop.Word.SmartTag, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTag Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTags which adds IDispose to the interface
	/// </summary>
	public interface ISmartTags : Microsoft.Office.Interop.Word.SmartTags, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTags Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for StyleSheet which adds IDispose to the interface
	/// </summary>
	public interface IStyleSheet : Microsoft.Office.Interop.Word.StyleSheet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.StyleSheet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for StyleSheets which adds IDispose to the interface
	/// </summary>
	public interface IStyleSheets : Microsoft.Office.Interop.Word.StyleSheets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.StyleSheets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MappedDataField which adds IDispose to the interface
	/// </summary>
	public interface IMappedDataField : Microsoft.Office.Interop.Word.MappedDataField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MappedDataField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MappedDataFields which adds IDispose to the interface
	/// </summary>
	public interface IMappedDataFields : Microsoft.Office.Interop.Word.MappedDataFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.MappedDataFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CanvasShapes which adds IDispose to the interface
	/// </summary>
	public interface ICanvasShapes : Microsoft.Office.Interop.Word.CanvasShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CanvasShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableStyle which adds IDispose to the interface
	/// </summary>
	public interface ITableStyle : Microsoft.Office.Interop.Word.TableStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TableStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ConditionalStyle which adds IDispose to the interface
	/// </summary>
	public interface IConditionalStyle : Microsoft.Office.Interop.Word.ConditionalStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ConditionalStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FootnoteOptions which adds IDispose to the interface
	/// </summary>
	public interface IFootnoteOptions : Microsoft.Office.Interop.Word.FootnoteOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.FootnoteOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EndnoteOptions which adds IDispose to the interface
	/// </summary>
	public interface IEndnoteOptions : Microsoft.Office.Interop.Word.EndnoteOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.EndnoteOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Reviewers which adds IDispose to the interface
	/// </summary>
	public interface IReviewers : Microsoft.Office.Interop.Word.Reviewers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Reviewers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Reviewer which adds IDispose to the interface
	/// </summary>
	public interface IReviewer : Microsoft.Office.Interop.Word.Reviewer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Reviewer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskPane which adds IDispose to the interface
	/// </summary>
	public interface ITaskPane : Microsoft.Office.Interop.Word.TaskPane, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TaskPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TaskPanes which adds IDispose to the interface
	/// </summary>
	public interface ITaskPanes : Microsoft.Office.Interop.Word.TaskPanes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TaskPanes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IApplicationEvents3 which adds IDispose to the interface
	/// </summary>
	public interface IIApplicationEvents3 : Microsoft.Office.Interop.Word.IApplicationEvents3, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.IApplicationEvents3 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents3 which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents3 : Microsoft.Office.Interop.Word.ApplicationEvents3, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents3 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagAction which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagAction : Microsoft.Office.Interop.Word.SmartTagAction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTagAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagActions which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagActions : Microsoft.Office.Interop.Word.SmartTagActions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTagActions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagRecognizer : Microsoft.Office.Interop.Word.SmartTagRecognizer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTagRecognizer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagRecognizers : Microsoft.Office.Interop.Word.SmartTagRecognizers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTagRecognizers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagType which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagType : Microsoft.Office.Interop.Word.SmartTagType, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTagType Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagTypes which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagTypes : Microsoft.Office.Interop.Word.SmartTagTypes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SmartTagTypes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Line which adds IDispose to the interface
	/// </summary>
	public interface ILine : Microsoft.Office.Interop.Word.Line, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Line Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Lines which adds IDispose to the interface
	/// </summary>
	public interface ILines : Microsoft.Office.Interop.Word.Lines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Lines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rectangle which adds IDispose to the interface
	/// </summary>
	public interface IRectangle : Microsoft.Office.Interop.Word.Rectangle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Rectangle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rectangles which adds IDispose to the interface
	/// </summary>
	public interface IRectangles : Microsoft.Office.Interop.Word.Rectangles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Rectangles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Break which adds IDispose to the interface
	/// </summary>
	public interface IBreak : Microsoft.Office.Interop.Word.Break, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Break Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Breaks which adds IDispose to the interface
	/// </summary>
	public interface IBreaks : Microsoft.Office.Interop.Word.Breaks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Breaks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Page which adds IDispose to the interface
	/// </summary>
	public interface IPage : Microsoft.Office.Interop.Word.Page, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Page Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pages which adds IDispose to the interface
	/// </summary>
	public interface IPages : Microsoft.Office.Interop.Word.Pages, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Pages Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLNode which adds IDispose to the interface
	/// </summary>
	public interface IXMLNode : Microsoft.Office.Interop.Word.XMLNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLNodes which adds IDispose to the interface
	/// </summary>
	public interface IXMLNodes : Microsoft.Office.Interop.Word.XMLNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLSchemaReference which adds IDispose to the interface
	/// </summary>
	public interface IXMLSchemaReference : Microsoft.Office.Interop.Word.XMLSchemaReference, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLSchemaReference Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLSchemaReferences which adds IDispose to the interface
	/// </summary>
	public interface IXMLSchemaReferences : Microsoft.Office.Interop.Word.XMLSchemaReferences, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLSchemaReferences Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLChildNodeSuggestion which adds IDispose to the interface
	/// </summary>
	public interface IXMLChildNodeSuggestion : Microsoft.Office.Interop.Word.XMLChildNodeSuggestion, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLChildNodeSuggestion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLChildNodeSuggestions which adds IDispose to the interface
	/// </summary>
	public interface IXMLChildNodeSuggestions : Microsoft.Office.Interop.Word.XMLChildNodeSuggestions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLChildNodeSuggestions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLNamespace which adds IDispose to the interface
	/// </summary>
	public interface IXMLNamespace : Microsoft.Office.Interop.Word.XMLNamespace, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLNamespace Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLNamespaces which adds IDispose to the interface
	/// </summary>
	public interface IXMLNamespaces : Microsoft.Office.Interop.Word.XMLNamespaces, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLNamespaces Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XSLTransform which adds IDispose to the interface
	/// </summary>
	public interface IXSLTransform : Microsoft.Office.Interop.Word.XSLTransform, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XSLTransform Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XSLTransforms which adds IDispose to the interface
	/// </summary>
	public interface IXSLTransforms : Microsoft.Office.Interop.Word.XSLTransforms, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XSLTransforms Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Editors which adds IDispose to the interface
	/// </summary>
	public interface IEditors : Microsoft.Office.Interop.Word.Editors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Editors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Editor which adds IDispose to the interface
	/// </summary>
	public interface IEditor : Microsoft.Office.Interop.Word.Editor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Editor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IApplicationEvents4 which adds IDispose to the interface
	/// </summary>
	public interface IIApplicationEvents4 : Microsoft.Office.Interop.Word.IApplicationEvents4, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.IApplicationEvents4 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ApplicationEvents4 which adds IDispose to the interface
	/// </summary>
	public interface IApplicationEvents4 : Microsoft.Office.Interop.Word.ApplicationEvents4, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ApplicationEvents4 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentEvents2 which adds IDispose to the interface
	/// </summary>
	public interface IDocumentEvents2 : Microsoft.Office.Interop.Word.DocumentEvents2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DocumentEvents2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Source which adds IDispose to the interface
	/// </summary>
	public interface ISource : Microsoft.Office.Interop.Word.Source, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Source Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sources which adds IDispose to the interface
	/// </summary>
	public interface ISources : Microsoft.Office.Interop.Word.Sources, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Sources Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Bibliography which adds IDispose to the interface
	/// </summary>
	public interface IBibliography : Microsoft.Office.Interop.Word.Bibliography, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Bibliography Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMaths which adds IDispose to the interface
	/// </summary>
	public interface IOMaths : Microsoft.Office.Interop.Word.OMaths, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMaths Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMath which adds IDispose to the interface
	/// </summary>
	public interface IOMath : Microsoft.Office.Interop.Word.OMath, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMath Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathFunctions which adds IDispose to the interface
	/// </summary>
	public interface IOMathFunctions : Microsoft.Office.Interop.Word.OMathFunctions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathFunctions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathArgs which adds IDispose to the interface
	/// </summary>
	public interface IOMathArgs : Microsoft.Office.Interop.Word.OMathArgs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathArgs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathFunction which adds IDispose to the interface
	/// </summary>
	public interface IOMathFunction : Microsoft.Office.Interop.Word.OMathFunction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathFunction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathAcc which adds IDispose to the interface
	/// </summary>
	public interface IOMathAcc : Microsoft.Office.Interop.Word.OMathAcc, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathAcc Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathBar which adds IDispose to the interface
	/// </summary>
	public interface IOMathBar : Microsoft.Office.Interop.Word.OMathBar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathBox which adds IDispose to the interface
	/// </summary>
	public interface IOMathBox : Microsoft.Office.Interop.Word.OMathBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathBorderBox which adds IDispose to the interface
	/// </summary>
	public interface IOMathBorderBox : Microsoft.Office.Interop.Word.OMathBorderBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathBorderBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathDelim which adds IDispose to the interface
	/// </summary>
	public interface IOMathDelim : Microsoft.Office.Interop.Word.OMathDelim, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathDelim Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathEqArray which adds IDispose to the interface
	/// </summary>
	public interface IOMathEqArray : Microsoft.Office.Interop.Word.OMathEqArray, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathEqArray Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathFrac which adds IDispose to the interface
	/// </summary>
	public interface IOMathFrac : Microsoft.Office.Interop.Word.OMathFrac, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathFrac Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathFunc which adds IDispose to the interface
	/// </summary>
	public interface IOMathFunc : Microsoft.Office.Interop.Word.OMathFunc, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathFunc Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathGroupChar which adds IDispose to the interface
	/// </summary>
	public interface IOMathGroupChar : Microsoft.Office.Interop.Word.OMathGroupChar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathGroupChar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathLimLow which adds IDispose to the interface
	/// </summary>
	public interface IOMathLimLow : Microsoft.Office.Interop.Word.OMathLimLow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathLimLow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathLimUpp which adds IDispose to the interface
	/// </summary>
	public interface IOMathLimUpp : Microsoft.Office.Interop.Word.OMathLimUpp, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathLimUpp Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathMat which adds IDispose to the interface
	/// </summary>
	public interface IOMathMat : Microsoft.Office.Interop.Word.OMathMat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathMat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathMatRows which adds IDispose to the interface
	/// </summary>
	public interface IOMathMatRows : Microsoft.Office.Interop.Word.OMathMatRows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathMatRows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathMatCols which adds IDispose to the interface
	/// </summary>
	public interface IOMathMatCols : Microsoft.Office.Interop.Word.OMathMatCols, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathMatCols Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathMatRow which adds IDispose to the interface
	/// </summary>
	public interface IOMathMatRow : Microsoft.Office.Interop.Word.OMathMatRow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathMatRow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathMatCol which adds IDispose to the interface
	/// </summary>
	public interface IOMathMatCol : Microsoft.Office.Interop.Word.OMathMatCol, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathMatCol Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathNary which adds IDispose to the interface
	/// </summary>
	public interface IOMathNary : Microsoft.Office.Interop.Word.OMathNary, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathNary Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathPhantom which adds IDispose to the interface
	/// </summary>
	public interface IOMathPhantom : Microsoft.Office.Interop.Word.OMathPhantom, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathPhantom Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathScrPre which adds IDispose to the interface
	/// </summary>
	public interface IOMathScrPre : Microsoft.Office.Interop.Word.OMathScrPre, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathScrPre Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathRad which adds IDispose to the interface
	/// </summary>
	public interface IOMathRad : Microsoft.Office.Interop.Word.OMathRad, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathRad Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathScrSub which adds IDispose to the interface
	/// </summary>
	public interface IOMathScrSub : Microsoft.Office.Interop.Word.OMathScrSub, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathScrSub Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathScrSubSup which adds IDispose to the interface
	/// </summary>
	public interface IOMathScrSubSup : Microsoft.Office.Interop.Word.OMathScrSubSup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathScrSubSup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathScrSup which adds IDispose to the interface
	/// </summary>
	public interface IOMathScrSup : Microsoft.Office.Interop.Word.OMathScrSup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathScrSup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathAutoCorrect which adds IDispose to the interface
	/// </summary>
	public interface IOMathAutoCorrect : Microsoft.Office.Interop.Word.OMathAutoCorrect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathAutoCorrect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathAutoCorrectEntries which adds IDispose to the interface
	/// </summary>
	public interface IOMathAutoCorrectEntries : Microsoft.Office.Interop.Word.OMathAutoCorrectEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathAutoCorrectEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathAutoCorrectEntry which adds IDispose to the interface
	/// </summary>
	public interface IOMathAutoCorrectEntry : Microsoft.Office.Interop.Word.OMathAutoCorrectEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathAutoCorrectEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathRecognizedFunctions which adds IDispose to the interface
	/// </summary>
	public interface IOMathRecognizedFunctions : Microsoft.Office.Interop.Word.OMathRecognizedFunctions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathRecognizedFunctions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathRecognizedFunction which adds IDispose to the interface
	/// </summary>
	public interface IOMathRecognizedFunction : Microsoft.Office.Interop.Word.OMathRecognizedFunction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathRecognizedFunction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContentControls which adds IDispose to the interface
	/// </summary>
	public interface IContentControls : Microsoft.Office.Interop.Word.ContentControls, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ContentControls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContentControl which adds IDispose to the interface
	/// </summary>
	public interface IContentControl : Microsoft.Office.Interop.Word.ContentControl, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ContentControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XMLMapping which adds IDispose to the interface
	/// </summary>
	public interface IXMLMapping : Microsoft.Office.Interop.Word.XMLMapping, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.XMLMapping Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContentControlListEntries which adds IDispose to the interface
	/// </summary>
	public interface IContentControlListEntries : Microsoft.Office.Interop.Word.ContentControlListEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ContentControlListEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ContentControlListEntry which adds IDispose to the interface
	/// </summary>
	public interface IContentControlListEntry : Microsoft.Office.Interop.Word.ContentControlListEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ContentControlListEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BuildingBlockTypes which adds IDispose to the interface
	/// </summary>
	public interface IBuildingBlockTypes : Microsoft.Office.Interop.Word.BuildingBlockTypes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.BuildingBlockTypes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BuildingBlockType which adds IDispose to the interface
	/// </summary>
	public interface IBuildingBlockType : Microsoft.Office.Interop.Word.BuildingBlockType, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.BuildingBlockType Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Categories which adds IDispose to the interface
	/// </summary>
	public interface ICategories : Microsoft.Office.Interop.Word.Categories, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Categories Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Category which adds IDispose to the interface
	/// </summary>
	public interface ICategory : Microsoft.Office.Interop.Word.Category, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Category Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BuildingBlocks which adds IDispose to the interface
	/// </summary>
	public interface IBuildingBlocks : Microsoft.Office.Interop.Word.BuildingBlocks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.BuildingBlocks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BuildingBlock which adds IDispose to the interface
	/// </summary>
	public interface IBuildingBlock : Microsoft.Office.Interop.Word.BuildingBlock, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.BuildingBlock Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BuildingBlockEntries which adds IDispose to the interface
	/// </summary>
	public interface IBuildingBlockEntries : Microsoft.Office.Interop.Word.BuildingBlockEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.BuildingBlockEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathBreaks which adds IDispose to the interface
	/// </summary>
	public interface IOMathBreaks : Microsoft.Office.Interop.Word.OMathBreaks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathBreaks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OMathBreak which adds IDispose to the interface
	/// </summary>
	public interface IOMathBreak : Microsoft.Office.Interop.Word.OMathBreak, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.OMathBreak Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Research which adds IDispose to the interface
	/// </summary>
	public interface IResearch : Microsoft.Office.Interop.Word.Research, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Research Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SoftEdgeFormat which adds IDispose to the interface
	/// </summary>
	public interface ISoftEdgeFormat : Microsoft.Office.Interop.Word.SoftEdgeFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SoftEdgeFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GlowFormat which adds IDispose to the interface
	/// </summary>
	public interface IGlowFormat : Microsoft.Office.Interop.Word.GlowFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.GlowFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ReflectionFormat which adds IDispose to the interface
	/// </summary>
	public interface IReflectionFormat : Microsoft.Office.Interop.Word.ReflectionFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ReflectionFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartData which adds IDispose to the interface
	/// </summary>
	public interface IChartData : Microsoft.Office.Interop.Word.ChartData, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartData Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Chart which adds IDispose to the interface
	/// </summary>
	public interface IChart : Microsoft.Office.Interop.Word.Chart, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Chart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Corners which adds IDispose to the interface
	/// </summary>
	public interface ICorners : Microsoft.Office.Interop.Word.Corners, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Corners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Legend which adds IDispose to the interface
	/// </summary>
	public interface ILegend : Microsoft.Office.Interop.Word.Legend, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Legend Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartBorder which adds IDispose to the interface
	/// </summary>
	public interface IChartBorder : Microsoft.Office.Interop.Word.ChartBorder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartBorder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Walls which adds IDispose to the interface
	/// </summary>
	public interface IWalls : Microsoft.Office.Interop.Word.Walls, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Walls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Floor which adds IDispose to the interface
	/// </summary>
	public interface IFloor : Microsoft.Office.Interop.Word.Floor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Floor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PlotArea which adds IDispose to the interface
	/// </summary>
	public interface IPlotArea : Microsoft.Office.Interop.Word.PlotArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.PlotArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartArea which adds IDispose to the interface
	/// </summary>
	public interface IChartArea : Microsoft.Office.Interop.Word.ChartArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesLines which adds IDispose to the interface
	/// </summary>
	public interface ISeriesLines : Microsoft.Office.Interop.Word.SeriesLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SeriesLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LeaderLines which adds IDispose to the interface
	/// </summary>
	public interface ILeaderLines : Microsoft.Office.Interop.Word.LeaderLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LeaderLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Gridlines which adds IDispose to the interface
	/// </summary>
	public interface IGridlines : Microsoft.Office.Interop.Word.Gridlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Gridlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UpBars which adds IDispose to the interface
	/// </summary>
	public interface IUpBars : Microsoft.Office.Interop.Word.UpBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.UpBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DownBars which adds IDispose to the interface
	/// </summary>
	public interface IDownBars : Microsoft.Office.Interop.Word.DownBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DownBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Interior which adds IDispose to the interface
	/// </summary>
	public interface IInterior : Microsoft.Office.Interop.Word.Interior, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Interior Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFillFormat : Microsoft.Office.Interop.Word.ChartFillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartFillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntries which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntries : Microsoft.Office.Interop.Word.LegendEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LegendEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFont which adds IDispose to the interface
	/// </summary>
	public interface IChartFont : Microsoft.Office.Interop.Word.ChartFont, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartColorFormat : Microsoft.Office.Interop.Word.ChartColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntry which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntry : Microsoft.Office.Interop.Word.LegendEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LegendEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendKey which adds IDispose to the interface
	/// </summary>
	public interface ILegendKey : Microsoft.Office.Interop.Word.LegendKey, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.LegendKey Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface ISeriesCollection : Microsoft.Office.Interop.Word.SeriesCollection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.SeriesCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Series which adds IDispose to the interface
	/// </summary>
	public interface ISeries : Microsoft.Office.Interop.Word.Series, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Series Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IErrorBars : Microsoft.Office.Interop.Word.ErrorBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ErrorBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendline which adds IDispose to the interface
	/// </summary>
	public interface ITrendline : Microsoft.Office.Interop.Word.Trendline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Trendline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendlines which adds IDispose to the interface
	/// </summary>
	public interface ITrendlines : Microsoft.Office.Interop.Word.Trendlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Trendlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataLabels which adds IDispose to the interface
	/// </summary>
	public interface IDataLabels : Microsoft.Office.Interop.Word.DataLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DataLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataLabel which adds IDispose to the interface
	/// </summary>
	public interface IDataLabel : Microsoft.Office.Interop.Word.DataLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DataLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Points which adds IDispose to the interface
	/// </summary>
	public interface IPoints : Microsoft.Office.Interop.Word.Points, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Points Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Point which adds IDispose to the interface
	/// </summary>
	public interface IPoint : Microsoft.Office.Interop.Word.Point, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Point Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axes which adds IDispose to the interface
	/// </summary>
	public interface IAxes : Microsoft.Office.Interop.Word.Axes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Axes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axis which adds IDispose to the interface
	/// </summary>
	public interface IAxis : Microsoft.Office.Interop.Word.Axis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Axis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataTable which adds IDispose to the interface
	/// </summary>
	public interface IDataTable : Microsoft.Office.Interop.Word.DataTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DataTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IChartTitle : Microsoft.Office.Interop.Word.ChartTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IAxisTitle : Microsoft.Office.Interop.Word.AxisTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.AxisTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IDisplayUnitLabel : Microsoft.Office.Interop.Word.DisplayUnitLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DisplayUnitLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TickLabels which adds IDispose to the interface
	/// </summary>
	public interface ITickLabels : Microsoft.Office.Interop.Word.TickLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.TickLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropLines which adds IDispose to the interface
	/// </summary>
	public interface IDropLines : Microsoft.Office.Interop.Word.DropLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.DropLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IHiLoLines : Microsoft.Office.Interop.Word.HiLoLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.HiLoLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IChartGroup : Microsoft.Office.Interop.Word.ChartGroup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IChartGroups : Microsoft.Office.Interop.Word.ChartGroups, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartCharacters which adds IDispose to the interface
	/// </summary>
	public interface IChartCharacters : Microsoft.Office.Interop.Word.ChartCharacters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartCharacters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFormat : Microsoft.Office.Interop.Word.ChartFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ChartFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UndoRecord which adds IDispose to the interface
	/// </summary>
	public interface IUndoRecord : Microsoft.Office.Interop.Word.UndoRecord, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.UndoRecord Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthLock which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthLock : Microsoft.Office.Interop.Word.CoAuthLock, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthLock Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthLocks which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthLocks : Microsoft.Office.Interop.Word.CoAuthLocks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthLocks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthUpdate which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthUpdate : Microsoft.Office.Interop.Word.CoAuthUpdate, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthUpdate Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthUpdates which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthUpdates : Microsoft.Office.Interop.Word.CoAuthUpdates, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthUpdates Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthor which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthor : Microsoft.Office.Interop.Word.CoAuthor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthors which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthors : Microsoft.Office.Interop.Word.CoAuthors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CoAuthoring which adds IDispose to the interface
	/// </summary>
	public interface ICoAuthoring : Microsoft.Office.Interop.Word.CoAuthoring, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.CoAuthoring Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Conflicts which adds IDispose to the interface
	/// </summary>
	public interface IConflicts : Microsoft.Office.Interop.Word.Conflicts, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Conflicts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Conflict which adds IDispose to the interface
	/// </summary>
	public interface IConflict : Microsoft.Office.Interop.Word.Conflict, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.Conflict Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
	/// </summary>
	public interface IProtectedViewWindows : Microsoft.Office.Interop.Word.ProtectedViewWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ProtectedViewWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
	/// </summary>
	public interface IProtectedViewWindow : Microsoft.Office.Interop.Word.ProtectedViewWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Word.ProtectedViewWindow Resource { get; }
	}

	}