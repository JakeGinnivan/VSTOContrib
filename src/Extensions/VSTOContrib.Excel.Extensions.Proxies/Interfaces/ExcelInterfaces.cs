//Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.Excel.Extensions.Interfaces
{
	/// <summary>
	/// Wrapper interface for Adjustments which adds IDispose to the interface
	/// </summary>
	public interface IAdjustments : Microsoft.Office.Interop.Excel.Adjustments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Adjustments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalloutFormat which adds IDispose to the interface
	/// </summary>
	public interface ICalloutFormat : Microsoft.Office.Interop.Excel.CalloutFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CalloutFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IColorFormat : Microsoft.Office.Interop.Excel.ColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LineFormat which adds IDispose to the interface
	/// </summary>
	public interface ILineFormat : Microsoft.Office.Interop.Excel.LineFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LineFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNode which adds IDispose to the interface
	/// </summary>
	public interface IShapeNode : Microsoft.Office.Interop.Excel.ShapeNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ShapeNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNodes which adds IDispose to the interface
	/// </summary>
	public interface IShapeNodes : Microsoft.Office.Interop.Excel.ShapeNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ShapeNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PictureFormat which adds IDispose to the interface
	/// </summary>
	public interface IPictureFormat : Microsoft.Office.Interop.Excel.PictureFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PictureFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShadowFormat which adds IDispose to the interface
	/// </summary>
	public interface IShadowFormat : Microsoft.Office.Interop.Excel.ShadowFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ShadowFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
	/// </summary>
	public interface ITextEffectFormat : Microsoft.Office.Interop.Excel.TextEffectFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TextEffectFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
	/// </summary>
	public interface IThreeDFormat : Microsoft.Office.Interop.Excel.ThreeDFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ThreeDFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FillFormat which adds IDispose to the interface
	/// </summary>
	public interface IFillFormat : Microsoft.Office.Interop.Excel.FillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodes which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodes : Microsoft.Office.Interop.Excel.DiagramNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DiagramNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodeChildren : Microsoft.Office.Interop.Excel.DiagramNodeChildren, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DiagramNodeChildren Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNode which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNode : Microsoft.Office.Interop.Excel.DiagramNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DiagramNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRTDUpdateEvent which adds IDispose to the interface
	/// </summary>
	public interface IIRTDUpdateEvent : Microsoft.Office.Interop.Excel.IRTDUpdateEvent, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRTDUpdateEvent Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRtdServer which adds IDispose to the interface
	/// </summary>
	public interface IIRtdServer : Microsoft.Office.Interop.Excel.IRtdServer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRtdServer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame2 which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame2 : Microsoft.Office.Interop.Excel.TextFrame2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TextFrame2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFont which adds IDispose to the interface
	/// </summary>
	public interface IIFont : Microsoft.Office.Interop.Excel.IFont, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWindow which adds IDispose to the interface
	/// </summary>
	public interface IIWindow : Microsoft.Office.Interop.Excel.IWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWindow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWindows which adds IDispose to the interface
	/// </summary>
	public interface IIWindows : Microsoft.Office.Interop.Excel.IWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAppEvents which adds IDispose to the interface
	/// </summary>
	public interface IIAppEvents : Microsoft.Office.Interop.Excel.IAppEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAppEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Application which adds IDispose to the interface
	/// </summary>
	public interface I_Application : Microsoft.Office.Interop.Excel._Application, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWorksheetFunction which adds IDispose to the interface
	/// </summary>
	public interface IIWorksheetFunction : Microsoft.Office.Interop.Excel.IWorksheetFunction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWorksheetFunction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRange which adds IDispose to the interface
	/// </summary>
	public interface IIRange : Microsoft.Office.Interop.Excel.IRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartEvents which adds IDispose to the interface
	/// </summary>
	public interface IIChartEvents : Microsoft.Office.Interop.Excel.IChartEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Chart which adds IDispose to the interface
	/// </summary>
	public interface I_Chart : Microsoft.Office.Interop.Excel._Chart, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._Chart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sheets which adds IDispose to the interface
	/// </summary>
	public interface ISheets : Microsoft.Office.Interop.Excel.Sheets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Sheets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IVPageBreak which adds IDispose to the interface
	/// </summary>
	public interface IIVPageBreak : Microsoft.Office.Interop.Excel.IVPageBreak, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IVPageBreak Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IHPageBreak which adds IDispose to the interface
	/// </summary>
	public interface IIHPageBreak : Microsoft.Office.Interop.Excel.IHPageBreak, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IHPageBreak Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IHPageBreaks which adds IDispose to the interface
	/// </summary>
	public interface IIHPageBreaks : Microsoft.Office.Interop.Excel.IHPageBreaks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IHPageBreaks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IVPageBreaks which adds IDispose to the interface
	/// </summary>
	public interface IIVPageBreaks : Microsoft.Office.Interop.Excel.IVPageBreaks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IVPageBreaks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRecentFile which adds IDispose to the interface
	/// </summary>
	public interface IIRecentFile : Microsoft.Office.Interop.Excel.IRecentFile, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRecentFile Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRecentFiles which adds IDispose to the interface
	/// </summary>
	public interface IIRecentFiles : Microsoft.Office.Interop.Excel.IRecentFiles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRecentFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDocEvents which adds IDispose to the interface
	/// </summary>
	public interface IIDocEvents : Microsoft.Office.Interop.Excel.IDocEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDocEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Worksheet which adds IDispose to the interface
	/// </summary>
	public interface I_Worksheet : Microsoft.Office.Interop.Excel._Worksheet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._Worksheet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IStyle which adds IDispose to the interface
	/// </summary>
	public interface IIStyle : Microsoft.Office.Interop.Excel.IStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IStyles which adds IDispose to the interface
	/// </summary>
	public interface IIStyles : Microsoft.Office.Interop.Excel.IStyles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IStyles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IBorders which adds IDispose to the interface
	/// </summary>
	public interface IIBorders : Microsoft.Office.Interop.Excel.IBorders, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IBorders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Global which adds IDispose to the interface
	/// </summary>
	public interface I_Global : Microsoft.Office.Interop.Excel._Global, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._Global Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAddIn which adds IDispose to the interface
	/// </summary>
	public interface IIAddIn : Microsoft.Office.Interop.Excel.IAddIn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAddIn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAddIns which adds IDispose to the interface
	/// </summary>
	public interface IIAddIns : Microsoft.Office.Interop.Excel.IAddIns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAddIns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IToolbar which adds IDispose to the interface
	/// </summary>
	public interface IIToolbar : Microsoft.Office.Interop.Excel.IToolbar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IToolbar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IToolbars which adds IDispose to the interface
	/// </summary>
	public interface IIToolbars : Microsoft.Office.Interop.Excel.IToolbars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IToolbars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IToolbarButton which adds IDispose to the interface
	/// </summary>
	public interface IIToolbarButton : Microsoft.Office.Interop.Excel.IToolbarButton, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IToolbarButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IToolbarButtons which adds IDispose to the interface
	/// </summary>
	public interface IIToolbarButtons : Microsoft.Office.Interop.Excel.IToolbarButtons, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IToolbarButtons Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAreas which adds IDispose to the interface
	/// </summary>
	public interface IIAreas : Microsoft.Office.Interop.Excel.IAreas, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAreas Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWorkbookEvents which adds IDispose to the interface
	/// </summary>
	public interface IIWorkbookEvents : Microsoft.Office.Interop.Excel.IWorkbookEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWorkbookEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Workbook which adds IDispose to the interface
	/// </summary>
	public interface I_Workbook : Microsoft.Office.Interop.Excel._Workbook, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._Workbook Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Workbooks which adds IDispose to the interface
	/// </summary>
	public interface IWorkbooks : Microsoft.Office.Interop.Excel.Workbooks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Workbooks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMenuBars which adds IDispose to the interface
	/// </summary>
	public interface IIMenuBars : Microsoft.Office.Interop.Excel.IMenuBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMenuBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMenuBar which adds IDispose to the interface
	/// </summary>
	public interface IIMenuBar : Microsoft.Office.Interop.Excel.IMenuBar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMenuBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMenus which adds IDispose to the interface
	/// </summary>
	public interface IIMenus : Microsoft.Office.Interop.Excel.IMenus, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMenus Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMenu which adds IDispose to the interface
	/// </summary>
	public interface IIMenu : Microsoft.Office.Interop.Excel.IMenu, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMenu Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMenuItems which adds IDispose to the interface
	/// </summary>
	public interface IIMenuItems : Microsoft.Office.Interop.Excel.IMenuItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMenuItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMenuItem which adds IDispose to the interface
	/// </summary>
	public interface IIMenuItem : Microsoft.Office.Interop.Excel.IMenuItem, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMenuItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICharts which adds IDispose to the interface
	/// </summary>
	public interface IICharts : Microsoft.Office.Interop.Excel.ICharts, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICharts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDrawingObjects which adds IDispose to the interface
	/// </summary>
	public interface IIDrawingObjects : Microsoft.Office.Interop.Excel.IDrawingObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDrawingObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotCache which adds IDispose to the interface
	/// </summary>
	public interface IIPivotCache : Microsoft.Office.Interop.Excel.IPivotCache, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotCache Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotCaches which adds IDispose to the interface
	/// </summary>
	public interface IIPivotCaches : Microsoft.Office.Interop.Excel.IPivotCaches, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotCaches Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotFormula which adds IDispose to the interface
	/// </summary>
	public interface IIPivotFormula : Microsoft.Office.Interop.Excel.IPivotFormula, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotFormula Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotFormulas which adds IDispose to the interface
	/// </summary>
	public interface IIPivotFormulas : Microsoft.Office.Interop.Excel.IPivotFormulas, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotFormulas Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotTable which adds IDispose to the interface
	/// </summary>
	public interface IIPivotTable : Microsoft.Office.Interop.Excel.IPivotTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotTables which adds IDispose to the interface
	/// </summary>
	public interface IIPivotTables : Microsoft.Office.Interop.Excel.IPivotTables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotTables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotField which adds IDispose to the interface
	/// </summary>
	public interface IIPivotField : Microsoft.Office.Interop.Excel.IPivotField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotFields which adds IDispose to the interface
	/// </summary>
	public interface IIPivotFields : Microsoft.Office.Interop.Excel.IPivotFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICalculatedFields which adds IDispose to the interface
	/// </summary>
	public interface IICalculatedFields : Microsoft.Office.Interop.Excel.ICalculatedFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICalculatedFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotItem which adds IDispose to the interface
	/// </summary>
	public interface IIPivotItem : Microsoft.Office.Interop.Excel.IPivotItem, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotItems which adds IDispose to the interface
	/// </summary>
	public interface IIPivotItems : Microsoft.Office.Interop.Excel.IPivotItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICalculatedItems which adds IDispose to the interface
	/// </summary>
	public interface IICalculatedItems : Microsoft.Office.Interop.Excel.ICalculatedItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICalculatedItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICharacters which adds IDispose to the interface
	/// </summary>
	public interface IICharacters : Microsoft.Office.Interop.Excel.ICharacters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICharacters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDialogs which adds IDispose to the interface
	/// </summary>
	public interface IIDialogs : Microsoft.Office.Interop.Excel.IDialogs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDialogs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDialog which adds IDispose to the interface
	/// </summary>
	public interface IIDialog : Microsoft.Office.Interop.Excel.IDialog, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISoundNote which adds IDispose to the interface
	/// </summary>
	public interface IISoundNote : Microsoft.Office.Interop.Excel.ISoundNote, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISoundNote Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IButton which adds IDispose to the interface
	/// </summary>
	public interface IIButton : Microsoft.Office.Interop.Excel.IButton, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IButtons which adds IDispose to the interface
	/// </summary>
	public interface IIButtons : Microsoft.Office.Interop.Excel.IButtons, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IButtons Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICheckBox which adds IDispose to the interface
	/// </summary>
	public interface IICheckBox : Microsoft.Office.Interop.Excel.ICheckBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICheckBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICheckBoxes which adds IDispose to the interface
	/// </summary>
	public interface IICheckBoxes : Microsoft.Office.Interop.Excel.ICheckBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICheckBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOptionButton which adds IDispose to the interface
	/// </summary>
	public interface IIOptionButton : Microsoft.Office.Interop.Excel.IOptionButton, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOptionButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOptionButtons which adds IDispose to the interface
	/// </summary>
	public interface IIOptionButtons : Microsoft.Office.Interop.Excel.IOptionButtons, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOptionButtons Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IEditBox which adds IDispose to the interface
	/// </summary>
	public interface IIEditBox : Microsoft.Office.Interop.Excel.IEditBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IEditBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IEditBoxes which adds IDispose to the interface
	/// </summary>
	public interface IIEditBoxes : Microsoft.Office.Interop.Excel.IEditBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IEditBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IScrollBar which adds IDispose to the interface
	/// </summary>
	public interface IIScrollBar : Microsoft.Office.Interop.Excel.IScrollBar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IScrollBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IScrollBars which adds IDispose to the interface
	/// </summary>
	public interface IIScrollBars : Microsoft.Office.Interop.Excel.IScrollBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IScrollBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListBox which adds IDispose to the interface
	/// </summary>
	public interface IIListBox : Microsoft.Office.Interop.Excel.IListBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListBoxes which adds IDispose to the interface
	/// </summary>
	public interface IIListBoxes : Microsoft.Office.Interop.Excel.IListBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGroupBox which adds IDispose to the interface
	/// </summary>
	public interface IIGroupBox : Microsoft.Office.Interop.Excel.IGroupBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGroupBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGroupBoxes which adds IDispose to the interface
	/// </summary>
	public interface IIGroupBoxes : Microsoft.Office.Interop.Excel.IGroupBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGroupBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDropDown which adds IDispose to the interface
	/// </summary>
	public interface IIDropDown : Microsoft.Office.Interop.Excel.IDropDown, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDropDown Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDropDowns which adds IDispose to the interface
	/// </summary>
	public interface IIDropDowns : Microsoft.Office.Interop.Excel.IDropDowns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDropDowns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISpinner which adds IDispose to the interface
	/// </summary>
	public interface IISpinner : Microsoft.Office.Interop.Excel.ISpinner, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISpinner Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISpinners which adds IDispose to the interface
	/// </summary>
	public interface IISpinners : Microsoft.Office.Interop.Excel.ISpinners, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISpinners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDialogFrame which adds IDispose to the interface
	/// </summary>
	public interface IIDialogFrame : Microsoft.Office.Interop.Excel.IDialogFrame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDialogFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILabel which adds IDispose to the interface
	/// </summary>
	public interface IILabel : Microsoft.Office.Interop.Excel.ILabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILabels which adds IDispose to the interface
	/// </summary>
	public interface IILabels : Microsoft.Office.Interop.Excel.ILabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPanes which adds IDispose to the interface
	/// </summary>
	public interface IIPanes : Microsoft.Office.Interop.Excel.IPanes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPanes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPane which adds IDispose to the interface
	/// </summary>
	public interface IIPane : Microsoft.Office.Interop.Excel.IPane, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IScenarios which adds IDispose to the interface
	/// </summary>
	public interface IIScenarios : Microsoft.Office.Interop.Excel.IScenarios, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IScenarios Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IScenario which adds IDispose to the interface
	/// </summary>
	public interface IIScenario : Microsoft.Office.Interop.Excel.IScenario, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IScenario Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGroupObject which adds IDispose to the interface
	/// </summary>
	public interface IIGroupObject : Microsoft.Office.Interop.Excel.IGroupObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGroupObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGroupObjects which adds IDispose to the interface
	/// </summary>
	public interface IIGroupObjects : Microsoft.Office.Interop.Excel.IGroupObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGroupObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILine which adds IDispose to the interface
	/// </summary>
	public interface IILine : Microsoft.Office.Interop.Excel.ILine, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILine Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILines which adds IDispose to the interface
	/// </summary>
	public interface IILines : Microsoft.Office.Interop.Excel.ILines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRectangle which adds IDispose to the interface
	/// </summary>
	public interface IIRectangle : Microsoft.Office.Interop.Excel.IRectangle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRectangle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRectangles which adds IDispose to the interface
	/// </summary>
	public interface IIRectangles : Microsoft.Office.Interop.Excel.IRectangles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRectangles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOval which adds IDispose to the interface
	/// </summary>
	public interface IIOval : Microsoft.Office.Interop.Excel.IOval, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOval Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOvals which adds IDispose to the interface
	/// </summary>
	public interface IIOvals : Microsoft.Office.Interop.Excel.IOvals, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOvals Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IArc which adds IDispose to the interface
	/// </summary>
	public interface IIArc : Microsoft.Office.Interop.Excel.IArc, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IArc Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IArcs which adds IDispose to the interface
	/// </summary>
	public interface IIArcs : Microsoft.Office.Interop.Excel.IArcs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IArcs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOLEObjectEvents which adds IDispose to the interface
	/// </summary>
	public interface IIOLEObjectEvents : Microsoft.Office.Interop.Excel.IOLEObjectEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOLEObjectEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _IOLEObject which adds IDispose to the interface
	/// </summary>
	public interface I_IOLEObject : Microsoft.Office.Interop.Excel._IOLEObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._IOLEObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOLEObjects which adds IDispose to the interface
	/// </summary>
	public interface IIOLEObjects : Microsoft.Office.Interop.Excel.IOLEObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOLEObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITextBox which adds IDispose to the interface
	/// </summary>
	public interface IITextBox : Microsoft.Office.Interop.Excel.ITextBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITextBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITextBoxes which adds IDispose to the interface
	/// </summary>
	public interface IITextBoxes : Microsoft.Office.Interop.Excel.ITextBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITextBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPicture which adds IDispose to the interface
	/// </summary>
	public interface IIPicture : Microsoft.Office.Interop.Excel.IPicture, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPicture Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPictures which adds IDispose to the interface
	/// </summary>
	public interface IIPictures : Microsoft.Office.Interop.Excel.IPictures, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPictures Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDrawing which adds IDispose to the interface
	/// </summary>
	public interface IIDrawing : Microsoft.Office.Interop.Excel.IDrawing, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDrawing Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDrawings which adds IDispose to the interface
	/// </summary>
	public interface IIDrawings : Microsoft.Office.Interop.Excel.IDrawings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDrawings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRoutingSlip which adds IDispose to the interface
	/// </summary>
	public interface IIRoutingSlip : Microsoft.Office.Interop.Excel.IRoutingSlip, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRoutingSlip Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOutline which adds IDispose to the interface
	/// </summary>
	public interface IIOutline : Microsoft.Office.Interop.Excel.IOutline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOutline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IModule which adds IDispose to the interface
	/// </summary>
	public interface IIModule : Microsoft.Office.Interop.Excel.IModule, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IModule Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IModules which adds IDispose to the interface
	/// </summary>
	public interface IIModules : Microsoft.Office.Interop.Excel.IModules, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IModules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDialogSheet which adds IDispose to the interface
	/// </summary>
	public interface IIDialogSheet : Microsoft.Office.Interop.Excel.IDialogSheet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDialogSheet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDialogSheets which adds IDispose to the interface
	/// </summary>
	public interface IIDialogSheets : Microsoft.Office.Interop.Excel.IDialogSheets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDialogSheets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWorksheets which adds IDispose to the interface
	/// </summary>
	public interface IIWorksheets : Microsoft.Office.Interop.Excel.IWorksheets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWorksheets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPageSetup which adds IDispose to the interface
	/// </summary>
	public interface IIPageSetup : Microsoft.Office.Interop.Excel.IPageSetup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPageSetup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for INames which adds IDispose to the interface
	/// </summary>
	public interface IINames : Microsoft.Office.Interop.Excel.INames, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.INames Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IName which adds IDispose to the interface
	/// </summary>
	public interface IIName : Microsoft.Office.Interop.Excel.IName, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IName Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartObject which adds IDispose to the interface
	/// </summary>
	public interface IIChartObject : Microsoft.Office.Interop.Excel.IChartObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartObjects which adds IDispose to the interface
	/// </summary>
	public interface IIChartObjects : Microsoft.Office.Interop.Excel.IChartObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMailer which adds IDispose to the interface
	/// </summary>
	public interface IIMailer : Microsoft.Office.Interop.Excel.IMailer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMailer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomViews which adds IDispose to the interface
	/// </summary>
	public interface IICustomViews : Microsoft.Office.Interop.Excel.ICustomViews, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICustomViews Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomView which adds IDispose to the interface
	/// </summary>
	public interface IICustomView : Microsoft.Office.Interop.Excel.ICustomView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICustomView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFormatConditions which adds IDispose to the interface
	/// </summary>
	public interface IIFormatConditions : Microsoft.Office.Interop.Excel.IFormatConditions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFormatConditions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFormatCondition which adds IDispose to the interface
	/// </summary>
	public interface IIFormatCondition : Microsoft.Office.Interop.Excel.IFormatCondition, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFormatCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IComments which adds IDispose to the interface
	/// </summary>
	public interface IIComments : Microsoft.Office.Interop.Excel.IComments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IComments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IComment which adds IDispose to the interface
	/// </summary>
	public interface IIComment : Microsoft.Office.Interop.Excel.IComment, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IComment Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRefreshEvents which adds IDispose to the interface
	/// </summary>
	public interface IIRefreshEvents : Microsoft.Office.Interop.Excel.IRefreshEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRefreshEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _IQueryTable which adds IDispose to the interface
	/// </summary>
	public interface I_IQueryTable : Microsoft.Office.Interop.Excel._IQueryTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._IQueryTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IQueryTables which adds IDispose to the interface
	/// </summary>
	public interface IIQueryTables : Microsoft.Office.Interop.Excel.IQueryTables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IQueryTables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IParameter which adds IDispose to the interface
	/// </summary>
	public interface IIParameter : Microsoft.Office.Interop.Excel.IParameter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IParameter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IParameters which adds IDispose to the interface
	/// </summary>
	public interface IIParameters : Microsoft.Office.Interop.Excel.IParameters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IParameters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IODBCError which adds IDispose to the interface
	/// </summary>
	public interface IIODBCError : Microsoft.Office.Interop.Excel.IODBCError, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IODBCError Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IODBCErrors which adds IDispose to the interface
	/// </summary>
	public interface IIODBCErrors : Microsoft.Office.Interop.Excel.IODBCErrors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IODBCErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IValidation which adds IDispose to the interface
	/// </summary>
	public interface IIValidation : Microsoft.Office.Interop.Excel.IValidation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IValidation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IHyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IIHyperlinks : Microsoft.Office.Interop.Excel.IHyperlinks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IHyperlinks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IHyperlink which adds IDispose to the interface
	/// </summary>
	public interface IIHyperlink : Microsoft.Office.Interop.Excel.IHyperlink, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IHyperlink Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAutoFilter which adds IDispose to the interface
	/// </summary>
	public interface IIAutoFilter : Microsoft.Office.Interop.Excel.IAutoFilter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAutoFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFilters which adds IDispose to the interface
	/// </summary>
	public interface IIFilters : Microsoft.Office.Interop.Excel.IFilters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFilters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFilter which adds IDispose to the interface
	/// </summary>
	public interface IIFilter : Microsoft.Office.Interop.Excel.IFilter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAutoCorrect which adds IDispose to the interface
	/// </summary>
	public interface IIAutoCorrect : Microsoft.Office.Interop.Excel.IAutoCorrect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAutoCorrect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IBorder which adds IDispose to the interface
	/// </summary>
	public interface IIBorder : Microsoft.Office.Interop.Excel.IBorder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IBorder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IInterior which adds IDispose to the interface
	/// </summary>
	public interface IIInterior : Microsoft.Office.Interop.Excel.IInterior, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IInterior Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IIChartFillFormat : Microsoft.Office.Interop.Excel.IChartFillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartFillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IIChartColorFormat : Microsoft.Office.Interop.Excel.IChartColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAxis which adds IDispose to the interface
	/// </summary>
	public interface IIAxis : Microsoft.Office.Interop.Excel.IAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IIChartTitle : Microsoft.Office.Interop.Excel.IChartTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IIAxisTitle : Microsoft.Office.Interop.Excel.IAxisTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAxisTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IIChartGroup : Microsoft.Office.Interop.Excel.IChartGroup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IIChartGroups : Microsoft.Office.Interop.Excel.IChartGroups, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAxes which adds IDispose to the interface
	/// </summary>
	public interface IIAxes : Microsoft.Office.Interop.Excel.IAxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPoints which adds IDispose to the interface
	/// </summary>
	public interface IIPoints : Microsoft.Office.Interop.Excel.IPoints, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPoints Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPoint which adds IDispose to the interface
	/// </summary>
	public interface IIPoint : Microsoft.Office.Interop.Excel.IPoint, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPoint Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISeries which adds IDispose to the interface
	/// </summary>
	public interface IISeries : Microsoft.Office.Interop.Excel.ISeries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISeries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface IISeriesCollection : Microsoft.Office.Interop.Excel.ISeriesCollection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISeriesCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDataLabel which adds IDispose to the interface
	/// </summary>
	public interface IIDataLabel : Microsoft.Office.Interop.Excel.IDataLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDataLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDataLabels which adds IDispose to the interface
	/// </summary>
	public interface IIDataLabels : Microsoft.Office.Interop.Excel.IDataLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDataLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILegendEntry which adds IDispose to the interface
	/// </summary>
	public interface IILegendEntry : Microsoft.Office.Interop.Excel.ILegendEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILegendEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILegendEntries which adds IDispose to the interface
	/// </summary>
	public interface IILegendEntries : Microsoft.Office.Interop.Excel.ILegendEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILegendEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILegendKey which adds IDispose to the interface
	/// </summary>
	public interface IILegendKey : Microsoft.Office.Interop.Excel.ILegendKey, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILegendKey Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITrendlines which adds IDispose to the interface
	/// </summary>
	public interface IITrendlines : Microsoft.Office.Interop.Excel.ITrendlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITrendlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITrendline which adds IDispose to the interface
	/// </summary>
	public interface IITrendline : Microsoft.Office.Interop.Excel.ITrendline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITrendline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICorners which adds IDispose to the interface
	/// </summary>
	public interface IICorners : Microsoft.Office.Interop.Excel.ICorners, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICorners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISeriesLines which adds IDispose to the interface
	/// </summary>
	public interface IISeriesLines : Microsoft.Office.Interop.Excel.ISeriesLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISeriesLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IHiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IIHiLoLines : Microsoft.Office.Interop.Excel.IHiLoLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IHiLoLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGridlines which adds IDispose to the interface
	/// </summary>
	public interface IIGridlines : Microsoft.Office.Interop.Excel.IGridlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGridlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDropLines which adds IDispose to the interface
	/// </summary>
	public interface IIDropLines : Microsoft.Office.Interop.Excel.IDropLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDropLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILeaderLines which adds IDispose to the interface
	/// </summary>
	public interface IILeaderLines : Microsoft.Office.Interop.Excel.ILeaderLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILeaderLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IUpBars which adds IDispose to the interface
	/// </summary>
	public interface IIUpBars : Microsoft.Office.Interop.Excel.IUpBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IUpBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDownBars which adds IDispose to the interface
	/// </summary>
	public interface IIDownBars : Microsoft.Office.Interop.Excel.IDownBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDownBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFloor which adds IDispose to the interface
	/// </summary>
	public interface IIFloor : Microsoft.Office.Interop.Excel.IFloor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFloor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWalls which adds IDispose to the interface
	/// </summary>
	public interface IIWalls : Microsoft.Office.Interop.Excel.IWalls, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWalls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITickLabels which adds IDispose to the interface
	/// </summary>
	public interface IITickLabels : Microsoft.Office.Interop.Excel.ITickLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITickLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPlotArea which adds IDispose to the interface
	/// </summary>
	public interface IIPlotArea : Microsoft.Office.Interop.Excel.IPlotArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPlotArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartArea which adds IDispose to the interface
	/// </summary>
	public interface IIChartArea : Microsoft.Office.Interop.Excel.IChartArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILegend which adds IDispose to the interface
	/// </summary>
	public interface IILegend : Microsoft.Office.Interop.Excel.ILegend, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILegend Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IIErrorBars : Microsoft.Office.Interop.Excel.IErrorBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IErrorBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDataTable which adds IDispose to the interface
	/// </summary>
	public interface IIDataTable : Microsoft.Office.Interop.Excel.IDataTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDataTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPhonetic which adds IDispose to the interface
	/// </summary>
	public interface IIPhonetic : Microsoft.Office.Interop.Excel.IPhonetic, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPhonetic Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IShape which adds IDispose to the interface
	/// </summary>
	public interface IIShape : Microsoft.Office.Interop.Excel.IShape, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IShape Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IShapes which adds IDispose to the interface
	/// </summary>
	public interface IIShapes : Microsoft.Office.Interop.Excel.IShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IIShapeRange : Microsoft.Office.Interop.Excel.IShapeRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IShapeRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IIGroupShapes : Microsoft.Office.Interop.Excel.IGroupShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGroupShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITextFrame which adds IDispose to the interface
	/// </summary>
	public interface IITextFrame : Microsoft.Office.Interop.Excel.ITextFrame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITextFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IIConnectorFormat : Microsoft.Office.Interop.Excel.IConnectorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IConnectorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IIFreeformBuilder : Microsoft.Office.Interop.Excel.IFreeformBuilder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFreeformBuilder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IControlFormat which adds IDispose to the interface
	/// </summary>
	public interface IIControlFormat : Microsoft.Office.Interop.Excel.IControlFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IControlFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOLEFormat which adds IDispose to the interface
	/// </summary>
	public interface IIOLEFormat : Microsoft.Office.Interop.Excel.IOLEFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOLEFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILinkFormat which adds IDispose to the interface
	/// </summary>
	public interface IILinkFormat : Microsoft.Office.Interop.Excel.ILinkFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILinkFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPublishObjects which adds IDispose to the interface
	/// </summary>
	public interface IIPublishObjects : Microsoft.Office.Interop.Excel.IPublishObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPublishObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PublishObject which adds IDispose to the interface
	/// </summary>
	public interface IPublishObject : Microsoft.Office.Interop.Excel.PublishObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PublishObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOLEDBError which adds IDispose to the interface
	/// </summary>
	public interface IIOLEDBError : Microsoft.Office.Interop.Excel.IOLEDBError, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOLEDBError Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOLEDBErrors which adds IDispose to the interface
	/// </summary>
	public interface IIOLEDBErrors : Microsoft.Office.Interop.Excel.IOLEDBErrors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOLEDBErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPhonetics which adds IDispose to the interface
	/// </summary>
	public interface IIPhonetics : Microsoft.Office.Interop.Excel.IPhonetics, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPhonetics Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
	/// </summary>
	public interface IDefaultWebOptions : Microsoft.Office.Interop.Excel.DefaultWebOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DefaultWebOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebOptions which adds IDispose to the interface
	/// </summary>
	public interface IWebOptions : Microsoft.Office.Interop.Excel.WebOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.WebOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotLayout which adds IDispose to the interface
	/// </summary>
	public interface IIPivotLayout : Microsoft.Office.Interop.Excel.IPivotLayout, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotLayout Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TreeviewControl which adds IDispose to the interface
	/// </summary>
	public interface ITreeviewControl : Microsoft.Office.Interop.Excel.TreeviewControl, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TreeviewControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CubeField which adds IDispose to the interface
	/// </summary>
	public interface ICubeField : Microsoft.Office.Interop.Excel.CubeField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CubeField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CubeFields which adds IDispose to the interface
	/// </summary>
	public interface ICubeFields : Microsoft.Office.Interop.Excel.CubeFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CubeFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IIDisplayUnitLabel : Microsoft.Office.Interop.Excel.IDisplayUnitLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDisplayUnitLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICellFormat which adds IDispose to the interface
	/// </summary>
	public interface IICellFormat : Microsoft.Office.Interop.Excel.ICellFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICellFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IUsedObjects which adds IDispose to the interface
	/// </summary>
	public interface IIUsedObjects : Microsoft.Office.Interop.Excel.IUsedObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IUsedObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomProperties which adds IDispose to the interface
	/// </summary>
	public interface IICustomProperties : Microsoft.Office.Interop.Excel.ICustomProperties, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICustomProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICustomProperty which adds IDispose to the interface
	/// </summary>
	public interface IICustomProperty : Microsoft.Office.Interop.Excel.ICustomProperty, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICustomProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICalculatedMembers which adds IDispose to the interface
	/// </summary>
	public interface IICalculatedMembers : Microsoft.Office.Interop.Excel.ICalculatedMembers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICalculatedMembers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICalculatedMember which adds IDispose to the interface
	/// </summary>
	public interface IICalculatedMember : Microsoft.Office.Interop.Excel.ICalculatedMember, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICalculatedMember Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWatches which adds IDispose to the interface
	/// </summary>
	public interface IIWatches : Microsoft.Office.Interop.Excel.IWatches, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWatches Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWatch which adds IDispose to the interface
	/// </summary>
	public interface IIWatch : Microsoft.Office.Interop.Excel.IWatch, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWatch Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotCell which adds IDispose to the interface
	/// </summary>
	public interface IIPivotCell : Microsoft.Office.Interop.Excel.IPivotCell, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotCell Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IGraphic which adds IDispose to the interface
	/// </summary>
	public interface IIGraphic : Microsoft.Office.Interop.Excel.IGraphic, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IGraphic Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAutoRecover which adds IDispose to the interface
	/// </summary>
	public interface IIAutoRecover : Microsoft.Office.Interop.Excel.IAutoRecover, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAutoRecover Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IErrorCheckingOptions which adds IDispose to the interface
	/// </summary>
	public interface IIErrorCheckingOptions : Microsoft.Office.Interop.Excel.IErrorCheckingOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IErrorCheckingOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IErrors which adds IDispose to the interface
	/// </summary>
	public interface IIErrors : Microsoft.Office.Interop.Excel.IErrors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IError which adds IDispose to the interface
	/// </summary>
	public interface IIError : Microsoft.Office.Interop.Excel.IError, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IError Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTagAction which adds IDispose to the interface
	/// </summary>
	public interface IISmartTagAction : Microsoft.Office.Interop.Excel.ISmartTagAction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTagAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTagActions which adds IDispose to the interface
	/// </summary>
	public interface IISmartTagActions : Microsoft.Office.Interop.Excel.ISmartTagActions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTagActions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTag which adds IDispose to the interface
	/// </summary>
	public interface IISmartTag : Microsoft.Office.Interop.Excel.ISmartTag, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTag Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTags which adds IDispose to the interface
	/// </summary>
	public interface IISmartTags : Microsoft.Office.Interop.Excel.ISmartTags, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTags Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTagRecognizer which adds IDispose to the interface
	/// </summary>
	public interface IISmartTagRecognizer : Microsoft.Office.Interop.Excel.ISmartTagRecognizer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTagRecognizer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTagRecognizers which adds IDispose to the interface
	/// </summary>
	public interface IISmartTagRecognizers : Microsoft.Office.Interop.Excel.ISmartTagRecognizers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTagRecognizers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISmartTagOptions which adds IDispose to the interface
	/// </summary>
	public interface IISmartTagOptions : Microsoft.Office.Interop.Excel.ISmartTagOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISmartTagOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISpellingOptions which adds IDispose to the interface
	/// </summary>
	public interface IISpellingOptions : Microsoft.Office.Interop.Excel.ISpellingOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISpellingOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISpeech which adds IDispose to the interface
	/// </summary>
	public interface IISpeech : Microsoft.Office.Interop.Excel.ISpeech, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISpeech Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IProtection which adds IDispose to the interface
	/// </summary>
	public interface IIProtection : Microsoft.Office.Interop.Excel.IProtection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IProtection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotItemList which adds IDispose to the interface
	/// </summary>
	public interface IIPivotItemList : Microsoft.Office.Interop.Excel.IPivotItemList, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotItemList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITab which adds IDispose to the interface
	/// </summary>
	public interface IITab : Microsoft.Office.Interop.Excel.ITab, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITab Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAllowEditRanges which adds IDispose to the interface
	/// </summary>
	public interface IIAllowEditRanges : Microsoft.Office.Interop.Excel.IAllowEditRanges, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAllowEditRanges Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAllowEditRange which adds IDispose to the interface
	/// </summary>
	public interface IIAllowEditRange : Microsoft.Office.Interop.Excel.IAllowEditRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAllowEditRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IUserAccessList which adds IDispose to the interface
	/// </summary>
	public interface IIUserAccessList : Microsoft.Office.Interop.Excel.IUserAccessList, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IUserAccessList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IUserAccess which adds IDispose to the interface
	/// </summary>
	public interface IIUserAccess : Microsoft.Office.Interop.Excel.IUserAccess, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IUserAccess Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRTD which adds IDispose to the interface
	/// </summary>
	public interface IIRTD : Microsoft.Office.Interop.Excel.IRTD, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRTD Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDiagram which adds IDispose to the interface
	/// </summary>
	public interface IIDiagram : Microsoft.Office.Interop.Excel.IDiagram, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDiagram Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListObjects which adds IDispose to the interface
	/// </summary>
	public interface IIListObjects : Microsoft.Office.Interop.Excel.IListObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListObject which adds IDispose to the interface
	/// </summary>
	public interface IIListObject : Microsoft.Office.Interop.Excel.IListObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListColumns which adds IDispose to the interface
	/// </summary>
	public interface IIListColumns : Microsoft.Office.Interop.Excel.IListColumns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListColumns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListColumn which adds IDispose to the interface
	/// </summary>
	public interface IIListColumn : Microsoft.Office.Interop.Excel.IListColumn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListColumn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListRows which adds IDispose to the interface
	/// </summary>
	public interface IIListRows : Microsoft.Office.Interop.Excel.IListRows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListRows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListRow which adds IDispose to the interface
	/// </summary>
	public interface IIListRow : Microsoft.Office.Interop.Excel.IListRow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListRow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlNamespace which adds IDispose to the interface
	/// </summary>
	public interface IIXmlNamespace : Microsoft.Office.Interop.Excel.IXmlNamespace, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlNamespace Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlNamespaces which adds IDispose to the interface
	/// </summary>
	public interface IIXmlNamespaces : Microsoft.Office.Interop.Excel.IXmlNamespaces, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlNamespaces Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlDataBinding which adds IDispose to the interface
	/// </summary>
	public interface IIXmlDataBinding : Microsoft.Office.Interop.Excel.IXmlDataBinding, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlDataBinding Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlSchema which adds IDispose to the interface
	/// </summary>
	public interface IIXmlSchema : Microsoft.Office.Interop.Excel.IXmlSchema, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlSchema Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlSchemas which adds IDispose to the interface
	/// </summary>
	public interface IIXmlSchemas : Microsoft.Office.Interop.Excel.IXmlSchemas, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlSchemas Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlMap which adds IDispose to the interface
	/// </summary>
	public interface IIXmlMap : Microsoft.Office.Interop.Excel.IXmlMap, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlMap Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXmlMaps which adds IDispose to the interface
	/// </summary>
	public interface IIXmlMaps : Microsoft.Office.Interop.Excel.IXmlMaps, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXmlMaps Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IListDataFormat which adds IDispose to the interface
	/// </summary>
	public interface IIListDataFormat : Microsoft.Office.Interop.Excel.IListDataFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IListDataFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IXPath which adds IDispose to the interface
	/// </summary>
	public interface IIXPath : Microsoft.Office.Interop.Excel.IXPath, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IXPath Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotLineCells which adds IDispose to the interface
	/// </summary>
	public interface IIPivotLineCells : Microsoft.Office.Interop.Excel.IPivotLineCells, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotLineCells Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotLine which adds IDispose to the interface
	/// </summary>
	public interface IIPivotLine : Microsoft.Office.Interop.Excel.IPivotLine, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotLine Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotLines which adds IDispose to the interface
	/// </summary>
	public interface IIPivotLines : Microsoft.Office.Interop.Excel.IPivotLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotAxis which adds IDispose to the interface
	/// </summary>
	public interface IIPivotAxis : Microsoft.Office.Interop.Excel.IPivotAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotFilter which adds IDispose to the interface
	/// </summary>
	public interface IIPivotFilter : Microsoft.Office.Interop.Excel.IPivotFilter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotFilters which adds IDispose to the interface
	/// </summary>
	public interface IIPivotFilters : Microsoft.Office.Interop.Excel.IPivotFilters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotFilters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWorkbookConnection which adds IDispose to the interface
	/// </summary>
	public interface IIWorkbookConnection : Microsoft.Office.Interop.Excel.IWorkbookConnection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWorkbookConnection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConnections which adds IDispose to the interface
	/// </summary>
	public interface IIConnections : Microsoft.Office.Interop.Excel.IConnections, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IConnections Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IWorksheetView which adds IDispose to the interface
	/// </summary>
	public interface IIWorksheetView : Microsoft.Office.Interop.Excel.IWorksheetView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IWorksheetView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartView which adds IDispose to the interface
	/// </summary>
	public interface IIChartView : Microsoft.Office.Interop.Excel.IChartView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IModuleView which adds IDispose to the interface
	/// </summary>
	public interface IIModuleView : Microsoft.Office.Interop.Excel.IModuleView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IModuleView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDialogSheetView which adds IDispose to the interface
	/// </summary>
	public interface IIDialogSheetView : Microsoft.Office.Interop.Excel.IDialogSheetView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDialogSheetView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISheetViews which adds IDispose to the interface
	/// </summary>
	public interface IISheetViews : Microsoft.Office.Interop.Excel.ISheetViews, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISheetViews Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IOLEDBConnection which adds IDispose to the interface
	/// </summary>
	public interface IIOLEDBConnection : Microsoft.Office.Interop.Excel.IOLEDBConnection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IOLEDBConnection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IODBCConnection which adds IDispose to the interface
	/// </summary>
	public interface IIODBCConnection : Microsoft.Office.Interop.Excel.IODBCConnection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IODBCConnection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAction which adds IDispose to the interface
	/// </summary>
	public interface IIAction : Microsoft.Office.Interop.Excel.IAction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IActions which adds IDispose to the interface
	/// </summary>
	public interface IIActions : Microsoft.Office.Interop.Excel.IActions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IActions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFormatColor which adds IDispose to the interface
	/// </summary>
	public interface IIFormatColor : Microsoft.Office.Interop.Excel.IFormatColor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFormatColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IConditionValue which adds IDispose to the interface
	/// </summary>
	public interface IIConditionValue : Microsoft.Office.Interop.Excel.IConditionValue, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IConditionValue Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IColorScale which adds IDispose to the interface
	/// </summary>
	public interface IIColorScale : Microsoft.Office.Interop.Excel.IColorScale, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IColorScale Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IColorScaleCriteria which adds IDispose to the interface
	/// </summary>
	public interface IIColorScaleCriteria : Microsoft.Office.Interop.Excel.IColorScaleCriteria, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IColorScaleCriteria Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IColorScaleCriterion which adds IDispose to the interface
	/// </summary>
	public interface IIColorScaleCriterion : Microsoft.Office.Interop.Excel.IColorScaleCriterion, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IColorScaleCriterion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDatabar which adds IDispose to the interface
	/// </summary>
	public interface IIDatabar : Microsoft.Office.Interop.Excel.IDatabar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDatabar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IIconSetCondition which adds IDispose to the interface
	/// </summary>
	public interface IIIconSetCondition : Microsoft.Office.Interop.Excel.IIconSetCondition, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IIconSetCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IIconCriteria which adds IDispose to the interface
	/// </summary>
	public interface IIIconCriteria : Microsoft.Office.Interop.Excel.IIconCriteria, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IIconCriteria Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IIconCriterion which adds IDispose to the interface
	/// </summary>
	public interface IIIconCriterion : Microsoft.Office.Interop.Excel.IIconCriterion, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IIconCriterion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IIcon which adds IDispose to the interface
	/// </summary>
	public interface IIIcon : Microsoft.Office.Interop.Excel.IIcon, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IIcon Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IIconSet which adds IDispose to the interface
	/// </summary>
	public interface IIIconSet : Microsoft.Office.Interop.Excel.IIconSet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IIconSet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IIconSets which adds IDispose to the interface
	/// </summary>
	public interface IIIconSets : Microsoft.Office.Interop.Excel.IIconSets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IIconSets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITop10 which adds IDispose to the interface
	/// </summary>
	public interface IITop10 : Microsoft.Office.Interop.Excel.ITop10, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITop10 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAboveAverage which adds IDispose to the interface
	/// </summary>
	public interface IIAboveAverage : Microsoft.Office.Interop.Excel.IAboveAverage, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAboveAverage Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IUniqueValues which adds IDispose to the interface
	/// </summary>
	public interface IIUniqueValues : Microsoft.Office.Interop.Excel.IUniqueValues, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IUniqueValues Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRanges which adds IDispose to the interface
	/// </summary>
	public interface IIRanges : Microsoft.Office.Interop.Excel.IRanges, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRanges Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IHeaderFooter which adds IDispose to the interface
	/// </summary>
	public interface IIHeaderFooter : Microsoft.Office.Interop.Excel.IHeaderFooter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IHeaderFooter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPage which adds IDispose to the interface
	/// </summary>
	public interface IIPage : Microsoft.Office.Interop.Excel.IPage, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPage Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPages which adds IDispose to the interface
	/// </summary>
	public interface IIPages : Microsoft.Office.Interop.Excel.IPages, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPages Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IServerViewableItems which adds IDispose to the interface
	/// </summary>
	public interface IIServerViewableItems : Microsoft.Office.Interop.Excel.IServerViewableItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IServerViewableItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITableStyleElement which adds IDispose to the interface
	/// </summary>
	public interface IITableStyleElement : Microsoft.Office.Interop.Excel.ITableStyleElement, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITableStyleElement Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITableStyleElements which adds IDispose to the interface
	/// </summary>
	public interface IITableStyleElements : Microsoft.Office.Interop.Excel.ITableStyleElements, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITableStyleElements Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITableStyle which adds IDispose to the interface
	/// </summary>
	public interface IITableStyle : Microsoft.Office.Interop.Excel.ITableStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITableStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ITableStyles which adds IDispose to the interface
	/// </summary>
	public interface IITableStyles : Microsoft.Office.Interop.Excel.ITableStyles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ITableStyles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISortField which adds IDispose to the interface
	/// </summary>
	public interface IISortField : Microsoft.Office.Interop.Excel.ISortField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISortField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISortFields which adds IDispose to the interface
	/// </summary>
	public interface IISortFields : Microsoft.Office.Interop.Excel.ISortFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISortFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISort which adds IDispose to the interface
	/// </summary>
	public interface IISort : Microsoft.Office.Interop.Excel.ISort, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISort Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IResearch which adds IDispose to the interface
	/// </summary>
	public interface IIResearch : Microsoft.Office.Interop.Excel.IResearch, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IResearch Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IColorStop which adds IDispose to the interface
	/// </summary>
	public interface IIColorStop : Microsoft.Office.Interop.Excel.IColorStop, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IColorStop Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IColorStops which adds IDispose to the interface
	/// </summary>
	public interface IIColorStops : Microsoft.Office.Interop.Excel.IColorStops, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IColorStops Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ILinearGradient which adds IDispose to the interface
	/// </summary>
	public interface IILinearGradient : Microsoft.Office.Interop.Excel.ILinearGradient, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ILinearGradient Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IRectangularGradient which adds IDispose to the interface
	/// </summary>
	public interface IIRectangularGradient : Microsoft.Office.Interop.Excel.IRectangularGradient, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IRectangularGradient Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IMultiThreadedCalculation which adds IDispose to the interface
	/// </summary>
	public interface IIMultiThreadedCalculation : Microsoft.Office.Interop.Excel.IMultiThreadedCalculation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IMultiThreadedCalculation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IIChartFormat : Microsoft.Office.Interop.Excel.IChartFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IChartFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFileExportConverter which adds IDispose to the interface
	/// </summary>
	public interface IIFileExportConverter : Microsoft.Office.Interop.Excel.IFileExportConverter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFileExportConverter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IFileExportConverters which adds IDispose to the interface
	/// </summary>
	public interface IIFileExportConverters : Microsoft.Office.Interop.Excel.IFileExportConverters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IFileExportConverters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IAddIns2 which adds IDispose to the interface
	/// </summary>
	public interface IIAddIns2 : Microsoft.Office.Interop.Excel.IAddIns2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IAddIns2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparklineGroups which adds IDispose to the interface
	/// </summary>
	public interface IISparklineGroups : Microsoft.Office.Interop.Excel.ISparklineGroups, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparklineGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparklineGroup which adds IDispose to the interface
	/// </summary>
	public interface IISparklineGroup : Microsoft.Office.Interop.Excel.ISparklineGroup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparklineGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparkPoints which adds IDispose to the interface
	/// </summary>
	public interface IISparkPoints : Microsoft.Office.Interop.Excel.ISparkPoints, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparkPoints Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparkline which adds IDispose to the interface
	/// </summary>
	public interface IISparkline : Microsoft.Office.Interop.Excel.ISparkline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparkline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparkAxes which adds IDispose to the interface
	/// </summary>
	public interface IISparkAxes : Microsoft.Office.Interop.Excel.ISparkAxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparkAxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparkHorizontalAxis which adds IDispose to the interface
	/// </summary>
	public interface IISparkHorizontalAxis : Microsoft.Office.Interop.Excel.ISparkHorizontalAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparkHorizontalAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparkVerticalAxis which adds IDispose to the interface
	/// </summary>
	public interface IISparkVerticalAxis : Microsoft.Office.Interop.Excel.ISparkVerticalAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparkVerticalAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISparkColor which adds IDispose to the interface
	/// </summary>
	public interface IISparkColor : Microsoft.Office.Interop.Excel.ISparkColor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISparkColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDataBarBorder which adds IDispose to the interface
	/// </summary>
	public interface IIDataBarBorder : Microsoft.Office.Interop.Excel.IDataBarBorder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDataBarBorder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for INegativeBarFormat which adds IDispose to the interface
	/// </summary>
	public interface IINegativeBarFormat : Microsoft.Office.Interop.Excel.INegativeBarFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.INegativeBarFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IValueChange which adds IDispose to the interface
	/// </summary>
	public interface IIValueChange : Microsoft.Office.Interop.Excel.IValueChange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IValueChange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IPivotTableChangeList which adds IDispose to the interface
	/// </summary>
	public interface IIPivotTableChangeList : Microsoft.Office.Interop.Excel.IPivotTableChangeList, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IPivotTableChangeList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDisplayFormat which adds IDispose to the interface
	/// </summary>
	public interface IIDisplayFormat : Microsoft.Office.Interop.Excel.IDisplayFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDisplayFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerCaches which adds IDispose to the interface
	/// </summary>
	public interface IISlicerCaches : Microsoft.Office.Interop.Excel.ISlicerCaches, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerCaches Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerCache which adds IDispose to the interface
	/// </summary>
	public interface IISlicerCache : Microsoft.Office.Interop.Excel.ISlicerCache, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerCache Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerCacheLevels which adds IDispose to the interface
	/// </summary>
	public interface IISlicerCacheLevels : Microsoft.Office.Interop.Excel.ISlicerCacheLevels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerCacheLevels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerCacheLevel which adds IDispose to the interface
	/// </summary>
	public interface IISlicerCacheLevel : Microsoft.Office.Interop.Excel.ISlicerCacheLevel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerCacheLevel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicers which adds IDispose to the interface
	/// </summary>
	public interface IISlicers : Microsoft.Office.Interop.Excel.ISlicers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicer which adds IDispose to the interface
	/// </summary>
	public interface IISlicer : Microsoft.Office.Interop.Excel.ISlicer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerItem which adds IDispose to the interface
	/// </summary>
	public interface IISlicerItem : Microsoft.Office.Interop.Excel.ISlicerItem, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerItems which adds IDispose to the interface
	/// </summary>
	public interface IISlicerItems : Microsoft.Office.Interop.Excel.ISlicerItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ISlicerPivotTables which adds IDispose to the interface
	/// </summary>
	public interface IISlicerPivotTables : Microsoft.Office.Interop.Excel.ISlicerPivotTables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ISlicerPivotTables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IProtectedViewWindows which adds IDispose to the interface
	/// </summary>
	public interface IIProtectedViewWindows : Microsoft.Office.Interop.Excel.IProtectedViewWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IProtectedViewWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IProtectedViewWindow which adds IDispose to the interface
	/// </summary>
	public interface IIProtectedViewWindow : Microsoft.Office.Interop.Excel.IProtectedViewWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IProtectedViewWindow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Font which adds IDispose to the interface
	/// </summary>
	public interface IFont : Microsoft.Office.Interop.Excel.Font, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Font Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Window which adds IDispose to the interface
	/// </summary>
	public interface IWindow : Microsoft.Office.Interop.Excel.Window, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Window Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Windows which adds IDispose to the interface
	/// </summary>
	public interface IWindows : Microsoft.Office.Interop.Excel.Windows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Windows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AppEvents which adds IDispose to the interface
	/// </summary>
	public interface IAppEvents : Microsoft.Office.Interop.Excel.AppEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AppEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorksheetFunction which adds IDispose to the interface
	/// </summary>
	public interface IWorksheetFunction : Microsoft.Office.Interop.Excel.WorksheetFunction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.WorksheetFunction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Range which adds IDispose to the interface
	/// </summary>
	public interface IRange : Microsoft.Office.Interop.Excel.Range, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Range Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartEvents which adds IDispose to the interface
	/// </summary>
	public interface IChartEvents : Microsoft.Office.Interop.Excel.ChartEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for VPageBreak which adds IDispose to the interface
	/// </summary>
	public interface IVPageBreak : Microsoft.Office.Interop.Excel.VPageBreak, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.VPageBreak Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HPageBreak which adds IDispose to the interface
	/// </summary>
	public interface IHPageBreak : Microsoft.Office.Interop.Excel.HPageBreak, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.HPageBreak Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HPageBreaks which adds IDispose to the interface
	/// </summary>
	public interface IHPageBreaks : Microsoft.Office.Interop.Excel.HPageBreaks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.HPageBreaks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for VPageBreaks which adds IDispose to the interface
	/// </summary>
	public interface IVPageBreaks : Microsoft.Office.Interop.Excel.VPageBreaks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.VPageBreaks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RecentFile which adds IDispose to the interface
	/// </summary>
	public interface IRecentFile : Microsoft.Office.Interop.Excel.RecentFile, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RecentFile Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RecentFiles which adds IDispose to the interface
	/// </summary>
	public interface IRecentFiles : Microsoft.Office.Interop.Excel.RecentFiles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RecentFiles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocEvents which adds IDispose to the interface
	/// </summary>
	public interface IDocEvents : Microsoft.Office.Interop.Excel.DocEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DocEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Style which adds IDispose to the interface
	/// </summary>
	public interface IStyle : Microsoft.Office.Interop.Excel.Style, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Style Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Styles which adds IDispose to the interface
	/// </summary>
	public interface IStyles : Microsoft.Office.Interop.Excel.Styles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Styles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Borders which adds IDispose to the interface
	/// </summary>
	public interface IBorders : Microsoft.Office.Interop.Excel.Borders, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Borders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIn which adds IDispose to the interface
	/// </summary>
	public interface IAddIn : Microsoft.Office.Interop.Excel.AddIn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AddIn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIns which adds IDispose to the interface
	/// </summary>
	public interface IAddIns : Microsoft.Office.Interop.Excel.AddIns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AddIns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Toolbar which adds IDispose to the interface
	/// </summary>
	public interface IToolbar : Microsoft.Office.Interop.Excel.Toolbar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Toolbar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Toolbars which adds IDispose to the interface
	/// </summary>
	public interface IToolbars : Microsoft.Office.Interop.Excel.Toolbars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Toolbars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ToolbarButton which adds IDispose to the interface
	/// </summary>
	public interface IToolbarButton : Microsoft.Office.Interop.Excel.ToolbarButton, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ToolbarButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ToolbarButtons which adds IDispose to the interface
	/// </summary>
	public interface IToolbarButtons : Microsoft.Office.Interop.Excel.ToolbarButtons, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ToolbarButtons Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Areas which adds IDispose to the interface
	/// </summary>
	public interface IAreas : Microsoft.Office.Interop.Excel.Areas, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Areas Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkbookEvents which adds IDispose to the interface
	/// </summary>
	public interface IWorkbookEvents : Microsoft.Office.Interop.Excel.WorkbookEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.WorkbookEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MenuBars which adds IDispose to the interface
	/// </summary>
	public interface IMenuBars : Microsoft.Office.Interop.Excel.MenuBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.MenuBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MenuBar which adds IDispose to the interface
	/// </summary>
	public interface IMenuBar : Microsoft.Office.Interop.Excel.MenuBar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.MenuBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Menus which adds IDispose to the interface
	/// </summary>
	public interface IMenus : Microsoft.Office.Interop.Excel.Menus, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Menus Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Menu which adds IDispose to the interface
	/// </summary>
	public interface IMenu : Microsoft.Office.Interop.Excel.Menu, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Menu Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MenuItems which adds IDispose to the interface
	/// </summary>
	public interface IMenuItems : Microsoft.Office.Interop.Excel.MenuItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.MenuItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MenuItem which adds IDispose to the interface
	/// </summary>
	public interface IMenuItem : Microsoft.Office.Interop.Excel.MenuItem, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.MenuItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Charts which adds IDispose to the interface
	/// </summary>
	public interface ICharts : Microsoft.Office.Interop.Excel.Charts, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Charts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DrawingObjects which adds IDispose to the interface
	/// </summary>
	public interface IDrawingObjects : Microsoft.Office.Interop.Excel.DrawingObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DrawingObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotCache which adds IDispose to the interface
	/// </summary>
	public interface IPivotCache : Microsoft.Office.Interop.Excel.PivotCache, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotCache Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotCaches which adds IDispose to the interface
	/// </summary>
	public interface IPivotCaches : Microsoft.Office.Interop.Excel.PivotCaches, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotCaches Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotFormula which adds IDispose to the interface
	/// </summary>
	public interface IPivotFormula : Microsoft.Office.Interop.Excel.PivotFormula, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotFormula Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotFormulas which adds IDispose to the interface
	/// </summary>
	public interface IPivotFormulas : Microsoft.Office.Interop.Excel.PivotFormulas, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotFormulas Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotTable which adds IDispose to the interface
	/// </summary>
	public interface IPivotTable : Microsoft.Office.Interop.Excel.PivotTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotTables which adds IDispose to the interface
	/// </summary>
	public interface IPivotTables : Microsoft.Office.Interop.Excel.PivotTables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotTables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotField which adds IDispose to the interface
	/// </summary>
	public interface IPivotField : Microsoft.Office.Interop.Excel.PivotField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotFields which adds IDispose to the interface
	/// </summary>
	public interface IPivotFields : Microsoft.Office.Interop.Excel.PivotFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalculatedFields which adds IDispose to the interface
	/// </summary>
	public interface ICalculatedFields : Microsoft.Office.Interop.Excel.CalculatedFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CalculatedFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotItem which adds IDispose to the interface
	/// </summary>
	public interface IPivotItem : Microsoft.Office.Interop.Excel.PivotItem, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotItems which adds IDispose to the interface
	/// </summary>
	public interface IPivotItems : Microsoft.Office.Interop.Excel.PivotItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalculatedItems which adds IDispose to the interface
	/// </summary>
	public interface ICalculatedItems : Microsoft.Office.Interop.Excel.CalculatedItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CalculatedItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Characters which adds IDispose to the interface
	/// </summary>
	public interface ICharacters : Microsoft.Office.Interop.Excel.Characters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Characters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Dialogs which adds IDispose to the interface
	/// </summary>
	public interface IDialogs : Microsoft.Office.Interop.Excel.Dialogs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Dialogs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Dialog which adds IDispose to the interface
	/// </summary>
	public interface IDialog : Microsoft.Office.Interop.Excel.Dialog, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Dialog Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SoundNote which adds IDispose to the interface
	/// </summary>
	public interface ISoundNote : Microsoft.Office.Interop.Excel.SoundNote, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SoundNote Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Button which adds IDispose to the interface
	/// </summary>
	public interface IButton : Microsoft.Office.Interop.Excel.Button, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Button Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Buttons which adds IDispose to the interface
	/// </summary>
	public interface IButtons : Microsoft.Office.Interop.Excel.Buttons, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Buttons Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CheckBox which adds IDispose to the interface
	/// </summary>
	public interface ICheckBox : Microsoft.Office.Interop.Excel.CheckBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CheckBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CheckBoxes which adds IDispose to the interface
	/// </summary>
	public interface ICheckBoxes : Microsoft.Office.Interop.Excel.CheckBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CheckBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OptionButton which adds IDispose to the interface
	/// </summary>
	public interface IOptionButton : Microsoft.Office.Interop.Excel.OptionButton, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OptionButton Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OptionButtons which adds IDispose to the interface
	/// </summary>
	public interface IOptionButtons : Microsoft.Office.Interop.Excel.OptionButtons, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OptionButtons Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EditBox which adds IDispose to the interface
	/// </summary>
	public interface IEditBox : Microsoft.Office.Interop.Excel.EditBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.EditBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EditBoxes which adds IDispose to the interface
	/// </summary>
	public interface IEditBoxes : Microsoft.Office.Interop.Excel.EditBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.EditBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ScrollBar which adds IDispose to the interface
	/// </summary>
	public interface IScrollBar : Microsoft.Office.Interop.Excel.ScrollBar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ScrollBar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ScrollBars which adds IDispose to the interface
	/// </summary>
	public interface IScrollBars : Microsoft.Office.Interop.Excel.ScrollBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ScrollBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListBox which adds IDispose to the interface
	/// </summary>
	public interface IListBox : Microsoft.Office.Interop.Excel.ListBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListBoxes which adds IDispose to the interface
	/// </summary>
	public interface IListBoxes : Microsoft.Office.Interop.Excel.ListBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupBox which adds IDispose to the interface
	/// </summary>
	public interface IGroupBox : Microsoft.Office.Interop.Excel.GroupBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.GroupBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupBoxes which adds IDispose to the interface
	/// </summary>
	public interface IGroupBoxes : Microsoft.Office.Interop.Excel.GroupBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.GroupBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropDown which adds IDispose to the interface
	/// </summary>
	public interface IDropDown : Microsoft.Office.Interop.Excel.DropDown, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DropDown Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropDowns which adds IDispose to the interface
	/// </summary>
	public interface IDropDowns : Microsoft.Office.Interop.Excel.DropDowns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DropDowns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Spinner which adds IDispose to the interface
	/// </summary>
	public interface ISpinner : Microsoft.Office.Interop.Excel.Spinner, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Spinner Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Spinners which adds IDispose to the interface
	/// </summary>
	public interface ISpinners : Microsoft.Office.Interop.Excel.Spinners, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Spinners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DialogFrame which adds IDispose to the interface
	/// </summary>
	public interface IDialogFrame : Microsoft.Office.Interop.Excel.DialogFrame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DialogFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Label which adds IDispose to the interface
	/// </summary>
	public interface ILabel : Microsoft.Office.Interop.Excel.Label, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Label Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Labels which adds IDispose to the interface
	/// </summary>
	public interface ILabels : Microsoft.Office.Interop.Excel.Labels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Labels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Panes which adds IDispose to the interface
	/// </summary>
	public interface IPanes : Microsoft.Office.Interop.Excel.Panes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Panes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pane which adds IDispose to the interface
	/// </summary>
	public interface IPane : Microsoft.Office.Interop.Excel.Pane, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Pane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Scenarios which adds IDispose to the interface
	/// </summary>
	public interface IScenarios : Microsoft.Office.Interop.Excel.Scenarios, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Scenarios Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Scenario which adds IDispose to the interface
	/// </summary>
	public interface IScenario : Microsoft.Office.Interop.Excel.Scenario, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Scenario Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupObject which adds IDispose to the interface
	/// </summary>
	public interface IGroupObject : Microsoft.Office.Interop.Excel.GroupObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.GroupObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupObjects which adds IDispose to the interface
	/// </summary>
	public interface IGroupObjects : Microsoft.Office.Interop.Excel.GroupObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.GroupObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Line which adds IDispose to the interface
	/// </summary>
	public interface ILine : Microsoft.Office.Interop.Excel.Line, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Line Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Lines which adds IDispose to the interface
	/// </summary>
	public interface ILines : Microsoft.Office.Interop.Excel.Lines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Lines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rectangle which adds IDispose to the interface
	/// </summary>
	public interface IRectangle : Microsoft.Office.Interop.Excel.Rectangle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Rectangle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rectangles which adds IDispose to the interface
	/// </summary>
	public interface IRectangles : Microsoft.Office.Interop.Excel.Rectangles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Rectangles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Oval which adds IDispose to the interface
	/// </summary>
	public interface IOval : Microsoft.Office.Interop.Excel.Oval, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Oval Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Ovals which adds IDispose to the interface
	/// </summary>
	public interface IOvals : Microsoft.Office.Interop.Excel.Ovals, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Ovals Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Arc which adds IDispose to the interface
	/// </summary>
	public interface IArc : Microsoft.Office.Interop.Excel.Arc, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Arc Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Arcs which adds IDispose to the interface
	/// </summary>
	public interface IArcs : Microsoft.Office.Interop.Excel.Arcs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Arcs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEObjectEvents which adds IDispose to the interface
	/// </summary>
	public interface IOLEObjectEvents : Microsoft.Office.Interop.Excel.OLEObjectEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEObjectEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _OLEObject which adds IDispose to the interface
	/// </summary>
	public interface I_OLEObject : Microsoft.Office.Interop.Excel._OLEObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._OLEObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEObjects which adds IDispose to the interface
	/// </summary>
	public interface IOLEObjects : Microsoft.Office.Interop.Excel.OLEObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextBox which adds IDispose to the interface
	/// </summary>
	public interface ITextBox : Microsoft.Office.Interop.Excel.TextBox, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TextBox Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextBoxes which adds IDispose to the interface
	/// </summary>
	public interface ITextBoxes : Microsoft.Office.Interop.Excel.TextBoxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TextBoxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Picture which adds IDispose to the interface
	/// </summary>
	public interface IPicture : Microsoft.Office.Interop.Excel.Picture, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Picture Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pictures which adds IDispose to the interface
	/// </summary>
	public interface IPictures : Microsoft.Office.Interop.Excel.Pictures, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Pictures Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Drawing which adds IDispose to the interface
	/// </summary>
	public interface IDrawing : Microsoft.Office.Interop.Excel.Drawing, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Drawing Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Drawings which adds IDispose to the interface
	/// </summary>
	public interface IDrawings : Microsoft.Office.Interop.Excel.Drawings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Drawings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RoutingSlip which adds IDispose to the interface
	/// </summary>
	public interface IRoutingSlip : Microsoft.Office.Interop.Excel.RoutingSlip, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RoutingSlip Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Outline which adds IDispose to the interface
	/// </summary>
	public interface IOutline : Microsoft.Office.Interop.Excel.Outline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Outline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Module which adds IDispose to the interface
	/// </summary>
	public interface IModule : Microsoft.Office.Interop.Excel.Module, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Module Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Modules which adds IDispose to the interface
	/// </summary>
	public interface IModules : Microsoft.Office.Interop.Excel.Modules, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Modules Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DialogSheet which adds IDispose to the interface
	/// </summary>
	public interface IDialogSheet : Microsoft.Office.Interop.Excel.DialogSheet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DialogSheet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DialogSheets which adds IDispose to the interface
	/// </summary>
	public interface IDialogSheets : Microsoft.Office.Interop.Excel.DialogSheets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DialogSheets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Worksheets which adds IDispose to the interface
	/// </summary>
	public interface IWorksheets : Microsoft.Office.Interop.Excel.Worksheets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Worksheets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PageSetup which adds IDispose to the interface
	/// </summary>
	public interface IPageSetup : Microsoft.Office.Interop.Excel.PageSetup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PageSetup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Names which adds IDispose to the interface
	/// </summary>
	public interface INames : Microsoft.Office.Interop.Excel.Names, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Names Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Name which adds IDispose to the interface
	/// </summary>
	public interface IName : Microsoft.Office.Interop.Excel.Name, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Name Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartObject which adds IDispose to the interface
	/// </summary>
	public interface IChartObject : Microsoft.Office.Interop.Excel.ChartObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartObjects which adds IDispose to the interface
	/// </summary>
	public interface IChartObjects : Microsoft.Office.Interop.Excel.ChartObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Mailer which adds IDispose to the interface
	/// </summary>
	public interface IMailer : Microsoft.Office.Interop.Excel.Mailer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Mailer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomViews which adds IDispose to the interface
	/// </summary>
	public interface ICustomViews : Microsoft.Office.Interop.Excel.CustomViews, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CustomViews Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomView which adds IDispose to the interface
	/// </summary>
	public interface ICustomView : Microsoft.Office.Interop.Excel.CustomView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CustomView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormatConditions which adds IDispose to the interface
	/// </summary>
	public interface IFormatConditions : Microsoft.Office.Interop.Excel.FormatConditions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FormatConditions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormatCondition which adds IDispose to the interface
	/// </summary>
	public interface IFormatCondition : Microsoft.Office.Interop.Excel.FormatCondition, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FormatCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Comments which adds IDispose to the interface
	/// </summary>
	public interface IComments : Microsoft.Office.Interop.Excel.Comments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Comments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Comment which adds IDispose to the interface
	/// </summary>
	public interface IComment : Microsoft.Office.Interop.Excel.Comment, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Comment Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RefreshEvents which adds IDispose to the interface
	/// </summary>
	public interface IRefreshEvents : Microsoft.Office.Interop.Excel.RefreshEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RefreshEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _QueryTable which adds IDispose to the interface
	/// </summary>
	public interface I_QueryTable : Microsoft.Office.Interop.Excel._QueryTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel._QueryTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for QueryTables which adds IDispose to the interface
	/// </summary>
	public interface IQueryTables : Microsoft.Office.Interop.Excel.QueryTables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.QueryTables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Parameter which adds IDispose to the interface
	/// </summary>
	public interface IParameter : Microsoft.Office.Interop.Excel.Parameter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Parameter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Parameters which adds IDispose to the interface
	/// </summary>
	public interface IParameters : Microsoft.Office.Interop.Excel.Parameters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Parameters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODBCError which adds IDispose to the interface
	/// </summary>
	public interface IODBCError : Microsoft.Office.Interop.Excel.ODBCError, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ODBCError Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODBCErrors which adds IDispose to the interface
	/// </summary>
	public interface IODBCErrors : Microsoft.Office.Interop.Excel.ODBCErrors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ODBCErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Validation which adds IDispose to the interface
	/// </summary>
	public interface IValidation : Microsoft.Office.Interop.Excel.Validation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Validation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Hyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IHyperlinks : Microsoft.Office.Interop.Excel.Hyperlinks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Hyperlinks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Hyperlink which adds IDispose to the interface
	/// </summary>
	public interface IHyperlink : Microsoft.Office.Interop.Excel.Hyperlink, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Hyperlink Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoFilter which adds IDispose to the interface
	/// </summary>
	public interface IAutoFilter : Microsoft.Office.Interop.Excel.AutoFilter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AutoFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Filters which adds IDispose to the interface
	/// </summary>
	public interface IFilters : Microsoft.Office.Interop.Excel.Filters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Filters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Filter which adds IDispose to the interface
	/// </summary>
	public interface IFilter : Microsoft.Office.Interop.Excel.Filter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Filter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCorrect which adds IDispose to the interface
	/// </summary>
	public interface IAutoCorrect : Microsoft.Office.Interop.Excel.AutoCorrect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AutoCorrect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Border which adds IDispose to the interface
	/// </summary>
	public interface IBorder : Microsoft.Office.Interop.Excel.Border, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Border Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Interior which adds IDispose to the interface
	/// </summary>
	public interface IInterior : Microsoft.Office.Interop.Excel.Interior, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Interior Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFillFormat : Microsoft.Office.Interop.Excel.ChartFillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartFillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartColorFormat : Microsoft.Office.Interop.Excel.ChartColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axis which adds IDispose to the interface
	/// </summary>
	public interface IAxis : Microsoft.Office.Interop.Excel.Axis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Axis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IChartTitle : Microsoft.Office.Interop.Excel.ChartTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IAxisTitle : Microsoft.Office.Interop.Excel.AxisTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AxisTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IChartGroup : Microsoft.Office.Interop.Excel.ChartGroup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IChartGroups : Microsoft.Office.Interop.Excel.ChartGroups, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axes which adds IDispose to the interface
	/// </summary>
	public interface IAxes : Microsoft.Office.Interop.Excel.Axes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Axes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Points which adds IDispose to the interface
	/// </summary>
	public interface IPoints : Microsoft.Office.Interop.Excel.Points, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Points Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Point which adds IDispose to the interface
	/// </summary>
	public interface IPoint : Microsoft.Office.Interop.Excel.Point, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Point Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Series which adds IDispose to the interface
	/// </summary>
	public interface ISeries : Microsoft.Office.Interop.Excel.Series, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Series Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface ISeriesCollection : Microsoft.Office.Interop.Excel.SeriesCollection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SeriesCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataLabel which adds IDispose to the interface
	/// </summary>
	public interface IDataLabel : Microsoft.Office.Interop.Excel.DataLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DataLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataLabels which adds IDispose to the interface
	/// </summary>
	public interface IDataLabels : Microsoft.Office.Interop.Excel.DataLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DataLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntry which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntry : Microsoft.Office.Interop.Excel.LegendEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LegendEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntries which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntries : Microsoft.Office.Interop.Excel.LegendEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LegendEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendKey which adds IDispose to the interface
	/// </summary>
	public interface ILegendKey : Microsoft.Office.Interop.Excel.LegendKey, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LegendKey Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendlines which adds IDispose to the interface
	/// </summary>
	public interface ITrendlines : Microsoft.Office.Interop.Excel.Trendlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Trendlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendline which adds IDispose to the interface
	/// </summary>
	public interface ITrendline : Microsoft.Office.Interop.Excel.Trendline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Trendline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Corners which adds IDispose to the interface
	/// </summary>
	public interface ICorners : Microsoft.Office.Interop.Excel.Corners, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Corners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesLines which adds IDispose to the interface
	/// </summary>
	public interface ISeriesLines : Microsoft.Office.Interop.Excel.SeriesLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SeriesLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IHiLoLines : Microsoft.Office.Interop.Excel.HiLoLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.HiLoLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Gridlines which adds IDispose to the interface
	/// </summary>
	public interface IGridlines : Microsoft.Office.Interop.Excel.Gridlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Gridlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropLines which adds IDispose to the interface
	/// </summary>
	public interface IDropLines : Microsoft.Office.Interop.Excel.DropLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DropLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LeaderLines which adds IDispose to the interface
	/// </summary>
	public interface ILeaderLines : Microsoft.Office.Interop.Excel.LeaderLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LeaderLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UpBars which adds IDispose to the interface
	/// </summary>
	public interface IUpBars : Microsoft.Office.Interop.Excel.UpBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.UpBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DownBars which adds IDispose to the interface
	/// </summary>
	public interface IDownBars : Microsoft.Office.Interop.Excel.DownBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DownBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Floor which adds IDispose to the interface
	/// </summary>
	public interface IFloor : Microsoft.Office.Interop.Excel.Floor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Floor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Walls which adds IDispose to the interface
	/// </summary>
	public interface IWalls : Microsoft.Office.Interop.Excel.Walls, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Walls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TickLabels which adds IDispose to the interface
	/// </summary>
	public interface ITickLabels : Microsoft.Office.Interop.Excel.TickLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TickLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PlotArea which adds IDispose to the interface
	/// </summary>
	public interface IPlotArea : Microsoft.Office.Interop.Excel.PlotArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PlotArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartArea which adds IDispose to the interface
	/// </summary>
	public interface IChartArea : Microsoft.Office.Interop.Excel.ChartArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Legend which adds IDispose to the interface
	/// </summary>
	public interface ILegend : Microsoft.Office.Interop.Excel.Legend, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Legend Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IErrorBars : Microsoft.Office.Interop.Excel.ErrorBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ErrorBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataTable which adds IDispose to the interface
	/// </summary>
	public interface IDataTable : Microsoft.Office.Interop.Excel.DataTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DataTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Phonetic which adds IDispose to the interface
	/// </summary>
	public interface IPhonetic : Microsoft.Office.Interop.Excel.Phonetic, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Phonetic Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shape which adds IDispose to the interface
	/// </summary>
	public interface IShape : Microsoft.Office.Interop.Excel.Shape, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Shape Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shapes which adds IDispose to the interface
	/// </summary>
	public interface IShapes : Microsoft.Office.Interop.Excel.Shapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Shapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IShapeRange : Microsoft.Office.Interop.Excel.ShapeRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ShapeRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IGroupShapes : Microsoft.Office.Interop.Excel.GroupShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.GroupShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame : Microsoft.Office.Interop.Excel.TextFrame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TextFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IConnectorFormat : Microsoft.Office.Interop.Excel.ConnectorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ConnectorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IFreeformBuilder : Microsoft.Office.Interop.Excel.FreeformBuilder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FreeformBuilder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ControlFormat which adds IDispose to the interface
	/// </summary>
	public interface IControlFormat : Microsoft.Office.Interop.Excel.ControlFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ControlFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEFormat which adds IDispose to the interface
	/// </summary>
	public interface IOLEFormat : Microsoft.Office.Interop.Excel.OLEFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LinkFormat which adds IDispose to the interface
	/// </summary>
	public interface ILinkFormat : Microsoft.Office.Interop.Excel.LinkFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LinkFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PublishObjects which adds IDispose to the interface
	/// </summary>
	public interface IPublishObjects : Microsoft.Office.Interop.Excel.PublishObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PublishObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEDBError which adds IDispose to the interface
	/// </summary>
	public interface IOLEDBError : Microsoft.Office.Interop.Excel.OLEDBError, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEDBError Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEDBErrors which adds IDispose to the interface
	/// </summary>
	public interface IOLEDBErrors : Microsoft.Office.Interop.Excel.OLEDBErrors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEDBErrors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Phonetics which adds IDispose to the interface
	/// </summary>
	public interface IPhonetics : Microsoft.Office.Interop.Excel.Phonetics, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Phonetics Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotLayout which adds IDispose to the interface
	/// </summary>
	public interface IPivotLayout : Microsoft.Office.Interop.Excel.PivotLayout, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotLayout Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IDisplayUnitLabel : Microsoft.Office.Interop.Excel.DisplayUnitLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DisplayUnitLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CellFormat which adds IDispose to the interface
	/// </summary>
	public interface ICellFormat : Microsoft.Office.Interop.Excel.CellFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CellFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UsedObjects which adds IDispose to the interface
	/// </summary>
	public interface IUsedObjects : Microsoft.Office.Interop.Excel.UsedObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.UsedObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomProperties which adds IDispose to the interface
	/// </summary>
	public interface ICustomProperties : Microsoft.Office.Interop.Excel.CustomProperties, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CustomProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomProperty which adds IDispose to the interface
	/// </summary>
	public interface ICustomProperty : Microsoft.Office.Interop.Excel.CustomProperty, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CustomProperty Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalculatedMembers which adds IDispose to the interface
	/// </summary>
	public interface ICalculatedMembers : Microsoft.Office.Interop.Excel.CalculatedMembers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CalculatedMembers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalculatedMember which adds IDispose to the interface
	/// </summary>
	public interface ICalculatedMember : Microsoft.Office.Interop.Excel.CalculatedMember, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.CalculatedMember Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Watches which adds IDispose to the interface
	/// </summary>
	public interface IWatches : Microsoft.Office.Interop.Excel.Watches, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Watches Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Watch which adds IDispose to the interface
	/// </summary>
	public interface IWatch : Microsoft.Office.Interop.Excel.Watch, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Watch Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotCell which adds IDispose to the interface
	/// </summary>
	public interface IPivotCell : Microsoft.Office.Interop.Excel.PivotCell, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotCell Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Graphic which adds IDispose to the interface
	/// </summary>
	public interface IGraphic : Microsoft.Office.Interop.Excel.Graphic, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Graphic Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoRecover which adds IDispose to the interface
	/// </summary>
	public interface IAutoRecover : Microsoft.Office.Interop.Excel.AutoRecover, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AutoRecover Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ErrorCheckingOptions which adds IDispose to the interface
	/// </summary>
	public interface IErrorCheckingOptions : Microsoft.Office.Interop.Excel.ErrorCheckingOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ErrorCheckingOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Errors which adds IDispose to the interface
	/// </summary>
	public interface IErrors : Microsoft.Office.Interop.Excel.Errors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Errors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Error which adds IDispose to the interface
	/// </summary>
	public interface IError : Microsoft.Office.Interop.Excel.Error, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Error Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagAction which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagAction : Microsoft.Office.Interop.Excel.SmartTagAction, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTagAction Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagActions which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagActions : Microsoft.Office.Interop.Excel.SmartTagActions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTagActions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTag which adds IDispose to the interface
	/// </summary>
	public interface ISmartTag : Microsoft.Office.Interop.Excel.SmartTag, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTag Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTags which adds IDispose to the interface
	/// </summary>
	public interface ISmartTags : Microsoft.Office.Interop.Excel.SmartTags, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTags Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagRecognizer : Microsoft.Office.Interop.Excel.SmartTagRecognizer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTagRecognizer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagRecognizers : Microsoft.Office.Interop.Excel.SmartTagRecognizers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTagRecognizers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SmartTagOptions which adds IDispose to the interface
	/// </summary>
	public interface ISmartTagOptions : Microsoft.Office.Interop.Excel.SmartTagOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SmartTagOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SpellingOptions which adds IDispose to the interface
	/// </summary>
	public interface ISpellingOptions : Microsoft.Office.Interop.Excel.SpellingOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SpellingOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Speech which adds IDispose to the interface
	/// </summary>
	public interface ISpeech : Microsoft.Office.Interop.Excel.Speech, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Speech Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Protection which adds IDispose to the interface
	/// </summary>
	public interface IProtection : Microsoft.Office.Interop.Excel.Protection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Protection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotItemList which adds IDispose to the interface
	/// </summary>
	public interface IPivotItemList : Microsoft.Office.Interop.Excel.PivotItemList, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotItemList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Tab which adds IDispose to the interface
	/// </summary>
	public interface ITab : Microsoft.Office.Interop.Excel.Tab, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Tab Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AllowEditRanges which adds IDispose to the interface
	/// </summary>
	public interface IAllowEditRanges : Microsoft.Office.Interop.Excel.AllowEditRanges, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AllowEditRanges Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AllowEditRange which adds IDispose to the interface
	/// </summary>
	public interface IAllowEditRange : Microsoft.Office.Interop.Excel.AllowEditRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AllowEditRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserAccessList which adds IDispose to the interface
	/// </summary>
	public interface IUserAccessList : Microsoft.Office.Interop.Excel.UserAccessList, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.UserAccessList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UserAccess which adds IDispose to the interface
	/// </summary>
	public interface IUserAccess : Microsoft.Office.Interop.Excel.UserAccess, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.UserAccess Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RTD which adds IDispose to the interface
	/// </summary>
	public interface IRTD : Microsoft.Office.Interop.Excel.RTD, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RTD Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Diagram which adds IDispose to the interface
	/// </summary>
	public interface IDiagram : Microsoft.Office.Interop.Excel.Diagram, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Diagram Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListObjects which adds IDispose to the interface
	/// </summary>
	public interface IListObjects : Microsoft.Office.Interop.Excel.ListObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListObject which adds IDispose to the interface
	/// </summary>
	public interface IListObject : Microsoft.Office.Interop.Excel.ListObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListColumns which adds IDispose to the interface
	/// </summary>
	public interface IListColumns : Microsoft.Office.Interop.Excel.ListColumns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListColumns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListColumn which adds IDispose to the interface
	/// </summary>
	public interface IListColumn : Microsoft.Office.Interop.Excel.ListColumn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListColumn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListRows which adds IDispose to the interface
	/// </summary>
	public interface IListRows : Microsoft.Office.Interop.Excel.ListRows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListRows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListRow which adds IDispose to the interface
	/// </summary>
	public interface IListRow : Microsoft.Office.Interop.Excel.ListRow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListRow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlNamespace which adds IDispose to the interface
	/// </summary>
	public interface IXmlNamespace : Microsoft.Office.Interop.Excel.XmlNamespace, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlNamespace Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlNamespaces which adds IDispose to the interface
	/// </summary>
	public interface IXmlNamespaces : Microsoft.Office.Interop.Excel.XmlNamespaces, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlNamespaces Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlDataBinding which adds IDispose to the interface
	/// </summary>
	public interface IXmlDataBinding : Microsoft.Office.Interop.Excel.XmlDataBinding, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlDataBinding Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlSchema which adds IDispose to the interface
	/// </summary>
	public interface IXmlSchema : Microsoft.Office.Interop.Excel.XmlSchema, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlSchema Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlSchemas which adds IDispose to the interface
	/// </summary>
	public interface IXmlSchemas : Microsoft.Office.Interop.Excel.XmlSchemas, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlSchemas Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlMap which adds IDispose to the interface
	/// </summary>
	public interface IXmlMap : Microsoft.Office.Interop.Excel.XmlMap, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlMap Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XmlMaps which adds IDispose to the interface
	/// </summary>
	public interface IXmlMaps : Microsoft.Office.Interop.Excel.XmlMaps, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XmlMaps Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ListDataFormat which adds IDispose to the interface
	/// </summary>
	public interface IListDataFormat : Microsoft.Office.Interop.Excel.ListDataFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ListDataFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for XPath which adds IDispose to the interface
	/// </summary>
	public interface IXPath : Microsoft.Office.Interop.Excel.XPath, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.XPath Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotLineCells which adds IDispose to the interface
	/// </summary>
	public interface IPivotLineCells : Microsoft.Office.Interop.Excel.PivotLineCells, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotLineCells Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotLine which adds IDispose to the interface
	/// </summary>
	public interface IPivotLine : Microsoft.Office.Interop.Excel.PivotLine, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotLine Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotLines which adds IDispose to the interface
	/// </summary>
	public interface IPivotLines : Microsoft.Office.Interop.Excel.PivotLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotAxis which adds IDispose to the interface
	/// </summary>
	public interface IPivotAxis : Microsoft.Office.Interop.Excel.PivotAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotFilter which adds IDispose to the interface
	/// </summary>
	public interface IPivotFilter : Microsoft.Office.Interop.Excel.PivotFilter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotFilter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotFilters which adds IDispose to the interface
	/// </summary>
	public interface IPivotFilters : Microsoft.Office.Interop.Excel.PivotFilters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotFilters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkbookConnection which adds IDispose to the interface
	/// </summary>
	public interface IWorkbookConnection : Microsoft.Office.Interop.Excel.WorkbookConnection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.WorkbookConnection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Connections which adds IDispose to the interface
	/// </summary>
	public interface IConnections : Microsoft.Office.Interop.Excel.Connections, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Connections Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorksheetView which adds IDispose to the interface
	/// </summary>
	public interface IWorksheetView : Microsoft.Office.Interop.Excel.WorksheetView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.WorksheetView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartView which adds IDispose to the interface
	/// </summary>
	public interface IChartView : Microsoft.Office.Interop.Excel.ChartView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ModuleView which adds IDispose to the interface
	/// </summary>
	public interface IModuleView : Microsoft.Office.Interop.Excel.ModuleView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ModuleView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DialogSheetView which adds IDispose to the interface
	/// </summary>
	public interface IDialogSheetView : Microsoft.Office.Interop.Excel.DialogSheetView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DialogSheetView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SheetViews which adds IDispose to the interface
	/// </summary>
	public interface ISheetViews : Microsoft.Office.Interop.Excel.SheetViews, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SheetViews Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEDBConnection which adds IDispose to the interface
	/// </summary>
	public interface IOLEDBConnection : Microsoft.Office.Interop.Excel.OLEDBConnection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEDBConnection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ODBCConnection which adds IDispose to the interface
	/// </summary>
	public interface IODBCConnection : Microsoft.Office.Interop.Excel.ODBCConnection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ODBCConnection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Action which adds IDispose to the interface
	/// </summary>
	public interface IAction : Microsoft.Office.Interop.Excel.Action, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Action Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Actions which adds IDispose to the interface
	/// </summary>
	public interface IActions : Microsoft.Office.Interop.Excel.Actions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Actions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FormatColor which adds IDispose to the interface
	/// </summary>
	public interface IFormatColor : Microsoft.Office.Interop.Excel.FormatColor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FormatColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ConditionValue which adds IDispose to the interface
	/// </summary>
	public interface IConditionValue : Microsoft.Office.Interop.Excel.ConditionValue, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ConditionValue Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorScale which adds IDispose to the interface
	/// </summary>
	public interface IColorScale : Microsoft.Office.Interop.Excel.ColorScale, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ColorScale Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorScaleCriteria which adds IDispose to the interface
	/// </summary>
	public interface IColorScaleCriteria : Microsoft.Office.Interop.Excel.ColorScaleCriteria, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ColorScaleCriteria Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorScaleCriterion which adds IDispose to the interface
	/// </summary>
	public interface IColorScaleCriterion : Microsoft.Office.Interop.Excel.ColorScaleCriterion, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ColorScaleCriterion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Databar which adds IDispose to the interface
	/// </summary>
	public interface IDatabar : Microsoft.Office.Interop.Excel.Databar, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Databar Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IconSetCondition which adds IDispose to the interface
	/// </summary>
	public interface IIconSetCondition : Microsoft.Office.Interop.Excel.IconSetCondition, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IconSetCondition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IconCriteria which adds IDispose to the interface
	/// </summary>
	public interface IIconCriteria : Microsoft.Office.Interop.Excel.IconCriteria, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IconCriteria Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IconCriterion which adds IDispose to the interface
	/// </summary>
	public interface IIconCriterion : Microsoft.Office.Interop.Excel.IconCriterion, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IconCriterion Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Icon which adds IDispose to the interface
	/// </summary>
	public interface IIcon : Microsoft.Office.Interop.Excel.Icon, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Icon Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IconSet which adds IDispose to the interface
	/// </summary>
	public interface IIconSet : Microsoft.Office.Interop.Excel.IconSet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IconSet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IconSets which adds IDispose to the interface
	/// </summary>
	public interface IIconSets : Microsoft.Office.Interop.Excel.IconSets, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IconSets Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Top10 which adds IDispose to the interface
	/// </summary>
	public interface ITop10 : Microsoft.Office.Interop.Excel.Top10, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Top10 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AboveAverage which adds IDispose to the interface
	/// </summary>
	public interface IAboveAverage : Microsoft.Office.Interop.Excel.AboveAverage, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AboveAverage Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UniqueValues which adds IDispose to the interface
	/// </summary>
	public interface IUniqueValues : Microsoft.Office.Interop.Excel.UniqueValues, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.UniqueValues Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Ranges which adds IDispose to the interface
	/// </summary>
	public interface IRanges : Microsoft.Office.Interop.Excel.Ranges, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Ranges Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeaderFooter which adds IDispose to the interface
	/// </summary>
	public interface IHeaderFooter : Microsoft.Office.Interop.Excel.HeaderFooter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.HeaderFooter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Page which adds IDispose to the interface
	/// </summary>
	public interface IPage : Microsoft.Office.Interop.Excel.Page, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Page Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pages which adds IDispose to the interface
	/// </summary>
	public interface IPages : Microsoft.Office.Interop.Excel.Pages, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Pages Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ServerViewableItems which adds IDispose to the interface
	/// </summary>
	public interface IServerViewableItems : Microsoft.Office.Interop.Excel.ServerViewableItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ServerViewableItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableStyleElement which adds IDispose to the interface
	/// </summary>
	public interface ITableStyleElement : Microsoft.Office.Interop.Excel.TableStyleElement, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TableStyleElement Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableStyleElements which adds IDispose to the interface
	/// </summary>
	public interface ITableStyleElements : Microsoft.Office.Interop.Excel.TableStyleElements, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TableStyleElements Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableStyle which adds IDispose to the interface
	/// </summary>
	public interface ITableStyle : Microsoft.Office.Interop.Excel.TableStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TableStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableStyles which adds IDispose to the interface
	/// </summary>
	public interface ITableStyles : Microsoft.Office.Interop.Excel.TableStyles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.TableStyles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SortField which adds IDispose to the interface
	/// </summary>
	public interface ISortField : Microsoft.Office.Interop.Excel.SortField, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SortField Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SortFields which adds IDispose to the interface
	/// </summary>
	public interface ISortFields : Microsoft.Office.Interop.Excel.SortFields, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SortFields Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sort which adds IDispose to the interface
	/// </summary>
	public interface ISort : Microsoft.Office.Interop.Excel.Sort, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Sort Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Research which adds IDispose to the interface
	/// </summary>
	public interface IResearch : Microsoft.Office.Interop.Excel.Research, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Research Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorStop which adds IDispose to the interface
	/// </summary>
	public interface IColorStop : Microsoft.Office.Interop.Excel.ColorStop, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ColorStop Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorStops which adds IDispose to the interface
	/// </summary>
	public interface IColorStops : Microsoft.Office.Interop.Excel.ColorStops, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ColorStops Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LinearGradient which adds IDispose to the interface
	/// </summary>
	public interface ILinearGradient : Microsoft.Office.Interop.Excel.LinearGradient, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.LinearGradient Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RectangularGradient which adds IDispose to the interface
	/// </summary>
	public interface IRectangularGradient : Microsoft.Office.Interop.Excel.RectangularGradient, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RectangularGradient Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MultiThreadedCalculation which adds IDispose to the interface
	/// </summary>
	public interface IMultiThreadedCalculation : Microsoft.Office.Interop.Excel.MultiThreadedCalculation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.MultiThreadedCalculation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFormat : Microsoft.Office.Interop.Excel.ChartFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileExportConverter which adds IDispose to the interface
	/// </summary>
	public interface IFileExportConverter : Microsoft.Office.Interop.Excel.FileExportConverter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FileExportConverter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileExportConverters which adds IDispose to the interface
	/// </summary>
	public interface IFileExportConverters : Microsoft.Office.Interop.Excel.FileExportConverters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.FileExportConverters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIns2 which adds IDispose to the interface
	/// </summary>
	public interface IAddIns2 : Microsoft.Office.Interop.Excel.AddIns2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AddIns2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparklineGroups which adds IDispose to the interface
	/// </summary>
	public interface ISparklineGroups : Microsoft.Office.Interop.Excel.SparklineGroups, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparklineGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparklineGroup which adds IDispose to the interface
	/// </summary>
	public interface ISparklineGroup : Microsoft.Office.Interop.Excel.SparklineGroup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparklineGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparkPoints which adds IDispose to the interface
	/// </summary>
	public interface ISparkPoints : Microsoft.Office.Interop.Excel.SparkPoints, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparkPoints Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sparkline which adds IDispose to the interface
	/// </summary>
	public interface ISparkline : Microsoft.Office.Interop.Excel.Sparkline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Sparkline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparkAxes which adds IDispose to the interface
	/// </summary>
	public interface ISparkAxes : Microsoft.Office.Interop.Excel.SparkAxes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparkAxes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparkHorizontalAxis which adds IDispose to the interface
	/// </summary>
	public interface ISparkHorizontalAxis : Microsoft.Office.Interop.Excel.SparkHorizontalAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparkHorizontalAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparkVerticalAxis which adds IDispose to the interface
	/// </summary>
	public interface ISparkVerticalAxis : Microsoft.Office.Interop.Excel.SparkVerticalAxis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparkVerticalAxis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SparkColor which adds IDispose to the interface
	/// </summary>
	public interface ISparkColor : Microsoft.Office.Interop.Excel.SparkColor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SparkColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataBarBorder which adds IDispose to the interface
	/// </summary>
	public interface IDataBarBorder : Microsoft.Office.Interop.Excel.DataBarBorder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DataBarBorder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NegativeBarFormat which adds IDispose to the interface
	/// </summary>
	public interface INegativeBarFormat : Microsoft.Office.Interop.Excel.NegativeBarFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.NegativeBarFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ValueChange which adds IDispose to the interface
	/// </summary>
	public interface IValueChange : Microsoft.Office.Interop.Excel.ValueChange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ValueChange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PivotTableChangeList which adds IDispose to the interface
	/// </summary>
	public interface IPivotTableChangeList : Microsoft.Office.Interop.Excel.PivotTableChangeList, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.PivotTableChangeList Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DisplayFormat which adds IDispose to the interface
	/// </summary>
	public interface IDisplayFormat : Microsoft.Office.Interop.Excel.DisplayFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DisplayFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerCaches which adds IDispose to the interface
	/// </summary>
	public interface ISlicerCaches : Microsoft.Office.Interop.Excel.SlicerCaches, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerCaches Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerCache which adds IDispose to the interface
	/// </summary>
	public interface ISlicerCache : Microsoft.Office.Interop.Excel.SlicerCache, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerCache Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerCacheLevels which adds IDispose to the interface
	/// </summary>
	public interface ISlicerCacheLevels : Microsoft.Office.Interop.Excel.SlicerCacheLevels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerCacheLevels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerCacheLevel which adds IDispose to the interface
	/// </summary>
	public interface ISlicerCacheLevel : Microsoft.Office.Interop.Excel.SlicerCacheLevel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerCacheLevel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Slicers which adds IDispose to the interface
	/// </summary>
	public interface ISlicers : Microsoft.Office.Interop.Excel.Slicers, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Slicers Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Slicer which adds IDispose to the interface
	/// </summary>
	public interface ISlicer : Microsoft.Office.Interop.Excel.Slicer, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Slicer Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerItem which adds IDispose to the interface
	/// </summary>
	public interface ISlicerItem : Microsoft.Office.Interop.Excel.SlicerItem, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerItem Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerItems which adds IDispose to the interface
	/// </summary>
	public interface ISlicerItems : Microsoft.Office.Interop.Excel.SlicerItems, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerItems Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlicerPivotTables which adds IDispose to the interface
	/// </summary>
	public interface ISlicerPivotTables : Microsoft.Office.Interop.Excel.SlicerPivotTables, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.SlicerPivotTables Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
	/// </summary>
	public interface IProtectedViewWindows : Microsoft.Office.Interop.Excel.ProtectedViewWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ProtectedViewWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
	/// </summary>
	public interface IProtectedViewWindow : Microsoft.Office.Interop.Excel.ProtectedViewWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ProtectedViewWindow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for IDummy which adds IDispose to the interface
	/// </summary>
	public interface IIDummy : Microsoft.Office.Interop.Excel.IDummy, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.IDummy Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ICanvasShapes which adds IDispose to the interface
	/// </summary>
	public interface IICanvasShapes : Microsoft.Office.Interop.Excel.ICanvasShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ICanvasShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RefreshEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IRefreshEvents_Event : Microsoft.Office.Interop.Excel.RefreshEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.RefreshEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for QueryTable which adds IDispose to the interface
	/// </summary>
	public interface IQueryTable : Microsoft.Office.Interop.Excel.QueryTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.QueryTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AppEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IAppEvents_Event : Microsoft.Office.Interop.Excel.AppEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.AppEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Application which adds IDispose to the interface
	/// </summary>
	public interface IApplication : Microsoft.Office.Interop.Excel.Application, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IChartEvents_Event : Microsoft.Office.Interop.Excel.ChartEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.ChartEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Chart which adds IDispose to the interface
	/// </summary>
	public interface IChart : Microsoft.Office.Interop.Excel.Chart, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Chart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IDocEvents_Event : Microsoft.Office.Interop.Excel.DocEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.DocEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Worksheet which adds IDispose to the interface
	/// </summary>
	public interface IWorksheet : Microsoft.Office.Interop.Excel.Worksheet, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Worksheet Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Global which adds IDispose to the interface
	/// </summary>
	public interface IGlobal : Microsoft.Office.Interop.Excel.Global, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Global Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WorkbookEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IWorkbookEvents_Event : Microsoft.Office.Interop.Excel.WorkbookEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.WorkbookEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Workbook which adds IDispose to the interface
	/// </summary>
	public interface IWorkbook : Microsoft.Office.Interop.Excel.Workbook, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.Workbook Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEObjectEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOLEObjectEvents_Event : Microsoft.Office.Interop.Excel.OLEObjectEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEObjectEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEObject which adds IDispose to the interface
	/// </summary>
	public interface IOLEObject : Microsoft.Office.Interop.Excel.OLEObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.Excel.OLEObject Resource { get; }
	}

	}