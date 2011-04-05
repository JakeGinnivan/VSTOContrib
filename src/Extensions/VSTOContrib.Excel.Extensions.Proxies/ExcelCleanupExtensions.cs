using VSTOContrib.Extensions.Proxies;

//Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.Excel.Extensions.Proxies
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAdjustments WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Adjustments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Adjustments, Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalloutFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CalloutFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CalloutFormat, Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ColorFormat, Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LineFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LineFormat, Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNode WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ShapeNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ShapeNode, Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNodes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ShapeNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ShapeNodes, Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PictureFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PictureFormat, Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShadowFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ShadowFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ShadowFormat, Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextEffectFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TextEffectFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TextEffectFormat, Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThreeDFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ThreeDFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ThreeDFormat, Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FillFormat, Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DiagramNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DiagramNodes, Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodeChildren WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DiagramNodeChildren resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DiagramNodeChildren, Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNode WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DiagramNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DiagramNode, Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for IRTDUpdateEvent which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRTDUpdateEvent WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRTDUpdateEvent resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRTDUpdateEvent, Interfaces.IIRTDUpdateEvent>();
		}

		/// <summary>
		/// Wrapper interface for IRtdServer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRtdServer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRtdServer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRtdServer, Interfaces.IIRtdServer>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame2 WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TextFrame2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TextFrame2, Interfaces.ITextFrame2>();
		}

		/// <summary>
		/// Wrapper interface for IFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFont WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFont, Interfaces.IIFont>();
		}

		/// <summary>
		/// Wrapper interface for IWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWindow WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWindow, Interfaces.IIWindow>();
		}

		/// <summary>
		/// Wrapper interface for IWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWindows WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWindows, Interfaces.IIWindows>();
		}

		/// <summary>
		/// Wrapper interface for IAppEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAppEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAppEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAppEvents, Interfaces.IIAppEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Application WithComCleanupProxy(this Microsoft.Office.Interop.Excel._Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._Application, Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for IWorksheetFunction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWorksheetFunction WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWorksheetFunction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWorksheetFunction, Interfaces.IIWorksheetFunction>();
		}

		/// <summary>
		/// Wrapper interface for IRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRange, Interfaces.IIRange>();
		}

		/// <summary>
		/// Wrapper interface for IChartEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartEvents, Interfaces.IIChartEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Chart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Chart WithComCleanupProxy(this Microsoft.Office.Interop.Excel._Chart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._Chart, Interfaces.I_Chart>();
		}

		/// <summary>
		/// Wrapper interface for Sheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISheets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Sheets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Sheets, Interfaces.ISheets>();
		}

		/// <summary>
		/// Wrapper interface for IVPageBreak which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIVPageBreak WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IVPageBreak resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IVPageBreak, Interfaces.IIVPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for IHPageBreak which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIHPageBreak WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IHPageBreak resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IHPageBreak, Interfaces.IIHPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for IHPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIHPageBreaks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IHPageBreaks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IHPageBreaks, Interfaces.IIHPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for IVPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIVPageBreaks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IVPageBreaks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IVPageBreaks, Interfaces.IIVPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for IRecentFile which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRecentFile WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRecentFile resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRecentFile, Interfaces.IIRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for IRecentFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRecentFiles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRecentFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRecentFiles, Interfaces.IIRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for IDocEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDocEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDocEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDocEvents, Interfaces.IIDocEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Worksheet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Worksheet WithComCleanupProxy(this Microsoft.Office.Interop.Excel._Worksheet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._Worksheet, Interfaces.I_Worksheet>();
		}

		/// <summary>
		/// Wrapper interface for IStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIStyle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IStyle, Interfaces.IIStyle>();
		}

		/// <summary>
		/// Wrapper interface for IStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIStyles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IStyles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IStyles, Interfaces.IIStyles>();
		}

		/// <summary>
		/// Wrapper interface for IBorders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIBorders WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IBorders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IBorders, Interfaces.IIBorders>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Global WithComCleanupProxy(this Microsoft.Office.Interop.Excel._Global resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._Global, Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for IAddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAddIn WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAddIn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAddIn, Interfaces.IIAddIn>();
		}

		/// <summary>
		/// Wrapper interface for IAddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAddIns WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAddIns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAddIns, Interfaces.IIAddIns>();
		}

		/// <summary>
		/// Wrapper interface for IToolbar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIToolbar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IToolbar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IToolbar, Interfaces.IIToolbar>();
		}

		/// <summary>
		/// Wrapper interface for IToolbars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIToolbars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IToolbars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IToolbars, Interfaces.IIToolbars>();
		}

		/// <summary>
		/// Wrapper interface for IToolbarButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIToolbarButton WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IToolbarButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IToolbarButton, Interfaces.IIToolbarButton>();
		}

		/// <summary>
		/// Wrapper interface for IToolbarButtons which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIToolbarButtons WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IToolbarButtons resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IToolbarButtons, Interfaces.IIToolbarButtons>();
		}

		/// <summary>
		/// Wrapper interface for IAreas which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAreas WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAreas resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAreas, Interfaces.IIAreas>();
		}

		/// <summary>
		/// Wrapper interface for IWorkbookEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWorkbookEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWorkbookEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWorkbookEvents, Interfaces.IIWorkbookEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Workbook which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Workbook WithComCleanupProxy(this Microsoft.Office.Interop.Excel._Workbook resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._Workbook, Interfaces.I_Workbook>();
		}

		/// <summary>
		/// Wrapper interface for Workbooks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkbooks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Workbooks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Workbooks, Interfaces.IWorkbooks>();
		}

		/// <summary>
		/// Wrapper interface for IMenuBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMenuBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMenuBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMenuBars, Interfaces.IIMenuBars>();
		}

		/// <summary>
		/// Wrapper interface for IMenuBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMenuBar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMenuBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMenuBar, Interfaces.IIMenuBar>();
		}

		/// <summary>
		/// Wrapper interface for IMenus which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMenus WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMenus resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMenus, Interfaces.IIMenus>();
		}

		/// <summary>
		/// Wrapper interface for IMenu which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMenu WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMenu resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMenu, Interfaces.IIMenu>();
		}

		/// <summary>
		/// Wrapper interface for IMenuItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMenuItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMenuItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMenuItems, Interfaces.IIMenuItems>();
		}

		/// <summary>
		/// Wrapper interface for IMenuItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMenuItem WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMenuItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMenuItem, Interfaces.IIMenuItem>();
		}

		/// <summary>
		/// Wrapper interface for ICharts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICharts WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICharts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICharts, Interfaces.IICharts>();
		}

		/// <summary>
		/// Wrapper interface for IDrawingObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDrawingObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDrawingObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDrawingObjects, Interfaces.IIDrawingObjects>();
		}

		/// <summary>
		/// Wrapper interface for IPivotCache which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotCache WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotCache resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotCache, Interfaces.IIPivotCache>();
		}

		/// <summary>
		/// Wrapper interface for IPivotCaches which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotCaches WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotCaches resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotCaches, Interfaces.IIPivotCaches>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFormula which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotFormula WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotFormula resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotFormula, Interfaces.IIPivotFormula>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFormulas which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotFormulas WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotFormulas resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotFormulas, Interfaces.IIPivotFormulas>();
		}

		/// <summary>
		/// Wrapper interface for IPivotTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotTable, Interfaces.IIPivotTable>();
		}

		/// <summary>
		/// Wrapper interface for IPivotTables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotTables WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotTables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotTables, Interfaces.IIPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for IPivotField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotField WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotField, Interfaces.IIPivotField>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotFields, Interfaces.IIPivotFields>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICalculatedFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICalculatedFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICalculatedFields, Interfaces.IICalculatedFields>();
		}

		/// <summary>
		/// Wrapper interface for IPivotItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotItem WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotItem, Interfaces.IIPivotItem>();
		}

		/// <summary>
		/// Wrapper interface for IPivotItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotItems, Interfaces.IIPivotItems>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICalculatedItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICalculatedItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICalculatedItems, Interfaces.IICalculatedItems>();
		}

		/// <summary>
		/// Wrapper interface for ICharacters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICharacters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICharacters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICharacters, Interfaces.IICharacters>();
		}

		/// <summary>
		/// Wrapper interface for IDialogs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDialogs WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDialogs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDialogs, Interfaces.IIDialogs>();
		}

		/// <summary>
		/// Wrapper interface for IDialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDialog WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDialog, Interfaces.IIDialog>();
		}

		/// <summary>
		/// Wrapper interface for ISoundNote which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISoundNote WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISoundNote resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISoundNote, Interfaces.IISoundNote>();
		}

		/// <summary>
		/// Wrapper interface for IButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIButton WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IButton, Interfaces.IIButton>();
		}

		/// <summary>
		/// Wrapper interface for IButtons which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIButtons WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IButtons resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IButtons, Interfaces.IIButtons>();
		}

		/// <summary>
		/// Wrapper interface for ICheckBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICheckBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICheckBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICheckBox, Interfaces.IICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for ICheckBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICheckBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICheckBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICheckBoxes, Interfaces.IICheckBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IOptionButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOptionButton WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOptionButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOptionButton, Interfaces.IIOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for IOptionButtons which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOptionButtons WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOptionButtons resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOptionButtons, Interfaces.IIOptionButtons>();
		}

		/// <summary>
		/// Wrapper interface for IEditBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIEditBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IEditBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IEditBox, Interfaces.IIEditBox>();
		}

		/// <summary>
		/// Wrapper interface for IEditBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIEditBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IEditBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IEditBoxes, Interfaces.IIEditBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IScrollBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIScrollBar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IScrollBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IScrollBar, Interfaces.IIScrollBar>();
		}

		/// <summary>
		/// Wrapper interface for IScrollBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIScrollBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IScrollBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IScrollBars, Interfaces.IIScrollBars>();
		}

		/// <summary>
		/// Wrapper interface for IListBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListBox, Interfaces.IIListBox>();
		}

		/// <summary>
		/// Wrapper interface for IListBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListBoxes, Interfaces.IIListBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IGroupBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGroupBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGroupBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGroupBox, Interfaces.IIGroupBox>();
		}

		/// <summary>
		/// Wrapper interface for IGroupBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGroupBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGroupBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGroupBoxes, Interfaces.IIGroupBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IDropDown which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDropDown WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDropDown resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDropDown, Interfaces.IIDropDown>();
		}

		/// <summary>
		/// Wrapper interface for IDropDowns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDropDowns WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDropDowns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDropDowns, Interfaces.IIDropDowns>();
		}

		/// <summary>
		/// Wrapper interface for ISpinner which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISpinner WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISpinner resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISpinner, Interfaces.IISpinner>();
		}

		/// <summary>
		/// Wrapper interface for ISpinners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISpinners WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISpinners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISpinners, Interfaces.IISpinners>();
		}

		/// <summary>
		/// Wrapper interface for IDialogFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDialogFrame WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDialogFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDialogFrame, Interfaces.IIDialogFrame>();
		}

		/// <summary>
		/// Wrapper interface for ILabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILabel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILabel, Interfaces.IILabel>();
		}

		/// <summary>
		/// Wrapper interface for ILabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILabels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILabels, Interfaces.IILabels>();
		}

		/// <summary>
		/// Wrapper interface for IPanes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPanes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPanes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPanes, Interfaces.IIPanes>();
		}

		/// <summary>
		/// Wrapper interface for IPane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPane WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPane, Interfaces.IIPane>();
		}

		/// <summary>
		/// Wrapper interface for IScenarios which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIScenarios WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IScenarios resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IScenarios, Interfaces.IIScenarios>();
		}

		/// <summary>
		/// Wrapper interface for IScenario which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIScenario WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IScenario resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IScenario, Interfaces.IIScenario>();
		}

		/// <summary>
		/// Wrapper interface for IGroupObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGroupObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGroupObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGroupObject, Interfaces.IIGroupObject>();
		}

		/// <summary>
		/// Wrapper interface for IGroupObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGroupObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGroupObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGroupObjects, Interfaces.IIGroupObjects>();
		}

		/// <summary>
		/// Wrapper interface for ILine which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILine WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILine resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILine, Interfaces.IILine>();
		}

		/// <summary>
		/// Wrapper interface for ILines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILines, Interfaces.IILines>();
		}

		/// <summary>
		/// Wrapper interface for IRectangle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRectangle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRectangle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRectangle, Interfaces.IIRectangle>();
		}

		/// <summary>
		/// Wrapper interface for IRectangles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRectangles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRectangles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRectangles, Interfaces.IIRectangles>();
		}

		/// <summary>
		/// Wrapper interface for IOval which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOval WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOval resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOval, Interfaces.IIOval>();
		}

		/// <summary>
		/// Wrapper interface for IOvals which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOvals WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOvals resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOvals, Interfaces.IIOvals>();
		}

		/// <summary>
		/// Wrapper interface for IArc which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIArc WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IArc resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IArc, Interfaces.IIArc>();
		}

		/// <summary>
		/// Wrapper interface for IArcs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIArcs WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IArcs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IArcs, Interfaces.IIArcs>();
		}

		/// <summary>
		/// Wrapper interface for IOLEObjectEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOLEObjectEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOLEObjectEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOLEObjectEvents, Interfaces.IIOLEObjectEvents>();
		}

		/// <summary>
		/// Wrapper interface for _IOLEObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IOLEObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel._IOLEObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._IOLEObject, Interfaces.I_IOLEObject>();
		}

		/// <summary>
		/// Wrapper interface for IOLEObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOLEObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOLEObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOLEObjects, Interfaces.IIOLEObjects>();
		}

		/// <summary>
		/// Wrapper interface for ITextBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITextBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITextBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITextBox, Interfaces.IITextBox>();
		}

		/// <summary>
		/// Wrapper interface for ITextBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITextBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITextBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITextBoxes, Interfaces.IITextBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IPicture which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPicture WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPicture resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPicture, Interfaces.IIPicture>();
		}

		/// <summary>
		/// Wrapper interface for IPictures which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPictures WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPictures resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPictures, Interfaces.IIPictures>();
		}

		/// <summary>
		/// Wrapper interface for IDrawing which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDrawing WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDrawing resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDrawing, Interfaces.IIDrawing>();
		}

		/// <summary>
		/// Wrapper interface for IDrawings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDrawings WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDrawings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDrawings, Interfaces.IIDrawings>();
		}

		/// <summary>
		/// Wrapper interface for IRoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRoutingSlip WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRoutingSlip resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRoutingSlip, Interfaces.IIRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for IOutline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOutline WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOutline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOutline, Interfaces.IIOutline>();
		}

		/// <summary>
		/// Wrapper interface for IModule which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIModule WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IModule resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IModule, Interfaces.IIModule>();
		}

		/// <summary>
		/// Wrapper interface for IModules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIModules WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IModules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IModules, Interfaces.IIModules>();
		}

		/// <summary>
		/// Wrapper interface for IDialogSheet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDialogSheet WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDialogSheet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDialogSheet, Interfaces.IIDialogSheet>();
		}

		/// <summary>
		/// Wrapper interface for IDialogSheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDialogSheets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDialogSheets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDialogSheets, Interfaces.IIDialogSheets>();
		}

		/// <summary>
		/// Wrapper interface for IWorksheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWorksheets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWorksheets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWorksheets, Interfaces.IIWorksheets>();
		}

		/// <summary>
		/// Wrapper interface for IPageSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPageSetup WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPageSetup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPageSetup, Interfaces.IIPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for INames which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IINames WithComCleanupProxy(this Microsoft.Office.Interop.Excel.INames resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.INames, Interfaces.IINames>();
		}

		/// <summary>
		/// Wrapper interface for IName which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIName WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IName resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IName, Interfaces.IIName>();
		}

		/// <summary>
		/// Wrapper interface for IChartObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartObject, Interfaces.IIChartObject>();
		}

		/// <summary>
		/// Wrapper interface for IChartObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartObjects, Interfaces.IIChartObjects>();
		}

		/// <summary>
		/// Wrapper interface for IMailer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMailer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMailer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMailer, Interfaces.IIMailer>();
		}

		/// <summary>
		/// Wrapper interface for ICustomViews which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomViews WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICustomViews resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICustomViews, Interfaces.IICustomViews>();
		}

		/// <summary>
		/// Wrapper interface for ICustomView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICustomView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICustomView, Interfaces.IICustomView>();
		}

		/// <summary>
		/// Wrapper interface for IFormatConditions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFormatConditions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFormatConditions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFormatConditions, Interfaces.IIFormatConditions>();
		}

		/// <summary>
		/// Wrapper interface for IFormatCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFormatCondition WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFormatCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFormatCondition, Interfaces.IIFormatCondition>();
		}

		/// <summary>
		/// Wrapper interface for IComments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIComments WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IComments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IComments, Interfaces.IIComments>();
		}

		/// <summary>
		/// Wrapper interface for IComment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIComment WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IComment resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IComment, Interfaces.IIComment>();
		}

		/// <summary>
		/// Wrapper interface for IRefreshEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRefreshEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRefreshEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRefreshEvents, Interfaces.IIRefreshEvents>();
		}

		/// <summary>
		/// Wrapper interface for _IQueryTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_IQueryTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel._IQueryTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._IQueryTable, Interfaces.I_IQueryTable>();
		}

		/// <summary>
		/// Wrapper interface for IQueryTables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIQueryTables WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IQueryTables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IQueryTables, Interfaces.IIQueryTables>();
		}

		/// <summary>
		/// Wrapper interface for IParameter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIParameter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IParameter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IParameter, Interfaces.IIParameter>();
		}

		/// <summary>
		/// Wrapper interface for IParameters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIParameters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IParameters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IParameters, Interfaces.IIParameters>();
		}

		/// <summary>
		/// Wrapper interface for IODBCError which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIODBCError WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IODBCError resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IODBCError, Interfaces.IIODBCError>();
		}

		/// <summary>
		/// Wrapper interface for IODBCErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIODBCErrors WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IODBCErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IODBCErrors, Interfaces.IIODBCErrors>();
		}

		/// <summary>
		/// Wrapper interface for IValidation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIValidation WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IValidation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IValidation, Interfaces.IIValidation>();
		}

		/// <summary>
		/// Wrapper interface for IHyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIHyperlinks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IHyperlinks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IHyperlinks, Interfaces.IIHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for IHyperlink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIHyperlink WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IHyperlink resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IHyperlink, Interfaces.IIHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for IAutoFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAutoFilter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAutoFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAutoFilter, Interfaces.IIAutoFilter>();
		}

		/// <summary>
		/// Wrapper interface for IFilters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFilters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFilters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFilters, Interfaces.IIFilters>();
		}

		/// <summary>
		/// Wrapper interface for IFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFilter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFilter, Interfaces.IIFilter>();
		}

		/// <summary>
		/// Wrapper interface for IAutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAutoCorrect WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAutoCorrect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAutoCorrect, Interfaces.IIAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for IBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIBorder WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IBorder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IBorder, Interfaces.IIBorder>();
		}

		/// <summary>
		/// Wrapper interface for IInterior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIInterior WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IInterior resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IInterior, Interfaces.IIInterior>();
		}

		/// <summary>
		/// Wrapper interface for IChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartFillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartFillFormat, Interfaces.IIChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for IChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartColorFormat, Interfaces.IIChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for IAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAxis, Interfaces.IIAxis>();
		}

		/// <summary>
		/// Wrapper interface for IChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartTitle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartTitle, Interfaces.IIChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for IAxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAxisTitle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAxisTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAxisTitle, Interfaces.IIAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for IChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartGroup WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartGroup, Interfaces.IIChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for IChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartGroups WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartGroups, Interfaces.IIChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for IAxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAxes, Interfaces.IIAxes>();
		}

		/// <summary>
		/// Wrapper interface for IPoints which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPoints WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPoints resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPoints, Interfaces.IIPoints>();
		}

		/// <summary>
		/// Wrapper interface for IPoint which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPoint WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPoint resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPoint, Interfaces.IIPoint>();
		}

		/// <summary>
		/// Wrapper interface for ISeries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISeries WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISeries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISeries, Interfaces.IISeries>();
		}

		/// <summary>
		/// Wrapper interface for ISeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISeriesCollection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISeriesCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISeriesCollection, Interfaces.IISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for IDataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDataLabel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDataLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDataLabel, Interfaces.IIDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for IDataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDataLabels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDataLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDataLabels, Interfaces.IIDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for ILegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILegendEntry WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILegendEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILegendEntry, Interfaces.IILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for ILegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILegendEntries WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILegendEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILegendEntries, Interfaces.IILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for ILegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILegendKey WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILegendKey resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILegendKey, Interfaces.IILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for ITrendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITrendlines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITrendlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITrendlines, Interfaces.IITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for ITrendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITrendline WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITrendline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITrendline, Interfaces.IITrendline>();
		}

		/// <summary>
		/// Wrapper interface for ICorners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICorners WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICorners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICorners, Interfaces.IICorners>();
		}

		/// <summary>
		/// Wrapper interface for ISeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISeriesLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISeriesLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISeriesLines, Interfaces.IISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for IHiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIHiLoLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IHiLoLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IHiLoLines, Interfaces.IIHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for IGridlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGridlines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGridlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGridlines, Interfaces.IIGridlines>();
		}

		/// <summary>
		/// Wrapper interface for IDropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDropLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDropLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDropLines, Interfaces.IIDropLines>();
		}

		/// <summary>
		/// Wrapper interface for ILeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILeaderLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILeaderLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILeaderLines, Interfaces.IILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for IUpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIUpBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IUpBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IUpBars, Interfaces.IIUpBars>();
		}

		/// <summary>
		/// Wrapper interface for IDownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDownBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDownBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDownBars, Interfaces.IIDownBars>();
		}

		/// <summary>
		/// Wrapper interface for IFloor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFloor WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFloor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFloor, Interfaces.IIFloor>();
		}

		/// <summary>
		/// Wrapper interface for IWalls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWalls WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWalls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWalls, Interfaces.IIWalls>();
		}

		/// <summary>
		/// Wrapper interface for ITickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITickLabels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITickLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITickLabels, Interfaces.IITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for IPlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPlotArea WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPlotArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPlotArea, Interfaces.IIPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for IChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartArea WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartArea, Interfaces.IIChartArea>();
		}

		/// <summary>
		/// Wrapper interface for ILegend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILegend WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILegend resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILegend, Interfaces.IILegend>();
		}

		/// <summary>
		/// Wrapper interface for IErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIErrorBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IErrorBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IErrorBars, Interfaces.IIErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for IDataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDataTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDataTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDataTable, Interfaces.IIDataTable>();
		}

		/// <summary>
		/// Wrapper interface for IPhonetic which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPhonetic WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPhonetic resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPhonetic, Interfaces.IIPhonetic>();
		}

		/// <summary>
		/// Wrapper interface for IShape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIShape WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IShape resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IShape, Interfaces.IIShape>();
		}

		/// <summary>
		/// Wrapper interface for IShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIShapes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IShapes, Interfaces.IIShapes>();
		}

		/// <summary>
		/// Wrapper interface for IShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIShapeRange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IShapeRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IShapeRange, Interfaces.IIShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for IGroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGroupShapes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGroupShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGroupShapes, Interfaces.IIGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for ITextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITextFrame WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITextFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITextFrame, Interfaces.IITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for IConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConnectorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IConnectorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IConnectorFormat, Interfaces.IIConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for IFreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFreeformBuilder WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFreeformBuilder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFreeformBuilder, Interfaces.IIFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for IControlFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIControlFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IControlFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IControlFormat, Interfaces.IIControlFormat>();
		}

		/// <summary>
		/// Wrapper interface for IOLEFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOLEFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOLEFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOLEFormat, Interfaces.IIOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for ILinkFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILinkFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILinkFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILinkFormat, Interfaces.IILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for IPublishObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPublishObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPublishObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPublishObjects, Interfaces.IIPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for PublishObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPublishObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PublishObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PublishObject, Interfaces.IPublishObject>();
		}

		/// <summary>
		/// Wrapper interface for IOLEDBError which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOLEDBError WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOLEDBError resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOLEDBError, Interfaces.IIOLEDBError>();
		}

		/// <summary>
		/// Wrapper interface for IOLEDBErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOLEDBErrors WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOLEDBErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOLEDBErrors, Interfaces.IIOLEDBErrors>();
		}

		/// <summary>
		/// Wrapper interface for IPhonetics which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPhonetics WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPhonetics resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPhonetics, Interfaces.IIPhonetics>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDefaultWebOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DefaultWebOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DefaultWebOptions, Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.WebOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.WebOptions, Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLayout which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotLayout WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotLayout resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotLayout, Interfaces.IIPivotLayout>();
		}

		/// <summary>
		/// Wrapper interface for TreeviewControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITreeviewControl WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TreeviewControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TreeviewControl, Interfaces.ITreeviewControl>();
		}

		/// <summary>
		/// Wrapper interface for CubeField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICubeField WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CubeField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CubeField, Interfaces.ICubeField>();
		}

		/// <summary>
		/// Wrapper interface for CubeFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICubeFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CubeFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CubeFields, Interfaces.ICubeFields>();
		}

		/// <summary>
		/// Wrapper interface for IDisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDisplayUnitLabel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDisplayUnitLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDisplayUnitLabel, Interfaces.IIDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for ICellFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICellFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICellFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICellFormat, Interfaces.IICellFormat>();
		}

		/// <summary>
		/// Wrapper interface for IUsedObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIUsedObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IUsedObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IUsedObjects, Interfaces.IIUsedObjects>();
		}

		/// <summary>
		/// Wrapper interface for ICustomProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomProperties WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICustomProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICustomProperties, Interfaces.IICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for ICustomProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICustomProperty WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICustomProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICustomProperty, Interfaces.IICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedMembers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICalculatedMembers WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICalculatedMembers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICalculatedMembers, Interfaces.IICalculatedMembers>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedMember which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICalculatedMember WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICalculatedMember resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICalculatedMember, Interfaces.IICalculatedMember>();
		}

		/// <summary>
		/// Wrapper interface for IWatches which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWatches WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWatches resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWatches, Interfaces.IIWatches>();
		}

		/// <summary>
		/// Wrapper interface for IWatch which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWatch WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWatch resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWatch, Interfaces.IIWatch>();
		}

		/// <summary>
		/// Wrapper interface for IPivotCell which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotCell WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotCell resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotCell, Interfaces.IIPivotCell>();
		}

		/// <summary>
		/// Wrapper interface for IGraphic which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIGraphic WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IGraphic resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IGraphic, Interfaces.IIGraphic>();
		}

		/// <summary>
		/// Wrapper interface for IAutoRecover which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAutoRecover WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAutoRecover resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAutoRecover, Interfaces.IIAutoRecover>();
		}

		/// <summary>
		/// Wrapper interface for IErrorCheckingOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIErrorCheckingOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IErrorCheckingOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IErrorCheckingOptions, Interfaces.IIErrorCheckingOptions>();
		}

		/// <summary>
		/// Wrapper interface for IErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIErrors WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IErrors, Interfaces.IIErrors>();
		}

		/// <summary>
		/// Wrapper interface for IError which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIError WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IError resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IError, Interfaces.IIError>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTagAction WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTagAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTagAction, Interfaces.IISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTagActions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTagActions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTagActions, Interfaces.IISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTag which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTag WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTag resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTag, Interfaces.IISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTags which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTags WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTags resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTags, Interfaces.IISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTagRecognizer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTagRecognizer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTagRecognizer, Interfaces.IISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTagRecognizers WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTagRecognizers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTagRecognizers, Interfaces.IISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISmartTagOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISmartTagOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISmartTagOptions, Interfaces.IISmartTagOptions>();
		}

		/// <summary>
		/// Wrapper interface for ISpellingOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISpellingOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISpellingOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISpellingOptions, Interfaces.IISpellingOptions>();
		}

		/// <summary>
		/// Wrapper interface for ISpeech which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISpeech WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISpeech resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISpeech, Interfaces.IISpeech>();
		}

		/// <summary>
		/// Wrapper interface for IProtection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIProtection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IProtection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IProtection, Interfaces.IIProtection>();
		}

		/// <summary>
		/// Wrapper interface for IPivotItemList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotItemList WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotItemList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotItemList, Interfaces.IIPivotItemList>();
		}

		/// <summary>
		/// Wrapper interface for ITab which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITab WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITab resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITab, Interfaces.IITab>();
		}

		/// <summary>
		/// Wrapper interface for IAllowEditRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAllowEditRanges WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAllowEditRanges resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAllowEditRanges, Interfaces.IIAllowEditRanges>();
		}

		/// <summary>
		/// Wrapper interface for IAllowEditRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAllowEditRange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAllowEditRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAllowEditRange, Interfaces.IIAllowEditRange>();
		}

		/// <summary>
		/// Wrapper interface for IUserAccessList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIUserAccessList WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IUserAccessList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IUserAccessList, Interfaces.IIUserAccessList>();
		}

		/// <summary>
		/// Wrapper interface for IUserAccess which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIUserAccess WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IUserAccess resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IUserAccess, Interfaces.IIUserAccess>();
		}

		/// <summary>
		/// Wrapper interface for IRTD which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRTD WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRTD resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRTD, Interfaces.IIRTD>();
		}

		/// <summary>
		/// Wrapper interface for IDiagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDiagram WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDiagram resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDiagram, Interfaces.IIDiagram>();
		}

		/// <summary>
		/// Wrapper interface for IListObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListObjects, Interfaces.IIListObjects>();
		}

		/// <summary>
		/// Wrapper interface for IListObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListObject, Interfaces.IIListObject>();
		}

		/// <summary>
		/// Wrapper interface for IListColumns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListColumns WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListColumns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListColumns, Interfaces.IIListColumns>();
		}

		/// <summary>
		/// Wrapper interface for IListColumn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListColumn WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListColumn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListColumn, Interfaces.IIListColumn>();
		}

		/// <summary>
		/// Wrapper interface for IListRows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListRows WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListRows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListRows, Interfaces.IIListRows>();
		}

		/// <summary>
		/// Wrapper interface for IListRow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListRow WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListRow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListRow, Interfaces.IIListRow>();
		}

		/// <summary>
		/// Wrapper interface for IXmlNamespace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlNamespace WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlNamespace resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlNamespace, Interfaces.IIXmlNamespace>();
		}

		/// <summary>
		/// Wrapper interface for IXmlNamespaces which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlNamespaces WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlNamespaces resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlNamespaces, Interfaces.IIXmlNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for IXmlDataBinding which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlDataBinding WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlDataBinding resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlDataBinding, Interfaces.IIXmlDataBinding>();
		}

		/// <summary>
		/// Wrapper interface for IXmlSchema which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlSchema WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlSchema resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlSchema, Interfaces.IIXmlSchema>();
		}

		/// <summary>
		/// Wrapper interface for IXmlSchemas which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlSchemas WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlSchemas resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlSchemas, Interfaces.IIXmlSchemas>();
		}

		/// <summary>
		/// Wrapper interface for IXmlMap which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlMap WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlMap resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlMap, Interfaces.IIXmlMap>();
		}

		/// <summary>
		/// Wrapper interface for IXmlMaps which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXmlMaps WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXmlMaps resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXmlMaps, Interfaces.IIXmlMaps>();
		}

		/// <summary>
		/// Wrapper interface for IListDataFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIListDataFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IListDataFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IListDataFormat, Interfaces.IIListDataFormat>();
		}

		/// <summary>
		/// Wrapper interface for IXPath which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIXPath WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IXPath resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IXPath, Interfaces.IIXPath>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLineCells which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotLineCells WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotLineCells resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotLineCells, Interfaces.IIPivotLineCells>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLine which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotLine WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotLine resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotLine, Interfaces.IIPivotLine>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotLines, Interfaces.IIPivotLines>();
		}

		/// <summary>
		/// Wrapper interface for IPivotAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotAxis, Interfaces.IIPivotAxis>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotFilter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotFilter, Interfaces.IIPivotFilter>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFilters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotFilters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotFilters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotFilters, Interfaces.IIPivotFilters>();
		}

		/// <summary>
		/// Wrapper interface for IWorkbookConnection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWorkbookConnection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWorkbookConnection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWorkbookConnection, Interfaces.IIWorkbookConnection>();
		}

		/// <summary>
		/// Wrapper interface for IConnections which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConnections WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IConnections resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IConnections, Interfaces.IIConnections>();
		}

		/// <summary>
		/// Wrapper interface for IWorksheetView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIWorksheetView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IWorksheetView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IWorksheetView, Interfaces.IIWorksheetView>();
		}

		/// <summary>
		/// Wrapper interface for IChartView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartView, Interfaces.IIChartView>();
		}

		/// <summary>
		/// Wrapper interface for IModuleView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIModuleView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IModuleView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IModuleView, Interfaces.IIModuleView>();
		}

		/// <summary>
		/// Wrapper interface for IDialogSheetView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDialogSheetView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDialogSheetView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDialogSheetView, Interfaces.IIDialogSheetView>();
		}

		/// <summary>
		/// Wrapper interface for ISheetViews which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISheetViews WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISheetViews resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISheetViews, Interfaces.IISheetViews>();
		}

		/// <summary>
		/// Wrapper interface for IOLEDBConnection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIOLEDBConnection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IOLEDBConnection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IOLEDBConnection, Interfaces.IIOLEDBConnection>();
		}

		/// <summary>
		/// Wrapper interface for IODBCConnection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIODBCConnection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IODBCConnection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IODBCConnection, Interfaces.IIODBCConnection>();
		}

		/// <summary>
		/// Wrapper interface for IAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAction WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAction, Interfaces.IIAction>();
		}

		/// <summary>
		/// Wrapper interface for IActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIActions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IActions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IActions, Interfaces.IIActions>();
		}

		/// <summary>
		/// Wrapper interface for IFormatColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFormatColor WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFormatColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFormatColor, Interfaces.IIFormatColor>();
		}

		/// <summary>
		/// Wrapper interface for IConditionValue which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIConditionValue WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IConditionValue resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IConditionValue, Interfaces.IIConditionValue>();
		}

		/// <summary>
		/// Wrapper interface for IColorScale which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIColorScale WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IColorScale resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IColorScale, Interfaces.IIColorScale>();
		}

		/// <summary>
		/// Wrapper interface for IColorScaleCriteria which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIColorScaleCriteria WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IColorScaleCriteria resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IColorScaleCriteria, Interfaces.IIColorScaleCriteria>();
		}

		/// <summary>
		/// Wrapper interface for IColorScaleCriterion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIColorScaleCriterion WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IColorScaleCriterion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IColorScaleCriterion, Interfaces.IIColorScaleCriterion>();
		}

		/// <summary>
		/// Wrapper interface for IDatabar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDatabar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDatabar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDatabar, Interfaces.IIDatabar>();
		}

		/// <summary>
		/// Wrapper interface for IIconSetCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIIconSetCondition WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IIconSetCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IIconSetCondition, Interfaces.IIIconSetCondition>();
		}

		/// <summary>
		/// Wrapper interface for IIconCriteria which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIIconCriteria WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IIconCriteria resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IIconCriteria, Interfaces.IIIconCriteria>();
		}

		/// <summary>
		/// Wrapper interface for IIconCriterion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIIconCriterion WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IIconCriterion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IIconCriterion, Interfaces.IIIconCriterion>();
		}

		/// <summary>
		/// Wrapper interface for IIcon which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIIcon WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IIcon resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IIcon, Interfaces.IIIcon>();
		}

		/// <summary>
		/// Wrapper interface for IIconSet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIIconSet WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IIconSet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IIconSet, Interfaces.IIIconSet>();
		}

		/// <summary>
		/// Wrapper interface for IIconSets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIIconSets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IIconSets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IIconSets, Interfaces.IIIconSets>();
		}

		/// <summary>
		/// Wrapper interface for ITop10 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITop10 WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITop10 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITop10, Interfaces.IITop10>();
		}

		/// <summary>
		/// Wrapper interface for IAboveAverage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAboveAverage WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAboveAverage resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAboveAverage, Interfaces.IIAboveAverage>();
		}

		/// <summary>
		/// Wrapper interface for IUniqueValues which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIUniqueValues WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IUniqueValues resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IUniqueValues, Interfaces.IIUniqueValues>();
		}

		/// <summary>
		/// Wrapper interface for IRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRanges WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRanges resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRanges, Interfaces.IIRanges>();
		}

		/// <summary>
		/// Wrapper interface for IHeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIHeaderFooter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IHeaderFooter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IHeaderFooter, Interfaces.IIHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for IPage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPage WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPage resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPage, Interfaces.IIPage>();
		}

		/// <summary>
		/// Wrapper interface for IPages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPages WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPages resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPages, Interfaces.IIPages>();
		}

		/// <summary>
		/// Wrapper interface for IServerViewableItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIServerViewableItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IServerViewableItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IServerViewableItems, Interfaces.IIServerViewableItems>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyleElement which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITableStyleElement WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITableStyleElement resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITableStyleElement, Interfaces.IITableStyleElement>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyleElements which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITableStyleElements WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITableStyleElements resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITableStyleElements, Interfaces.IITableStyleElements>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITableStyle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITableStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITableStyle, Interfaces.IITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IITableStyles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ITableStyles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ITableStyles, Interfaces.IITableStyles>();
		}

		/// <summary>
		/// Wrapper interface for ISortField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISortField WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISortField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISortField, Interfaces.IISortField>();
		}

		/// <summary>
		/// Wrapper interface for ISortFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISortFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISortFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISortFields, Interfaces.IISortFields>();
		}

		/// <summary>
		/// Wrapper interface for ISort which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISort WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISort resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISort, Interfaces.IISort>();
		}

		/// <summary>
		/// Wrapper interface for IResearch which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIResearch WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IResearch resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IResearch, Interfaces.IIResearch>();
		}

		/// <summary>
		/// Wrapper interface for IColorStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIColorStop WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IColorStop resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IColorStop, Interfaces.IIColorStop>();
		}

		/// <summary>
		/// Wrapper interface for IColorStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIColorStops WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IColorStops resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IColorStops, Interfaces.IIColorStops>();
		}

		/// <summary>
		/// Wrapper interface for ILinearGradient which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IILinearGradient WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ILinearGradient resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ILinearGradient, Interfaces.IILinearGradient>();
		}

		/// <summary>
		/// Wrapper interface for IRectangularGradient which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIRectangularGradient WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IRectangularGradient resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IRectangularGradient, Interfaces.IIRectangularGradient>();
		}

		/// <summary>
		/// Wrapper interface for IMultiThreadedCalculation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIMultiThreadedCalculation WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IMultiThreadedCalculation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IMultiThreadedCalculation, Interfaces.IIMultiThreadedCalculation>();
		}

		/// <summary>
		/// Wrapper interface for IChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIChartFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IChartFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IChartFormat, Interfaces.IIChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for IFileExportConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFileExportConverter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFileExportConverter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFileExportConverter, Interfaces.IIFileExportConverter>();
		}

		/// <summary>
		/// Wrapper interface for IFileExportConverters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIFileExportConverters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IFileExportConverters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IFileExportConverters, Interfaces.IIFileExportConverters>();
		}

		/// <summary>
		/// Wrapper interface for IAddIns2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIAddIns2 WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IAddIns2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IAddIns2, Interfaces.IIAddIns2>();
		}

		/// <summary>
		/// Wrapper interface for ISparklineGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparklineGroups WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparklineGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparklineGroups, Interfaces.IISparklineGroups>();
		}

		/// <summary>
		/// Wrapper interface for ISparklineGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparklineGroup WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparklineGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparklineGroup, Interfaces.IISparklineGroup>();
		}

		/// <summary>
		/// Wrapper interface for ISparkPoints which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparkPoints WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparkPoints resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparkPoints, Interfaces.IISparkPoints>();
		}

		/// <summary>
		/// Wrapper interface for ISparkline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparkline WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparkline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparkline, Interfaces.IISparkline>();
		}

		/// <summary>
		/// Wrapper interface for ISparkAxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparkAxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparkAxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparkAxes, Interfaces.IISparkAxes>();
		}

		/// <summary>
		/// Wrapper interface for ISparkHorizontalAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparkHorizontalAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparkHorizontalAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparkHorizontalAxis, Interfaces.IISparkHorizontalAxis>();
		}

		/// <summary>
		/// Wrapper interface for ISparkVerticalAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparkVerticalAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparkVerticalAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparkVerticalAxis, Interfaces.IISparkVerticalAxis>();
		}

		/// <summary>
		/// Wrapper interface for ISparkColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISparkColor WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISparkColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISparkColor, Interfaces.IISparkColor>();
		}

		/// <summary>
		/// Wrapper interface for IDataBarBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDataBarBorder WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDataBarBorder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDataBarBorder, Interfaces.IIDataBarBorder>();
		}

		/// <summary>
		/// Wrapper interface for INegativeBarFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IINegativeBarFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.INegativeBarFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.INegativeBarFormat, Interfaces.IINegativeBarFormat>();
		}

		/// <summary>
		/// Wrapper interface for IValueChange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIValueChange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IValueChange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IValueChange, Interfaces.IIValueChange>();
		}

		/// <summary>
		/// Wrapper interface for IPivotTableChangeList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIPivotTableChangeList WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IPivotTableChangeList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IPivotTableChangeList, Interfaces.IIPivotTableChangeList>();
		}

		/// <summary>
		/// Wrapper interface for IDisplayFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDisplayFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDisplayFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDisplayFormat, Interfaces.IIDisplayFormat>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCaches which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerCaches WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerCaches resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerCaches, Interfaces.IISlicerCaches>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCache which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerCache WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerCache resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerCache, Interfaces.IISlicerCache>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCacheLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerCacheLevels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerCacheLevels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerCacheLevels, Interfaces.IISlicerCacheLevels>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCacheLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerCacheLevel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerCacheLevel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerCacheLevel, Interfaces.IISlicerCacheLevel>();
		}

		/// <summary>
		/// Wrapper interface for ISlicers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicers WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicers, Interfaces.IISlicers>();
		}

		/// <summary>
		/// Wrapper interface for ISlicer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicer, Interfaces.IISlicer>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerItem WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerItem, Interfaces.IISlicerItem>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerItems, Interfaces.IISlicerItems>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerPivotTables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IISlicerPivotTables WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ISlicerPivotTables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ISlicerPivotTables, Interfaces.IISlicerPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for IProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIProtectedViewWindows WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IProtectedViewWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IProtectedViewWindows, Interfaces.IIProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for IProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIProtectedViewWindow WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IProtectedViewWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IProtectedViewWindow, Interfaces.IIProtectedViewWindow>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFont WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Font resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Font, Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for Window which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWindow WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Window resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Window, Interfaces.IWindow>();
		}

		/// <summary>
		/// Wrapper interface for Windows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWindows WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Windows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Windows, Interfaces.IWindows>();
		}

		/// <summary>
		/// Wrapper interface for AppEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAppEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AppEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AppEvents, Interfaces.IAppEvents>();
		}

		/// <summary>
		/// Wrapper interface for WorksheetFunction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorksheetFunction WithComCleanupProxy(this Microsoft.Office.Interop.Excel.WorksheetFunction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.WorksheetFunction, Interfaces.IWorksheetFunction>();
		}

		/// <summary>
		/// Wrapper interface for Range which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Range resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Range, Interfaces.IRange>();
		}

		/// <summary>
		/// Wrapper interface for ChartEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartEvents, Interfaces.IChartEvents>();
		}

		/// <summary>
		/// Wrapper interface for VPageBreak which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVPageBreak WithComCleanupProxy(this Microsoft.Office.Interop.Excel.VPageBreak resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.VPageBreak, Interfaces.IVPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for HPageBreak which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHPageBreak WithComCleanupProxy(this Microsoft.Office.Interop.Excel.HPageBreak resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.HPageBreak, Interfaces.IHPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for HPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHPageBreaks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.HPageBreaks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.HPageBreaks, Interfaces.IHPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for VPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IVPageBreaks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.VPageBreaks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.VPageBreaks, Interfaces.IVPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for RecentFile which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecentFile WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RecentFile resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RecentFile, Interfaces.IRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for RecentFiles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRecentFiles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RecentFiles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RecentFiles, Interfaces.IRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for DocEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DocEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DocEvents, Interfaces.IDocEvents>();
		}

		/// <summary>
		/// Wrapper interface for Style which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Style resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Style, Interfaces.IStyle>();
		}

		/// <summary>
		/// Wrapper interface for Styles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IStyles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Styles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Styles, Interfaces.IStyles>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorders WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Borders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Borders, Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIn WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AddIn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AddIn, Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIns WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AddIns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AddIns, Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for Toolbar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IToolbar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Toolbar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Toolbar, Interfaces.IToolbar>();
		}

		/// <summary>
		/// Wrapper interface for Toolbars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IToolbars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Toolbars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Toolbars, Interfaces.IToolbars>();
		}

		/// <summary>
		/// Wrapper interface for ToolbarButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IToolbarButton WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ToolbarButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ToolbarButton, Interfaces.IToolbarButton>();
		}

		/// <summary>
		/// Wrapper interface for ToolbarButtons which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IToolbarButtons WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ToolbarButtons resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ToolbarButtons, Interfaces.IToolbarButtons>();
		}

		/// <summary>
		/// Wrapper interface for Areas which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAreas WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Areas resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Areas, Interfaces.IAreas>();
		}

		/// <summary>
		/// Wrapper interface for WorkbookEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkbookEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.WorkbookEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.WorkbookEvents, Interfaces.IWorkbookEvents>();
		}

		/// <summary>
		/// Wrapper interface for MenuBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMenuBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.MenuBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.MenuBars, Interfaces.IMenuBars>();
		}

		/// <summary>
		/// Wrapper interface for MenuBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMenuBar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.MenuBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.MenuBar, Interfaces.IMenuBar>();
		}

		/// <summary>
		/// Wrapper interface for Menus which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMenus WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Menus resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Menus, Interfaces.IMenus>();
		}

		/// <summary>
		/// Wrapper interface for Menu which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMenu WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Menu resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Menu, Interfaces.IMenu>();
		}

		/// <summary>
		/// Wrapper interface for MenuItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMenuItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.MenuItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.MenuItems, Interfaces.IMenuItems>();
		}

		/// <summary>
		/// Wrapper interface for MenuItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMenuItem WithComCleanupProxy(this Microsoft.Office.Interop.Excel.MenuItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.MenuItem, Interfaces.IMenuItem>();
		}

		/// <summary>
		/// Wrapper interface for Charts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICharts WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Charts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Charts, Interfaces.ICharts>();
		}

		/// <summary>
		/// Wrapper interface for DrawingObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDrawingObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DrawingObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DrawingObjects, Interfaces.IDrawingObjects>();
		}

		/// <summary>
		/// Wrapper interface for PivotCache which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotCache WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotCache resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotCache, Interfaces.IPivotCache>();
		}

		/// <summary>
		/// Wrapper interface for PivotCaches which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotCaches WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotCaches resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotCaches, Interfaces.IPivotCaches>();
		}

		/// <summary>
		/// Wrapper interface for PivotFormula which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotFormula WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotFormula resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotFormula, Interfaces.IPivotFormula>();
		}

		/// <summary>
		/// Wrapper interface for PivotFormulas which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotFormulas WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotFormulas resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotFormulas, Interfaces.IPivotFormulas>();
		}

		/// <summary>
		/// Wrapper interface for PivotTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotTable, Interfaces.IPivotTable>();
		}

		/// <summary>
		/// Wrapper interface for PivotTables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotTables WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotTables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotTables, Interfaces.IPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for PivotField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotField WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotField, Interfaces.IPivotField>();
		}

		/// <summary>
		/// Wrapper interface for PivotFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotFields, Interfaces.IPivotFields>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalculatedFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CalculatedFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CalculatedFields, Interfaces.ICalculatedFields>();
		}

		/// <summary>
		/// Wrapper interface for PivotItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotItem WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotItem, Interfaces.IPivotItem>();
		}

		/// <summary>
		/// Wrapper interface for PivotItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotItems, Interfaces.IPivotItems>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalculatedItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CalculatedItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CalculatedItems, Interfaces.ICalculatedItems>();
		}

		/// <summary>
		/// Wrapper interface for Characters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICharacters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Characters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Characters, Interfaces.ICharacters>();
		}

		/// <summary>
		/// Wrapper interface for Dialogs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogs WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Dialogs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Dialogs, Interfaces.IDialogs>();
		}

		/// <summary>
		/// Wrapper interface for Dialog which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialog WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Dialog resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Dialog, Interfaces.IDialog>();
		}

		/// <summary>
		/// Wrapper interface for SoundNote which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoundNote WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SoundNote resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SoundNote, Interfaces.ISoundNote>();
		}

		/// <summary>
		/// Wrapper interface for Button which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IButton WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Button resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Button, Interfaces.IButton>();
		}

		/// <summary>
		/// Wrapper interface for Buttons which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IButtons WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Buttons resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Buttons, Interfaces.IButtons>();
		}

		/// <summary>
		/// Wrapper interface for CheckBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICheckBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CheckBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CheckBox, Interfaces.ICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for CheckBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICheckBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CheckBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CheckBoxes, Interfaces.ICheckBoxes>();
		}

		/// <summary>
		/// Wrapper interface for OptionButton which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOptionButton WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OptionButton resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OptionButton, Interfaces.IOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for OptionButtons which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOptionButtons WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OptionButtons resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OptionButtons, Interfaces.IOptionButtons>();
		}

		/// <summary>
		/// Wrapper interface for EditBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEditBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.EditBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.EditBox, Interfaces.IEditBox>();
		}

		/// <summary>
		/// Wrapper interface for EditBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEditBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.EditBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.EditBoxes, Interfaces.IEditBoxes>();
		}

		/// <summary>
		/// Wrapper interface for ScrollBar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScrollBar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ScrollBar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ScrollBar, Interfaces.IScrollBar>();
		}

		/// <summary>
		/// Wrapper interface for ScrollBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScrollBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ScrollBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ScrollBars, Interfaces.IScrollBars>();
		}

		/// <summary>
		/// Wrapper interface for ListBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListBox, Interfaces.IListBox>();
		}

		/// <summary>
		/// Wrapper interface for ListBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListBoxes, Interfaces.IListBoxes>();
		}

		/// <summary>
		/// Wrapper interface for GroupBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.GroupBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.GroupBox, Interfaces.IGroupBox>();
		}

		/// <summary>
		/// Wrapper interface for GroupBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.GroupBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.GroupBoxes, Interfaces.IGroupBoxes>();
		}

		/// <summary>
		/// Wrapper interface for DropDown which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropDown WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DropDown resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DropDown, Interfaces.IDropDown>();
		}

		/// <summary>
		/// Wrapper interface for DropDowns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropDowns WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DropDowns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DropDowns, Interfaces.IDropDowns>();
		}

		/// <summary>
		/// Wrapper interface for Spinner which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpinner WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Spinner resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Spinner, Interfaces.ISpinner>();
		}

		/// <summary>
		/// Wrapper interface for Spinners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpinners WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Spinners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Spinners, Interfaces.ISpinners>();
		}

		/// <summary>
		/// Wrapper interface for DialogFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogFrame WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DialogFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DialogFrame, Interfaces.IDialogFrame>();
		}

		/// <summary>
		/// Wrapper interface for Label which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILabel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Label resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Label, Interfaces.ILabel>();
		}

		/// <summary>
		/// Wrapper interface for Labels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILabels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Labels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Labels, Interfaces.ILabels>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPanes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Panes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Panes, Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPane WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Pane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Pane, Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for Scenarios which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScenarios WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Scenarios resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Scenarios, Interfaces.IScenarios>();
		}

		/// <summary>
		/// Wrapper interface for Scenario which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScenario WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Scenario resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Scenario, Interfaces.IScenario>();
		}

		/// <summary>
		/// Wrapper interface for GroupObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.GroupObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.GroupObject, Interfaces.IGroupObject>();
		}

		/// <summary>
		/// Wrapper interface for GroupObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.GroupObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.GroupObjects, Interfaces.IGroupObjects>();
		}

		/// <summary>
		/// Wrapper interface for Line which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILine WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Line resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Line, Interfaces.ILine>();
		}

		/// <summary>
		/// Wrapper interface for Lines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Lines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Lines, Interfaces.ILines>();
		}

		/// <summary>
		/// Wrapper interface for Rectangle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Rectangle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Rectangle, Interfaces.IRectangle>();
		}

		/// <summary>
		/// Wrapper interface for Rectangles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Rectangles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Rectangles, Interfaces.IRectangles>();
		}

		/// <summary>
		/// Wrapper interface for Oval which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOval WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Oval resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Oval, Interfaces.IOval>();
		}

		/// <summary>
		/// Wrapper interface for Ovals which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOvals WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Ovals resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Ovals, Interfaces.IOvals>();
		}

		/// <summary>
		/// Wrapper interface for Arc which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IArc WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Arc resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Arc, Interfaces.IArc>();
		}

		/// <summary>
		/// Wrapper interface for Arcs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IArcs WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Arcs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Arcs, Interfaces.IArcs>();
		}

		/// <summary>
		/// Wrapper interface for OLEObjectEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEObjectEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEObjectEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEObjectEvents, Interfaces.IOLEObjectEvents>();
		}

		/// <summary>
		/// Wrapper interface for _OLEObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_OLEObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel._OLEObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._OLEObject, Interfaces.I_OLEObject>();
		}

		/// <summary>
		/// Wrapper interface for OLEObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEObjects, Interfaces.IOLEObjects>();
		}

		/// <summary>
		/// Wrapper interface for TextBox which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextBox WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TextBox resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TextBox, Interfaces.ITextBox>();
		}

		/// <summary>
		/// Wrapper interface for TextBoxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextBoxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TextBoxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TextBoxes, Interfaces.ITextBoxes>();
		}

		/// <summary>
		/// Wrapper interface for Picture which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPicture WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Picture resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Picture, Interfaces.IPicture>();
		}

		/// <summary>
		/// Wrapper interface for Pictures which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictures WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Pictures resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Pictures, Interfaces.IPictures>();
		}

		/// <summary>
		/// Wrapper interface for Drawing which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDrawing WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Drawing resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Drawing, Interfaces.IDrawing>();
		}

		/// <summary>
		/// Wrapper interface for Drawings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDrawings WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Drawings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Drawings, Interfaces.IDrawings>();
		}

		/// <summary>
		/// Wrapper interface for RoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRoutingSlip WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RoutingSlip resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RoutingSlip, Interfaces.IRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for Outline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOutline WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Outline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Outline, Interfaces.IOutline>();
		}

		/// <summary>
		/// Wrapper interface for Module which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IModule WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Module resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Module, Interfaces.IModule>();
		}

		/// <summary>
		/// Wrapper interface for Modules which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IModules WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Modules resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Modules, Interfaces.IModules>();
		}

		/// <summary>
		/// Wrapper interface for DialogSheet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogSheet WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DialogSheet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DialogSheet, Interfaces.IDialogSheet>();
		}

		/// <summary>
		/// Wrapper interface for DialogSheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogSheets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DialogSheets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DialogSheets, Interfaces.IDialogSheets>();
		}

		/// <summary>
		/// Wrapper interface for Worksheets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorksheets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Worksheets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Worksheets, Interfaces.IWorksheets>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageSetup WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PageSetup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PageSetup, Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for Names which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INames WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Names resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Names, Interfaces.INames>();
		}

		/// <summary>
		/// Wrapper interface for Name which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IName WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Name resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Name, Interfaces.IName>();
		}

		/// <summary>
		/// Wrapper interface for ChartObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartObject, Interfaces.IChartObject>();
		}

		/// <summary>
		/// Wrapper interface for ChartObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartObjects, Interfaces.IChartObjects>();
		}

		/// <summary>
		/// Wrapper interface for Mailer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMailer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Mailer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Mailer, Interfaces.IMailer>();
		}

		/// <summary>
		/// Wrapper interface for CustomViews which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomViews WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CustomViews resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CustomViews, Interfaces.ICustomViews>();
		}

		/// <summary>
		/// Wrapper interface for CustomView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CustomView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CustomView, Interfaces.ICustomView>();
		}

		/// <summary>
		/// Wrapper interface for FormatConditions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormatConditions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FormatConditions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FormatConditions, Interfaces.IFormatConditions>();
		}

		/// <summary>
		/// Wrapper interface for FormatCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormatCondition WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FormatCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FormatCondition, Interfaces.IFormatCondition>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComments WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Comments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Comments, Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComment WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Comment resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Comment, Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for RefreshEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRefreshEvents WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RefreshEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RefreshEvents, Interfaces.IRefreshEvents>();
		}

		/// <summary>
		/// Wrapper interface for _QueryTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_QueryTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel._QueryTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel._QueryTable, Interfaces.I_QueryTable>();
		}

		/// <summary>
		/// Wrapper interface for QueryTables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IQueryTables WithComCleanupProxy(this Microsoft.Office.Interop.Excel.QueryTables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.QueryTables, Interfaces.IQueryTables>();
		}

		/// <summary>
		/// Wrapper interface for Parameter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParameter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Parameter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Parameter, Interfaces.IParameter>();
		}

		/// <summary>
		/// Wrapper interface for Parameters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParameters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Parameters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Parameters, Interfaces.IParameters>();
		}

		/// <summary>
		/// Wrapper interface for ODBCError which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODBCError WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ODBCError resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ODBCError, Interfaces.IODBCError>();
		}

		/// <summary>
		/// Wrapper interface for ODBCErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODBCErrors WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ODBCErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ODBCErrors, Interfaces.IODBCErrors>();
		}

		/// <summary>
		/// Wrapper interface for Validation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IValidation WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Validation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Validation, Interfaces.IValidation>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlinks WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Hyperlinks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Hyperlinks, Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlink WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Hyperlink resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Hyperlink, Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for AutoFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoFilter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AutoFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AutoFilter, Interfaces.IAutoFilter>();
		}

		/// <summary>
		/// Wrapper interface for Filters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFilters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Filters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Filters, Interfaces.IFilters>();
		}

		/// <summary>
		/// Wrapper interface for Filter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFilter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Filter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Filter, Interfaces.IFilter>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrect WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AutoCorrect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AutoCorrect, Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for Border which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorder WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Border resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Border, Interfaces.IBorder>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInterior WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Interior resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Interior, Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartFillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartFillFormat, Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartColorFormat, Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Axis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Axis, Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartTitle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartTitle, Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxisTitle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AxisTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AxisTitle, Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroup WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartGroup, Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroups WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartGroups, Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Axes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Axes, Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoints WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Points resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Points, Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoint WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Point resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Point, Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeries WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Series resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Series, Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesCollection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SeriesCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SeriesCollection, Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DataLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DataLabel, Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DataLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DataLabels, Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntry WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LegendEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LegendEntry, Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntries WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LegendEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LegendEntries, Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendKey WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LegendKey resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LegendKey, Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendlines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Trendlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Trendlines, Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendline WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Trendline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Trendline, Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICorners WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Corners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Corners, Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SeriesLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SeriesLines, Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHiLoLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.HiLoLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.HiLoLines, Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGridlines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Gridlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Gridlines, Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DropLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DropLines, Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILeaderLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LeaderLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LeaderLines, Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUpBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.UpBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.UpBars, Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDownBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DownBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DownBars, Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFloor WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Floor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Floor, Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWalls WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Walls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Walls, Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITickLabels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TickLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TickLabels, Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlotArea WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PlotArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PlotArea, Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartArea WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartArea, Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegend WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Legend resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Legend, Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrorBars WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ErrorBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ErrorBars, Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DataTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DataTable, Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for Phonetic which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPhonetic WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Phonetic resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Phonetic, Interfaces.IPhonetic>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShape WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Shape resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Shape, Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Shapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Shapes, Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeRange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ShapeRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ShapeRange, Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupShapes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.GroupShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.GroupShapes, Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TextFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TextFrame, Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnectorFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ConnectorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ConnectorFormat, Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFreeformBuilder WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FreeformBuilder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FreeformBuilder, Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for ControlFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IControlFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ControlFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ControlFormat, Interfaces.IControlFormat>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEFormat, Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinkFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LinkFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LinkFormat, Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for PublishObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPublishObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PublishObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PublishObjects, Interfaces.IPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for OLEDBError which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEDBError WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEDBError resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEDBError, Interfaces.IOLEDBError>();
		}

		/// <summary>
		/// Wrapper interface for OLEDBErrors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEDBErrors WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEDBErrors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEDBErrors, Interfaces.IOLEDBErrors>();
		}

		/// <summary>
		/// Wrapper interface for Phonetics which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPhonetics WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Phonetics resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Phonetics, Interfaces.IPhonetics>();
		}

		/// <summary>
		/// Wrapper interface for PivotLayout which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotLayout WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotLayout resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotLayout, Interfaces.IPivotLayout>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDisplayUnitLabel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DisplayUnitLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DisplayUnitLabel, Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for CellFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICellFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CellFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CellFormat, Interfaces.ICellFormat>();
		}

		/// <summary>
		/// Wrapper interface for UsedObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUsedObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.UsedObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.UsedObjects, Interfaces.IUsedObjects>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomProperties WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CustomProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CustomProperties, Interfaces.ICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperty which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomProperty WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CustomProperty resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CustomProperty, Interfaces.ICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedMembers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalculatedMembers WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CalculatedMembers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CalculatedMembers, Interfaces.ICalculatedMembers>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedMember which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalculatedMember WithComCleanupProxy(this Microsoft.Office.Interop.Excel.CalculatedMember resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.CalculatedMember, Interfaces.ICalculatedMember>();
		}

		/// <summary>
		/// Wrapper interface for Watches which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWatches WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Watches resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Watches, Interfaces.IWatches>();
		}

		/// <summary>
		/// Wrapper interface for Watch which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWatch WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Watch resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Watch, Interfaces.IWatch>();
		}

		/// <summary>
		/// Wrapper interface for PivotCell which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotCell WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotCell resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotCell, Interfaces.IPivotCell>();
		}

		/// <summary>
		/// Wrapper interface for Graphic which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGraphic WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Graphic resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Graphic, Interfaces.IGraphic>();
		}

		/// <summary>
		/// Wrapper interface for AutoRecover which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoRecover WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AutoRecover resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AutoRecover, Interfaces.IAutoRecover>();
		}

		/// <summary>
		/// Wrapper interface for ErrorCheckingOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrorCheckingOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ErrorCheckingOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ErrorCheckingOptions, Interfaces.IErrorCheckingOptions>();
		}

		/// <summary>
		/// Wrapper interface for Errors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrors WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Errors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Errors, Interfaces.IErrors>();
		}

		/// <summary>
		/// Wrapper interface for Error which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IError WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Error resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Error, Interfaces.IError>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagAction WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTagAction resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTagAction, Interfaces.ISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagActions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTagActions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTagActions, Interfaces.ISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for SmartTag which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTag WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTag resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTag, Interfaces.ISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for SmartTags which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTags WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTags resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTags, Interfaces.ISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagRecognizer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTagRecognizer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTagRecognizer, Interfaces.ISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagRecognizers WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTagRecognizers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTagRecognizers, Interfaces.ISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISmartTagOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SmartTagOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SmartTagOptions, Interfaces.ISmartTagOptions>();
		}

		/// <summary>
		/// Wrapper interface for SpellingOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpellingOptions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SpellingOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SpellingOptions, Interfaces.ISpellingOptions>();
		}

		/// <summary>
		/// Wrapper interface for Speech which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISpeech WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Speech resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Speech, Interfaces.ISpeech>();
		}

		/// <summary>
		/// Wrapper interface for Protection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Protection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Protection, Interfaces.IProtection>();
		}

		/// <summary>
		/// Wrapper interface for PivotItemList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotItemList WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotItemList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotItemList, Interfaces.IPivotItemList>();
		}

		/// <summary>
		/// Wrapper interface for Tab which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITab WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Tab resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Tab, Interfaces.ITab>();
		}

		/// <summary>
		/// Wrapper interface for AllowEditRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAllowEditRanges WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AllowEditRanges resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AllowEditRanges, Interfaces.IAllowEditRanges>();
		}

		/// <summary>
		/// Wrapper interface for AllowEditRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAllowEditRange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AllowEditRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AllowEditRange, Interfaces.IAllowEditRange>();
		}

		/// <summary>
		/// Wrapper interface for UserAccessList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserAccessList WithComCleanupProxy(this Microsoft.Office.Interop.Excel.UserAccessList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.UserAccessList, Interfaces.IUserAccessList>();
		}

		/// <summary>
		/// Wrapper interface for UserAccess which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUserAccess WithComCleanupProxy(this Microsoft.Office.Interop.Excel.UserAccess resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.UserAccess, Interfaces.IUserAccess>();
		}

		/// <summary>
		/// Wrapper interface for RTD which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRTD WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RTD resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RTD, Interfaces.IRTD>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagram WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Diagram resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Diagram, Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for ListObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListObjects WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListObjects, Interfaces.IListObjects>();
		}

		/// <summary>
		/// Wrapper interface for ListObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListObject, Interfaces.IListObject>();
		}

		/// <summary>
		/// Wrapper interface for ListColumns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListColumns WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListColumns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListColumns, Interfaces.IListColumns>();
		}

		/// <summary>
		/// Wrapper interface for ListColumn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListColumn WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListColumn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListColumn, Interfaces.IListColumn>();
		}

		/// <summary>
		/// Wrapper interface for ListRows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListRows WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListRows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListRows, Interfaces.IListRows>();
		}

		/// <summary>
		/// Wrapper interface for ListRow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListRow WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListRow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListRow, Interfaces.IListRow>();
		}

		/// <summary>
		/// Wrapper interface for XmlNamespace which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlNamespace WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlNamespace resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlNamespace, Interfaces.IXmlNamespace>();
		}

		/// <summary>
		/// Wrapper interface for XmlNamespaces which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlNamespaces WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlNamespaces resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlNamespaces, Interfaces.IXmlNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for XmlDataBinding which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlDataBinding WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlDataBinding resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlDataBinding, Interfaces.IXmlDataBinding>();
		}

		/// <summary>
		/// Wrapper interface for XmlSchema which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlSchema WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlSchema resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlSchema, Interfaces.IXmlSchema>();
		}

		/// <summary>
		/// Wrapper interface for XmlSchemas which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlSchemas WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlSchemas resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlSchemas, Interfaces.IXmlSchemas>();
		}

		/// <summary>
		/// Wrapper interface for XmlMap which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlMap WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlMap resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlMap, Interfaces.IXmlMap>();
		}

		/// <summary>
		/// Wrapper interface for XmlMaps which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXmlMaps WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XmlMaps resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XmlMaps, Interfaces.IXmlMaps>();
		}

		/// <summary>
		/// Wrapper interface for ListDataFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IListDataFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ListDataFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ListDataFormat, Interfaces.IListDataFormat>();
		}

		/// <summary>
		/// Wrapper interface for XPath which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IXPath WithComCleanupProxy(this Microsoft.Office.Interop.Excel.XPath resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.XPath, Interfaces.IXPath>();
		}

		/// <summary>
		/// Wrapper interface for PivotLineCells which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotLineCells WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotLineCells resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotLineCells, Interfaces.IPivotLineCells>();
		}

		/// <summary>
		/// Wrapper interface for PivotLine which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotLine WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotLine resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotLine, Interfaces.IPivotLine>();
		}

		/// <summary>
		/// Wrapper interface for PivotLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotLines WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotLines, Interfaces.IPivotLines>();
		}

		/// <summary>
		/// Wrapper interface for PivotAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotAxis, Interfaces.IPivotAxis>();
		}

		/// <summary>
		/// Wrapper interface for PivotFilter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotFilter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotFilter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotFilter, Interfaces.IPivotFilter>();
		}

		/// <summary>
		/// Wrapper interface for PivotFilters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotFilters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotFilters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotFilters, Interfaces.IPivotFilters>();
		}

		/// <summary>
		/// Wrapper interface for WorkbookConnection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkbookConnection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.WorkbookConnection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.WorkbookConnection, Interfaces.IWorkbookConnection>();
		}

		/// <summary>
		/// Wrapper interface for Connections which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnections WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Connections resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Connections, Interfaces.IConnections>();
		}

		/// <summary>
		/// Wrapper interface for WorksheetView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorksheetView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.WorksheetView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.WorksheetView, Interfaces.IWorksheetView>();
		}

		/// <summary>
		/// Wrapper interface for ChartView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartView, Interfaces.IChartView>();
		}

		/// <summary>
		/// Wrapper interface for ModuleView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IModuleView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ModuleView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ModuleView, Interfaces.IModuleView>();
		}

		/// <summary>
		/// Wrapper interface for DialogSheetView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDialogSheetView WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DialogSheetView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DialogSheetView, Interfaces.IDialogSheetView>();
		}

		/// <summary>
		/// Wrapper interface for SheetViews which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISheetViews WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SheetViews resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SheetViews, Interfaces.ISheetViews>();
		}

		/// <summary>
		/// Wrapper interface for OLEDBConnection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEDBConnection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEDBConnection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEDBConnection, Interfaces.IOLEDBConnection>();
		}

		/// <summary>
		/// Wrapper interface for ODBCConnection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IODBCConnection WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ODBCConnection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ODBCConnection, Interfaces.IODBCConnection>();
		}

		/// <summary>
		/// Wrapper interface for Action which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAction WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Action resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Action, Interfaces.IAction>();
		}

		/// <summary>
		/// Wrapper interface for Actions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IActions WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Actions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Actions, Interfaces.IActions>();
		}

		/// <summary>
		/// Wrapper interface for FormatColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFormatColor WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FormatColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FormatColor, Interfaces.IFormatColor>();
		}

		/// <summary>
		/// Wrapper interface for ConditionValue which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConditionValue WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ConditionValue resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ConditionValue, Interfaces.IConditionValue>();
		}

		/// <summary>
		/// Wrapper interface for ColorScale which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorScale WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ColorScale resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ColorScale, Interfaces.IColorScale>();
		}

		/// <summary>
		/// Wrapper interface for ColorScaleCriteria which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorScaleCriteria WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ColorScaleCriteria resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ColorScaleCriteria, Interfaces.IColorScaleCriteria>();
		}

		/// <summary>
		/// Wrapper interface for ColorScaleCriterion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorScaleCriterion WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ColorScaleCriterion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ColorScaleCriterion, Interfaces.IColorScaleCriterion>();
		}

		/// <summary>
		/// Wrapper interface for Databar which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDatabar WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Databar resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Databar, Interfaces.IDatabar>();
		}

		/// <summary>
		/// Wrapper interface for IconSetCondition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIconSetCondition WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IconSetCondition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IconSetCondition, Interfaces.IIconSetCondition>();
		}

		/// <summary>
		/// Wrapper interface for IconCriteria which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIconCriteria WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IconCriteria resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IconCriteria, Interfaces.IIconCriteria>();
		}

		/// <summary>
		/// Wrapper interface for IconCriterion which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIconCriterion WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IconCriterion resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IconCriterion, Interfaces.IIconCriterion>();
		}

		/// <summary>
		/// Wrapper interface for Icon which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIcon WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Icon resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Icon, Interfaces.IIcon>();
		}

		/// <summary>
		/// Wrapper interface for IconSet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIconSet WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IconSet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IconSet, Interfaces.IIconSet>();
		}

		/// <summary>
		/// Wrapper interface for IconSets which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIconSets WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IconSets resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IconSets, Interfaces.IIconSets>();
		}

		/// <summary>
		/// Wrapper interface for Top10 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITop10 WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Top10 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Top10, Interfaces.ITop10>();
		}

		/// <summary>
		/// Wrapper interface for AboveAverage which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAboveAverage WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AboveAverage resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AboveAverage, Interfaces.IAboveAverage>();
		}

		/// <summary>
		/// Wrapper interface for UniqueValues which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUniqueValues WithComCleanupProxy(this Microsoft.Office.Interop.Excel.UniqueValues resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.UniqueValues, Interfaces.IUniqueValues>();
		}

		/// <summary>
		/// Wrapper interface for Ranges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRanges WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Ranges resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Ranges, Interfaces.IRanges>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeaderFooter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.HeaderFooter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.HeaderFooter, Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for Page which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPage WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Page resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Page, Interfaces.IPage>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPages WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Pages resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Pages, Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for ServerViewableItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IServerViewableItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ServerViewableItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ServerViewableItems, Interfaces.IServerViewableItems>();
		}

		/// <summary>
		/// Wrapper interface for TableStyleElement which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyleElement WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TableStyleElement resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TableStyleElement, Interfaces.ITableStyleElement>();
		}

		/// <summary>
		/// Wrapper interface for TableStyleElements which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyleElements WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TableStyleElements resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TableStyleElements, Interfaces.ITableStyleElements>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyle WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TableStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TableStyle, Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for TableStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyles WithComCleanupProxy(this Microsoft.Office.Interop.Excel.TableStyles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.TableStyles, Interfaces.ITableStyles>();
		}

		/// <summary>
		/// Wrapper interface for SortField which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISortField WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SortField resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SortField, Interfaces.ISortField>();
		}

		/// <summary>
		/// Wrapper interface for SortFields which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISortFields WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SortFields resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SortFields, Interfaces.ISortFields>();
		}

		/// <summary>
		/// Wrapper interface for Sort which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISort WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Sort resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Sort, Interfaces.ISort>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResearch WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Research resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Research, Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for ColorStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorStop WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ColorStop resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ColorStop, Interfaces.IColorStop>();
		}

		/// <summary>
		/// Wrapper interface for ColorStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorStops WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ColorStops resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ColorStops, Interfaces.IColorStops>();
		}

		/// <summary>
		/// Wrapper interface for LinearGradient which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinearGradient WithComCleanupProxy(this Microsoft.Office.Interop.Excel.LinearGradient resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.LinearGradient, Interfaces.ILinearGradient>();
		}

		/// <summary>
		/// Wrapper interface for RectangularGradient which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRectangularGradient WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RectangularGradient resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RectangularGradient, Interfaces.IRectangularGradient>();
		}

		/// <summary>
		/// Wrapper interface for MultiThreadedCalculation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMultiThreadedCalculation WithComCleanupProxy(this Microsoft.Office.Interop.Excel.MultiThreadedCalculation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.MultiThreadedCalculation, Interfaces.IMultiThreadedCalculation>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartFormat, Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for FileExportConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileExportConverter WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FileExportConverter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FileExportConverter, Interfaces.IFileExportConverter>();
		}

		/// <summary>
		/// Wrapper interface for FileExportConverters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileExportConverters WithComCleanupProxy(this Microsoft.Office.Interop.Excel.FileExportConverters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.FileExportConverters, Interfaces.IFileExportConverters>();
		}

		/// <summary>
		/// Wrapper interface for AddIns2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIns2 WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AddIns2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AddIns2, Interfaces.IAddIns2>();
		}

		/// <summary>
		/// Wrapper interface for SparklineGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparklineGroups WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparklineGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparklineGroups, Interfaces.ISparklineGroups>();
		}

		/// <summary>
		/// Wrapper interface for SparklineGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparklineGroup WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparklineGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparklineGroup, Interfaces.ISparklineGroup>();
		}

		/// <summary>
		/// Wrapper interface for SparkPoints which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparkPoints WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparkPoints resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparkPoints, Interfaces.ISparkPoints>();
		}

		/// <summary>
		/// Wrapper interface for Sparkline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparkline WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Sparkline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Sparkline, Interfaces.ISparkline>();
		}

		/// <summary>
		/// Wrapper interface for SparkAxes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparkAxes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparkAxes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparkAxes, Interfaces.ISparkAxes>();
		}

		/// <summary>
		/// Wrapper interface for SparkHorizontalAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparkHorizontalAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparkHorizontalAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparkHorizontalAxis, Interfaces.ISparkHorizontalAxis>();
		}

		/// <summary>
		/// Wrapper interface for SparkVerticalAxis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparkVerticalAxis WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparkVerticalAxis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparkVerticalAxis, Interfaces.ISparkVerticalAxis>();
		}

		/// <summary>
		/// Wrapper interface for SparkColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISparkColor WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SparkColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SparkColor, Interfaces.ISparkColor>();
		}

		/// <summary>
		/// Wrapper interface for DataBarBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataBarBorder WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DataBarBorder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DataBarBorder, Interfaces.IDataBarBorder>();
		}

		/// <summary>
		/// Wrapper interface for NegativeBarFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INegativeBarFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.NegativeBarFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.NegativeBarFormat, Interfaces.INegativeBarFormat>();
		}

		/// <summary>
		/// Wrapper interface for ValueChange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IValueChange WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ValueChange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ValueChange, Interfaces.IValueChange>();
		}

		/// <summary>
		/// Wrapper interface for PivotTableChangeList which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPivotTableChangeList WithComCleanupProxy(this Microsoft.Office.Interop.Excel.PivotTableChangeList resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.PivotTableChangeList, Interfaces.IPivotTableChangeList>();
		}

		/// <summary>
		/// Wrapper interface for DisplayFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDisplayFormat WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DisplayFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DisplayFormat, Interfaces.IDisplayFormat>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCaches which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerCaches WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerCaches resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerCaches, Interfaces.ISlicerCaches>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCache which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerCache WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerCache resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerCache, Interfaces.ISlicerCache>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCacheLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerCacheLevels WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerCacheLevels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerCacheLevels, Interfaces.ISlicerCacheLevels>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCacheLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerCacheLevel WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerCacheLevel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerCacheLevel, Interfaces.ISlicerCacheLevel>();
		}

		/// <summary>
		/// Wrapper interface for Slicers which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicers WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Slicers resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Slicers, Interfaces.ISlicers>();
		}

		/// <summary>
		/// Wrapper interface for Slicer which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicer WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Slicer resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Slicer, Interfaces.ISlicer>();
		}

		/// <summary>
		/// Wrapper interface for SlicerItem which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerItem WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerItem resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerItem, Interfaces.ISlicerItem>();
		}

		/// <summary>
		/// Wrapper interface for SlicerItems which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerItems WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerItems resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerItems, Interfaces.ISlicerItems>();
		}

		/// <summary>
		/// Wrapper interface for SlicerPivotTables which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlicerPivotTables WithComCleanupProxy(this Microsoft.Office.Interop.Excel.SlicerPivotTables resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.SlicerPivotTables, Interfaces.ISlicerPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindows WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ProtectedViewWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ProtectedViewWindows, Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindow WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ProtectedViewWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ProtectedViewWindow, Interfaces.IProtectedViewWindow>();
		}

		/// <summary>
		/// Wrapper interface for IDummy which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IIDummy WithComCleanupProxy(this Microsoft.Office.Interop.Excel.IDummy resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.IDummy, Interfaces.IIDummy>();
		}

		/// <summary>
		/// Wrapper interface for ICanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IICanvasShapes WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ICanvasShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ICanvasShapes, Interfaces.IICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for RefreshEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRefreshEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Excel.RefreshEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.RefreshEvents_Event, Interfaces.IRefreshEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for QueryTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IQueryTable WithComCleanupProxy(this Microsoft.Office.Interop.Excel.QueryTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.QueryTable, Interfaces.IQueryTable>();
		}

		/// <summary>
		/// Wrapper interface for AppEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAppEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Excel.AppEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.AppEvents_Event, Interfaces.IAppEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplication WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Application, Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for ChartEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Excel.ChartEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.ChartEvents_Event, Interfaces.IChartEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChart WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Chart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Chart, Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for DocEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Excel.DocEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.DocEvents_Event, Interfaces.IDocEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Worksheet which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorksheet WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Worksheet resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Worksheet, Interfaces.IWorksheet>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlobal WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Global resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Global, Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for WorkbookEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkbookEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Excel.WorkbookEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.WorkbookEvents_Event, Interfaces.IWorkbookEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Workbook which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWorkbook WithComCleanupProxy(this Microsoft.Office.Interop.Excel.Workbook resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.Workbook, Interfaces.IWorkbook>();
		}

		/// <summary>
		/// Wrapper interface for OLEObjectEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEObjectEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEObjectEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEObjectEvents_Event, Interfaces.IOLEObjectEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEObject WithComCleanupProxy(this Microsoft.Office.Interop.Excel.OLEObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.Excel.OLEObject, Interfaces.IOLEObject>();
		}

	}
}