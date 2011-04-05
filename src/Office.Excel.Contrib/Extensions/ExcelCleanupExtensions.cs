//Microsoft.Office.Interop.Excel, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace Office.Contrib.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAdjustments WithComCleanup(this Microsoft.Office.Interop.Excel.Adjustments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Adjustments, Excel.Contrib.Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICalloutFormat WithComCleanup(this Microsoft.Office.Interop.Excel.CalloutFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CalloutFormat, Excel.Contrib.Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IColorFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ColorFormat, Excel.Contrib.Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILineFormat WithComCleanup(this Microsoft.Office.Interop.Excel.LineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LineFormat, Excel.Contrib.Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IShapeNode WithComCleanup(this Microsoft.Office.Interop.Excel.ShapeNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ShapeNode, Excel.Contrib.Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IShapeNodes WithComCleanup(this Microsoft.Office.Interop.Excel.ShapeNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ShapeNodes, Excel.Contrib.Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPictureFormat WithComCleanup(this Microsoft.Office.Interop.Excel.PictureFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PictureFormat, Excel.Contrib.Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IShadowFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ShadowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ShadowFormat, Excel.Contrib.Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITextEffectFormat WithComCleanup(this Microsoft.Office.Interop.Excel.TextEffectFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TextEffectFormat, Excel.Contrib.Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IThreeDFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ThreeDFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ThreeDFormat, Excel.Contrib.Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFillFormat WithComCleanup(this Microsoft.Office.Interop.Excel.FillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FillFormat, Excel.Contrib.Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDiagramNodes WithComCleanup(this Microsoft.Office.Interop.Excel.DiagramNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DiagramNodes, Excel.Contrib.Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDiagramNodeChildren WithComCleanup(this Microsoft.Office.Interop.Excel.DiagramNodeChildren resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DiagramNodeChildren, Excel.Contrib.Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDiagramNode WithComCleanup(this Microsoft.Office.Interop.Excel.DiagramNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DiagramNode, Excel.Contrib.Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for IRTDUpdateEvent which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRTDUpdateEvent WithComCleanup(this Microsoft.Office.Interop.Excel.IRTDUpdateEvent resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRTDUpdateEvent, Excel.Contrib.Interfaces.IIRTDUpdateEvent>();
		}

		/// <summary>
		/// Wrapper interface for IRtdServer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRtdServer WithComCleanup(this Microsoft.Office.Interop.Excel.IRtdServer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRtdServer, Excel.Contrib.Interfaces.IIRtdServer>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITextFrame2 WithComCleanup(this Microsoft.Office.Interop.Excel.TextFrame2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TextFrame2, Excel.Contrib.Interfaces.ITextFrame2>();
		}

		/// <summary>
		/// Wrapper interface for IFont which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFont WithComCleanup(this Microsoft.Office.Interop.Excel.IFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFont, Excel.Contrib.Interfaces.IIFont>();
		}

		/// <summary>
		/// Wrapper interface for IWindow which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWindow WithComCleanup(this Microsoft.Office.Interop.Excel.IWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWindow, Excel.Contrib.Interfaces.IIWindow>();
		}

		/// <summary>
		/// Wrapper interface for IWindows which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWindows WithComCleanup(this Microsoft.Office.Interop.Excel.IWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWindows, Excel.Contrib.Interfaces.IIWindows>();
		}

		/// <summary>
		/// Wrapper interface for IAppEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAppEvents WithComCleanup(this Microsoft.Office.Interop.Excel.IAppEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAppEvents, Excel.Contrib.Interfaces.IIAppEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_Application WithComCleanup(this Microsoft.Office.Interop.Excel._Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._Application, Excel.Contrib.Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for IWorksheetFunction which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWorksheetFunction WithComCleanup(this Microsoft.Office.Interop.Excel.IWorksheetFunction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWorksheetFunction, Excel.Contrib.Interfaces.IIWorksheetFunction>();
		}

		/// <summary>
		/// Wrapper interface for IRange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRange WithComCleanup(this Microsoft.Office.Interop.Excel.IRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRange, Excel.Contrib.Interfaces.IIRange>();
		}

		/// <summary>
		/// Wrapper interface for IChartEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartEvents WithComCleanup(this Microsoft.Office.Interop.Excel.IChartEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartEvents, Excel.Contrib.Interfaces.IIChartEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Chart which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_Chart WithComCleanup(this Microsoft.Office.Interop.Excel._Chart resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._Chart, Excel.Contrib.Interfaces.I_Chart>();
		}

		/// <summary>
		/// Wrapper interface for Sheets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISheets WithComCleanup(this Microsoft.Office.Interop.Excel.Sheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Sheets, Excel.Contrib.Interfaces.ISheets>();
		}

		/// <summary>
		/// Wrapper interface for IVPageBreak which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIVPageBreak WithComCleanup(this Microsoft.Office.Interop.Excel.IVPageBreak resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IVPageBreak, Excel.Contrib.Interfaces.IIVPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for IHPageBreak which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIHPageBreak WithComCleanup(this Microsoft.Office.Interop.Excel.IHPageBreak resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IHPageBreak, Excel.Contrib.Interfaces.IIHPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for IHPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIHPageBreaks WithComCleanup(this Microsoft.Office.Interop.Excel.IHPageBreaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IHPageBreaks, Excel.Contrib.Interfaces.IIHPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for IVPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIVPageBreaks WithComCleanup(this Microsoft.Office.Interop.Excel.IVPageBreaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IVPageBreaks, Excel.Contrib.Interfaces.IIVPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for IRecentFile which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRecentFile WithComCleanup(this Microsoft.Office.Interop.Excel.IRecentFile resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRecentFile, Excel.Contrib.Interfaces.IIRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for IRecentFiles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRecentFiles WithComCleanup(this Microsoft.Office.Interop.Excel.IRecentFiles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRecentFiles, Excel.Contrib.Interfaces.IIRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for IDocEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDocEvents WithComCleanup(this Microsoft.Office.Interop.Excel.IDocEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDocEvents, Excel.Contrib.Interfaces.IIDocEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Worksheet which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_Worksheet WithComCleanup(this Microsoft.Office.Interop.Excel._Worksheet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._Worksheet, Excel.Contrib.Interfaces.I_Worksheet>();
		}

		/// <summary>
		/// Wrapper interface for IStyle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIStyle WithComCleanup(this Microsoft.Office.Interop.Excel.IStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IStyle, Excel.Contrib.Interfaces.IIStyle>();
		}

		/// <summary>
		/// Wrapper interface for IStyles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIStyles WithComCleanup(this Microsoft.Office.Interop.Excel.IStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IStyles, Excel.Contrib.Interfaces.IIStyles>();
		}

		/// <summary>
		/// Wrapper interface for IBorders which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIBorders WithComCleanup(this Microsoft.Office.Interop.Excel.IBorders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IBorders, Excel.Contrib.Interfaces.IIBorders>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_Global WithComCleanup(this Microsoft.Office.Interop.Excel._Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._Global, Excel.Contrib.Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for IAddIn which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAddIn WithComCleanup(this Microsoft.Office.Interop.Excel.IAddIn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAddIn, Excel.Contrib.Interfaces.IIAddIn>();
		}

		/// <summary>
		/// Wrapper interface for IAddIns which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAddIns WithComCleanup(this Microsoft.Office.Interop.Excel.IAddIns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAddIns, Excel.Contrib.Interfaces.IIAddIns>();
		}

		/// <summary>
		/// Wrapper interface for IToolbar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIToolbar WithComCleanup(this Microsoft.Office.Interop.Excel.IToolbar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IToolbar, Excel.Contrib.Interfaces.IIToolbar>();
		}

		/// <summary>
		/// Wrapper interface for IToolbars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIToolbars WithComCleanup(this Microsoft.Office.Interop.Excel.IToolbars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IToolbars, Excel.Contrib.Interfaces.IIToolbars>();
		}

		/// <summary>
		/// Wrapper interface for IToolbarButton which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIToolbarButton WithComCleanup(this Microsoft.Office.Interop.Excel.IToolbarButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IToolbarButton, Excel.Contrib.Interfaces.IIToolbarButton>();
		}

		/// <summary>
		/// Wrapper interface for IToolbarButtons which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIToolbarButtons WithComCleanup(this Microsoft.Office.Interop.Excel.IToolbarButtons resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IToolbarButtons, Excel.Contrib.Interfaces.IIToolbarButtons>();
		}

		/// <summary>
		/// Wrapper interface for IAreas which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAreas WithComCleanup(this Microsoft.Office.Interop.Excel.IAreas resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAreas, Excel.Contrib.Interfaces.IIAreas>();
		}

		/// <summary>
		/// Wrapper interface for IWorkbookEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWorkbookEvents WithComCleanup(this Microsoft.Office.Interop.Excel.IWorkbookEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWorkbookEvents, Excel.Contrib.Interfaces.IIWorkbookEvents>();
		}

		/// <summary>
		/// Wrapper interface for _Workbook which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_Workbook WithComCleanup(this Microsoft.Office.Interop.Excel._Workbook resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._Workbook, Excel.Contrib.Interfaces.I_Workbook>();
		}

		/// <summary>
		/// Wrapper interface for Workbooks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorkbooks WithComCleanup(this Microsoft.Office.Interop.Excel.Workbooks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Workbooks, Excel.Contrib.Interfaces.IWorkbooks>();
		}

		/// <summary>
		/// Wrapper interface for IMenuBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMenuBars WithComCleanup(this Microsoft.Office.Interop.Excel.IMenuBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMenuBars, Excel.Contrib.Interfaces.IIMenuBars>();
		}

		/// <summary>
		/// Wrapper interface for IMenuBar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMenuBar WithComCleanup(this Microsoft.Office.Interop.Excel.IMenuBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMenuBar, Excel.Contrib.Interfaces.IIMenuBar>();
		}

		/// <summary>
		/// Wrapper interface for IMenus which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMenus WithComCleanup(this Microsoft.Office.Interop.Excel.IMenus resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMenus, Excel.Contrib.Interfaces.IIMenus>();
		}

		/// <summary>
		/// Wrapper interface for IMenu which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMenu WithComCleanup(this Microsoft.Office.Interop.Excel.IMenu resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMenu, Excel.Contrib.Interfaces.IIMenu>();
		}

		/// <summary>
		/// Wrapper interface for IMenuItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMenuItems WithComCleanup(this Microsoft.Office.Interop.Excel.IMenuItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMenuItems, Excel.Contrib.Interfaces.IIMenuItems>();
		}

		/// <summary>
		/// Wrapper interface for IMenuItem which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMenuItem WithComCleanup(this Microsoft.Office.Interop.Excel.IMenuItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMenuItem, Excel.Contrib.Interfaces.IIMenuItem>();
		}

		/// <summary>
		/// Wrapper interface for ICharts which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICharts WithComCleanup(this Microsoft.Office.Interop.Excel.ICharts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICharts, Excel.Contrib.Interfaces.IICharts>();
		}

		/// <summary>
		/// Wrapper interface for IDrawingObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDrawingObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IDrawingObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDrawingObjects, Excel.Contrib.Interfaces.IIDrawingObjects>();
		}

		/// <summary>
		/// Wrapper interface for IPivotCache which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotCache WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotCache resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotCache, Excel.Contrib.Interfaces.IIPivotCache>();
		}

		/// <summary>
		/// Wrapper interface for IPivotCaches which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotCaches WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotCaches resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotCaches, Excel.Contrib.Interfaces.IIPivotCaches>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFormula which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotFormula WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotFormula resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotFormula, Excel.Contrib.Interfaces.IIPivotFormula>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFormulas which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotFormulas WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotFormulas resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotFormulas, Excel.Contrib.Interfaces.IIPivotFormulas>();
		}

		/// <summary>
		/// Wrapper interface for IPivotTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotTable WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotTable, Excel.Contrib.Interfaces.IIPivotTable>();
		}

		/// <summary>
		/// Wrapper interface for IPivotTables which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotTables WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotTables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotTables, Excel.Contrib.Interfaces.IIPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for IPivotField which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotField WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotField, Excel.Contrib.Interfaces.IIPivotField>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotFields WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotFields, Excel.Contrib.Interfaces.IIPivotFields>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICalculatedFields WithComCleanup(this Microsoft.Office.Interop.Excel.ICalculatedFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICalculatedFields, Excel.Contrib.Interfaces.IICalculatedFields>();
		}

		/// <summary>
		/// Wrapper interface for IPivotItem which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotItem WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotItem, Excel.Contrib.Interfaces.IIPivotItem>();
		}

		/// <summary>
		/// Wrapper interface for IPivotItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotItems WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotItems, Excel.Contrib.Interfaces.IIPivotItems>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICalculatedItems WithComCleanup(this Microsoft.Office.Interop.Excel.ICalculatedItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICalculatedItems, Excel.Contrib.Interfaces.IICalculatedItems>();
		}

		/// <summary>
		/// Wrapper interface for ICharacters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICharacters WithComCleanup(this Microsoft.Office.Interop.Excel.ICharacters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICharacters, Excel.Contrib.Interfaces.IICharacters>();
		}

		/// <summary>
		/// Wrapper interface for IDialogs which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDialogs WithComCleanup(this Microsoft.Office.Interop.Excel.IDialogs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDialogs, Excel.Contrib.Interfaces.IIDialogs>();
		}

		/// <summary>
		/// Wrapper interface for IDialog which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDialog WithComCleanup(this Microsoft.Office.Interop.Excel.IDialog resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDialog, Excel.Contrib.Interfaces.IIDialog>();
		}

		/// <summary>
		/// Wrapper interface for ISoundNote which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISoundNote WithComCleanup(this Microsoft.Office.Interop.Excel.ISoundNote resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISoundNote, Excel.Contrib.Interfaces.IISoundNote>();
		}

		/// <summary>
		/// Wrapper interface for IButton which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIButton WithComCleanup(this Microsoft.Office.Interop.Excel.IButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IButton, Excel.Contrib.Interfaces.IIButton>();
		}

		/// <summary>
		/// Wrapper interface for IButtons which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIButtons WithComCleanup(this Microsoft.Office.Interop.Excel.IButtons resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IButtons, Excel.Contrib.Interfaces.IIButtons>();
		}

		/// <summary>
		/// Wrapper interface for ICheckBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICheckBox WithComCleanup(this Microsoft.Office.Interop.Excel.ICheckBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICheckBox, Excel.Contrib.Interfaces.IICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for ICheckBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICheckBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.ICheckBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICheckBoxes, Excel.Contrib.Interfaces.IICheckBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IOptionButton which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOptionButton WithComCleanup(this Microsoft.Office.Interop.Excel.IOptionButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOptionButton, Excel.Contrib.Interfaces.IIOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for IOptionButtons which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOptionButtons WithComCleanup(this Microsoft.Office.Interop.Excel.IOptionButtons resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOptionButtons, Excel.Contrib.Interfaces.IIOptionButtons>();
		}

		/// <summary>
		/// Wrapper interface for IEditBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIEditBox WithComCleanup(this Microsoft.Office.Interop.Excel.IEditBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IEditBox, Excel.Contrib.Interfaces.IIEditBox>();
		}

		/// <summary>
		/// Wrapper interface for IEditBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIEditBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.IEditBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IEditBoxes, Excel.Contrib.Interfaces.IIEditBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IScrollBar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIScrollBar WithComCleanup(this Microsoft.Office.Interop.Excel.IScrollBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IScrollBar, Excel.Contrib.Interfaces.IIScrollBar>();
		}

		/// <summary>
		/// Wrapper interface for IScrollBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIScrollBars WithComCleanup(this Microsoft.Office.Interop.Excel.IScrollBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IScrollBars, Excel.Contrib.Interfaces.IIScrollBars>();
		}

		/// <summary>
		/// Wrapper interface for IListBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListBox WithComCleanup(this Microsoft.Office.Interop.Excel.IListBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListBox, Excel.Contrib.Interfaces.IIListBox>();
		}

		/// <summary>
		/// Wrapper interface for IListBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.IListBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListBoxes, Excel.Contrib.Interfaces.IIListBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IGroupBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGroupBox WithComCleanup(this Microsoft.Office.Interop.Excel.IGroupBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGroupBox, Excel.Contrib.Interfaces.IIGroupBox>();
		}

		/// <summary>
		/// Wrapper interface for IGroupBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGroupBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.IGroupBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGroupBoxes, Excel.Contrib.Interfaces.IIGroupBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IDropDown which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDropDown WithComCleanup(this Microsoft.Office.Interop.Excel.IDropDown resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDropDown, Excel.Contrib.Interfaces.IIDropDown>();
		}

		/// <summary>
		/// Wrapper interface for IDropDowns which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDropDowns WithComCleanup(this Microsoft.Office.Interop.Excel.IDropDowns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDropDowns, Excel.Contrib.Interfaces.IIDropDowns>();
		}

		/// <summary>
		/// Wrapper interface for ISpinner which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISpinner WithComCleanup(this Microsoft.Office.Interop.Excel.ISpinner resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISpinner, Excel.Contrib.Interfaces.IISpinner>();
		}

		/// <summary>
		/// Wrapper interface for ISpinners which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISpinners WithComCleanup(this Microsoft.Office.Interop.Excel.ISpinners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISpinners, Excel.Contrib.Interfaces.IISpinners>();
		}

		/// <summary>
		/// Wrapper interface for IDialogFrame which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDialogFrame WithComCleanup(this Microsoft.Office.Interop.Excel.IDialogFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDialogFrame, Excel.Contrib.Interfaces.IIDialogFrame>();
		}

		/// <summary>
		/// Wrapper interface for ILabel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILabel WithComCleanup(this Microsoft.Office.Interop.Excel.ILabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILabel, Excel.Contrib.Interfaces.IILabel>();
		}

		/// <summary>
		/// Wrapper interface for ILabels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILabels WithComCleanup(this Microsoft.Office.Interop.Excel.ILabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILabels, Excel.Contrib.Interfaces.IILabels>();
		}

		/// <summary>
		/// Wrapper interface for IPanes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPanes WithComCleanup(this Microsoft.Office.Interop.Excel.IPanes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPanes, Excel.Contrib.Interfaces.IIPanes>();
		}

		/// <summary>
		/// Wrapper interface for IPane which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPane WithComCleanup(this Microsoft.Office.Interop.Excel.IPane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPane, Excel.Contrib.Interfaces.IIPane>();
		}

		/// <summary>
		/// Wrapper interface for IScenarios which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIScenarios WithComCleanup(this Microsoft.Office.Interop.Excel.IScenarios resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IScenarios, Excel.Contrib.Interfaces.IIScenarios>();
		}

		/// <summary>
		/// Wrapper interface for IScenario which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIScenario WithComCleanup(this Microsoft.Office.Interop.Excel.IScenario resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IScenario, Excel.Contrib.Interfaces.IIScenario>();
		}

		/// <summary>
		/// Wrapper interface for IGroupObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGroupObject WithComCleanup(this Microsoft.Office.Interop.Excel.IGroupObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGroupObject, Excel.Contrib.Interfaces.IIGroupObject>();
		}

		/// <summary>
		/// Wrapper interface for IGroupObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGroupObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IGroupObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGroupObjects, Excel.Contrib.Interfaces.IIGroupObjects>();
		}

		/// <summary>
		/// Wrapper interface for ILine which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILine WithComCleanup(this Microsoft.Office.Interop.Excel.ILine resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILine, Excel.Contrib.Interfaces.IILine>();
		}

		/// <summary>
		/// Wrapper interface for ILines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILines WithComCleanup(this Microsoft.Office.Interop.Excel.ILines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILines, Excel.Contrib.Interfaces.IILines>();
		}

		/// <summary>
		/// Wrapper interface for IRectangle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRectangle WithComCleanup(this Microsoft.Office.Interop.Excel.IRectangle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRectangle, Excel.Contrib.Interfaces.IIRectangle>();
		}

		/// <summary>
		/// Wrapper interface for IRectangles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRectangles WithComCleanup(this Microsoft.Office.Interop.Excel.IRectangles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRectangles, Excel.Contrib.Interfaces.IIRectangles>();
		}

		/// <summary>
		/// Wrapper interface for IOval which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOval WithComCleanup(this Microsoft.Office.Interop.Excel.IOval resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOval, Excel.Contrib.Interfaces.IIOval>();
		}

		/// <summary>
		/// Wrapper interface for IOvals which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOvals WithComCleanup(this Microsoft.Office.Interop.Excel.IOvals resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOvals, Excel.Contrib.Interfaces.IIOvals>();
		}

		/// <summary>
		/// Wrapper interface for IArc which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIArc WithComCleanup(this Microsoft.Office.Interop.Excel.IArc resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IArc, Excel.Contrib.Interfaces.IIArc>();
		}

		/// <summary>
		/// Wrapper interface for IArcs which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIArcs WithComCleanup(this Microsoft.Office.Interop.Excel.IArcs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IArcs, Excel.Contrib.Interfaces.IIArcs>();
		}

		/// <summary>
		/// Wrapper interface for IOLEObjectEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOLEObjectEvents WithComCleanup(this Microsoft.Office.Interop.Excel.IOLEObjectEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOLEObjectEvents, Excel.Contrib.Interfaces.IIOLEObjectEvents>();
		}

		/// <summary>
		/// Wrapper interface for _IOLEObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_IOLEObject WithComCleanup(this Microsoft.Office.Interop.Excel._IOLEObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._IOLEObject, Excel.Contrib.Interfaces.I_IOLEObject>();
		}

		/// <summary>
		/// Wrapper interface for IOLEObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOLEObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IOLEObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOLEObjects, Excel.Contrib.Interfaces.IIOLEObjects>();
		}

		/// <summary>
		/// Wrapper interface for ITextBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITextBox WithComCleanup(this Microsoft.Office.Interop.Excel.ITextBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITextBox, Excel.Contrib.Interfaces.IITextBox>();
		}

		/// <summary>
		/// Wrapper interface for ITextBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITextBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.ITextBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITextBoxes, Excel.Contrib.Interfaces.IITextBoxes>();
		}

		/// <summary>
		/// Wrapper interface for IPicture which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPicture WithComCleanup(this Microsoft.Office.Interop.Excel.IPicture resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPicture, Excel.Contrib.Interfaces.IIPicture>();
		}

		/// <summary>
		/// Wrapper interface for IPictures which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPictures WithComCleanup(this Microsoft.Office.Interop.Excel.IPictures resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPictures, Excel.Contrib.Interfaces.IIPictures>();
		}

		/// <summary>
		/// Wrapper interface for IDrawing which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDrawing WithComCleanup(this Microsoft.Office.Interop.Excel.IDrawing resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDrawing, Excel.Contrib.Interfaces.IIDrawing>();
		}

		/// <summary>
		/// Wrapper interface for IDrawings which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDrawings WithComCleanup(this Microsoft.Office.Interop.Excel.IDrawings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDrawings, Excel.Contrib.Interfaces.IIDrawings>();
		}

		/// <summary>
		/// Wrapper interface for IRoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRoutingSlip WithComCleanup(this Microsoft.Office.Interop.Excel.IRoutingSlip resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRoutingSlip, Excel.Contrib.Interfaces.IIRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for IOutline which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOutline WithComCleanup(this Microsoft.Office.Interop.Excel.IOutline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOutline, Excel.Contrib.Interfaces.IIOutline>();
		}

		/// <summary>
		/// Wrapper interface for IModule which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIModule WithComCleanup(this Microsoft.Office.Interop.Excel.IModule resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IModule, Excel.Contrib.Interfaces.IIModule>();
		}

		/// <summary>
		/// Wrapper interface for IModules which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIModules WithComCleanup(this Microsoft.Office.Interop.Excel.IModules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IModules, Excel.Contrib.Interfaces.IIModules>();
		}

		/// <summary>
		/// Wrapper interface for IDialogSheet which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDialogSheet WithComCleanup(this Microsoft.Office.Interop.Excel.IDialogSheet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDialogSheet, Excel.Contrib.Interfaces.IIDialogSheet>();
		}

		/// <summary>
		/// Wrapper interface for IDialogSheets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDialogSheets WithComCleanup(this Microsoft.Office.Interop.Excel.IDialogSheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDialogSheets, Excel.Contrib.Interfaces.IIDialogSheets>();
		}

		/// <summary>
		/// Wrapper interface for IWorksheets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWorksheets WithComCleanup(this Microsoft.Office.Interop.Excel.IWorksheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWorksheets, Excel.Contrib.Interfaces.IIWorksheets>();
		}

		/// <summary>
		/// Wrapper interface for IPageSetup which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPageSetup WithComCleanup(this Microsoft.Office.Interop.Excel.IPageSetup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPageSetup, Excel.Contrib.Interfaces.IIPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for INames which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IINames WithComCleanup(this Microsoft.Office.Interop.Excel.INames resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.INames, Excel.Contrib.Interfaces.IINames>();
		}

		/// <summary>
		/// Wrapper interface for IName which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIName WithComCleanup(this Microsoft.Office.Interop.Excel.IName resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IName, Excel.Contrib.Interfaces.IIName>();
		}

		/// <summary>
		/// Wrapper interface for IChartObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartObject WithComCleanup(this Microsoft.Office.Interop.Excel.IChartObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartObject, Excel.Contrib.Interfaces.IIChartObject>();
		}

		/// <summary>
		/// Wrapper interface for IChartObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IChartObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartObjects, Excel.Contrib.Interfaces.IIChartObjects>();
		}

		/// <summary>
		/// Wrapper interface for IMailer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMailer WithComCleanup(this Microsoft.Office.Interop.Excel.IMailer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMailer, Excel.Contrib.Interfaces.IIMailer>();
		}

		/// <summary>
		/// Wrapper interface for ICustomViews which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICustomViews WithComCleanup(this Microsoft.Office.Interop.Excel.ICustomViews resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICustomViews, Excel.Contrib.Interfaces.IICustomViews>();
		}

		/// <summary>
		/// Wrapper interface for ICustomView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICustomView WithComCleanup(this Microsoft.Office.Interop.Excel.ICustomView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICustomView, Excel.Contrib.Interfaces.IICustomView>();
		}

		/// <summary>
		/// Wrapper interface for IFormatConditions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFormatConditions WithComCleanup(this Microsoft.Office.Interop.Excel.IFormatConditions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFormatConditions, Excel.Contrib.Interfaces.IIFormatConditions>();
		}

		/// <summary>
		/// Wrapper interface for IFormatCondition which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFormatCondition WithComCleanup(this Microsoft.Office.Interop.Excel.IFormatCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFormatCondition, Excel.Contrib.Interfaces.IIFormatCondition>();
		}

		/// <summary>
		/// Wrapper interface for IComments which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIComments WithComCleanup(this Microsoft.Office.Interop.Excel.IComments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IComments, Excel.Contrib.Interfaces.IIComments>();
		}

		/// <summary>
		/// Wrapper interface for IComment which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIComment WithComCleanup(this Microsoft.Office.Interop.Excel.IComment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IComment, Excel.Contrib.Interfaces.IIComment>();
		}

		/// <summary>
		/// Wrapper interface for IRefreshEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRefreshEvents WithComCleanup(this Microsoft.Office.Interop.Excel.IRefreshEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRefreshEvents, Excel.Contrib.Interfaces.IIRefreshEvents>();
		}

		/// <summary>
		/// Wrapper interface for _IQueryTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_IQueryTable WithComCleanup(this Microsoft.Office.Interop.Excel._IQueryTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._IQueryTable, Excel.Contrib.Interfaces.I_IQueryTable>();
		}

		/// <summary>
		/// Wrapper interface for IQueryTables which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIQueryTables WithComCleanup(this Microsoft.Office.Interop.Excel.IQueryTables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IQueryTables, Excel.Contrib.Interfaces.IIQueryTables>();
		}

		/// <summary>
		/// Wrapper interface for IParameter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIParameter WithComCleanup(this Microsoft.Office.Interop.Excel.IParameter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IParameter, Excel.Contrib.Interfaces.IIParameter>();
		}

		/// <summary>
		/// Wrapper interface for IParameters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIParameters WithComCleanup(this Microsoft.Office.Interop.Excel.IParameters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IParameters, Excel.Contrib.Interfaces.IIParameters>();
		}

		/// <summary>
		/// Wrapper interface for IODBCError which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIODBCError WithComCleanup(this Microsoft.Office.Interop.Excel.IODBCError resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IODBCError, Excel.Contrib.Interfaces.IIODBCError>();
		}

		/// <summary>
		/// Wrapper interface for IODBCErrors which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIODBCErrors WithComCleanup(this Microsoft.Office.Interop.Excel.IODBCErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IODBCErrors, Excel.Contrib.Interfaces.IIODBCErrors>();
		}

		/// <summary>
		/// Wrapper interface for IValidation which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIValidation WithComCleanup(this Microsoft.Office.Interop.Excel.IValidation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IValidation, Excel.Contrib.Interfaces.IIValidation>();
		}

		/// <summary>
		/// Wrapper interface for IHyperlinks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIHyperlinks WithComCleanup(this Microsoft.Office.Interop.Excel.IHyperlinks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IHyperlinks, Excel.Contrib.Interfaces.IIHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for IHyperlink which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIHyperlink WithComCleanup(this Microsoft.Office.Interop.Excel.IHyperlink resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IHyperlink, Excel.Contrib.Interfaces.IIHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for IAutoFilter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAutoFilter WithComCleanup(this Microsoft.Office.Interop.Excel.IAutoFilter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAutoFilter, Excel.Contrib.Interfaces.IIAutoFilter>();
		}

		/// <summary>
		/// Wrapper interface for IFilters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFilters WithComCleanup(this Microsoft.Office.Interop.Excel.IFilters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFilters, Excel.Contrib.Interfaces.IIFilters>();
		}

		/// <summary>
		/// Wrapper interface for IFilter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFilter WithComCleanup(this Microsoft.Office.Interop.Excel.IFilter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFilter, Excel.Contrib.Interfaces.IIFilter>();
		}

		/// <summary>
		/// Wrapper interface for IAutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAutoCorrect WithComCleanup(this Microsoft.Office.Interop.Excel.IAutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAutoCorrect, Excel.Contrib.Interfaces.IIAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for IBorder which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIBorder WithComCleanup(this Microsoft.Office.Interop.Excel.IBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IBorder, Excel.Contrib.Interfaces.IIBorder>();
		}

		/// <summary>
		/// Wrapper interface for IInterior which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIInterior WithComCleanup(this Microsoft.Office.Interop.Excel.IInterior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IInterior, Excel.Contrib.Interfaces.IIInterior>();
		}

		/// <summary>
		/// Wrapper interface for IChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartFillFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IChartFillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartFillFormat, Excel.Contrib.Interfaces.IIChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for IChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartColorFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IChartColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartColorFormat, Excel.Contrib.Interfaces.IIChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for IAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAxis WithComCleanup(this Microsoft.Office.Interop.Excel.IAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAxis, Excel.Contrib.Interfaces.IIAxis>();
		}

		/// <summary>
		/// Wrapper interface for IChartTitle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartTitle WithComCleanup(this Microsoft.Office.Interop.Excel.IChartTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartTitle, Excel.Contrib.Interfaces.IIChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for IAxisTitle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAxisTitle WithComCleanup(this Microsoft.Office.Interop.Excel.IAxisTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAxisTitle, Excel.Contrib.Interfaces.IIAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for IChartGroup which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartGroup WithComCleanup(this Microsoft.Office.Interop.Excel.IChartGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartGroup, Excel.Contrib.Interfaces.IIChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for IChartGroups which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartGroups WithComCleanup(this Microsoft.Office.Interop.Excel.IChartGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartGroups, Excel.Contrib.Interfaces.IIChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for IAxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAxes WithComCleanup(this Microsoft.Office.Interop.Excel.IAxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAxes, Excel.Contrib.Interfaces.IIAxes>();
		}

		/// <summary>
		/// Wrapper interface for IPoints which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPoints WithComCleanup(this Microsoft.Office.Interop.Excel.IPoints resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPoints, Excel.Contrib.Interfaces.IIPoints>();
		}

		/// <summary>
		/// Wrapper interface for IPoint which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPoint WithComCleanup(this Microsoft.Office.Interop.Excel.IPoint resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPoint, Excel.Contrib.Interfaces.IIPoint>();
		}

		/// <summary>
		/// Wrapper interface for ISeries which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISeries WithComCleanup(this Microsoft.Office.Interop.Excel.ISeries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISeries, Excel.Contrib.Interfaces.IISeries>();
		}

		/// <summary>
		/// Wrapper interface for ISeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISeriesCollection WithComCleanup(this Microsoft.Office.Interop.Excel.ISeriesCollection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISeriesCollection, Excel.Contrib.Interfaces.IISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for IDataLabel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDataLabel WithComCleanup(this Microsoft.Office.Interop.Excel.IDataLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDataLabel, Excel.Contrib.Interfaces.IIDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for IDataLabels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDataLabels WithComCleanup(this Microsoft.Office.Interop.Excel.IDataLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDataLabels, Excel.Contrib.Interfaces.IIDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for ILegendEntry which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILegendEntry WithComCleanup(this Microsoft.Office.Interop.Excel.ILegendEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILegendEntry, Excel.Contrib.Interfaces.IILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for ILegendEntries which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILegendEntries WithComCleanup(this Microsoft.Office.Interop.Excel.ILegendEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILegendEntries, Excel.Contrib.Interfaces.IILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for ILegendKey which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILegendKey WithComCleanup(this Microsoft.Office.Interop.Excel.ILegendKey resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILegendKey, Excel.Contrib.Interfaces.IILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for ITrendlines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITrendlines WithComCleanup(this Microsoft.Office.Interop.Excel.ITrendlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITrendlines, Excel.Contrib.Interfaces.IITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for ITrendline which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITrendline WithComCleanup(this Microsoft.Office.Interop.Excel.ITrendline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITrendline, Excel.Contrib.Interfaces.IITrendline>();
		}

		/// <summary>
		/// Wrapper interface for ICorners which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICorners WithComCleanup(this Microsoft.Office.Interop.Excel.ICorners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICorners, Excel.Contrib.Interfaces.IICorners>();
		}

		/// <summary>
		/// Wrapper interface for ISeriesLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISeriesLines WithComCleanup(this Microsoft.Office.Interop.Excel.ISeriesLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISeriesLines, Excel.Contrib.Interfaces.IISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for IHiLoLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIHiLoLines WithComCleanup(this Microsoft.Office.Interop.Excel.IHiLoLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IHiLoLines, Excel.Contrib.Interfaces.IIHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for IGridlines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGridlines WithComCleanup(this Microsoft.Office.Interop.Excel.IGridlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGridlines, Excel.Contrib.Interfaces.IIGridlines>();
		}

		/// <summary>
		/// Wrapper interface for IDropLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDropLines WithComCleanup(this Microsoft.Office.Interop.Excel.IDropLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDropLines, Excel.Contrib.Interfaces.IIDropLines>();
		}

		/// <summary>
		/// Wrapper interface for ILeaderLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILeaderLines WithComCleanup(this Microsoft.Office.Interop.Excel.ILeaderLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILeaderLines, Excel.Contrib.Interfaces.IILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for IUpBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIUpBars WithComCleanup(this Microsoft.Office.Interop.Excel.IUpBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IUpBars, Excel.Contrib.Interfaces.IIUpBars>();
		}

		/// <summary>
		/// Wrapper interface for IDownBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDownBars WithComCleanup(this Microsoft.Office.Interop.Excel.IDownBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDownBars, Excel.Contrib.Interfaces.IIDownBars>();
		}

		/// <summary>
		/// Wrapper interface for IFloor which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFloor WithComCleanup(this Microsoft.Office.Interop.Excel.IFloor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFloor, Excel.Contrib.Interfaces.IIFloor>();
		}

		/// <summary>
		/// Wrapper interface for IWalls which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWalls WithComCleanup(this Microsoft.Office.Interop.Excel.IWalls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWalls, Excel.Contrib.Interfaces.IIWalls>();
		}

		/// <summary>
		/// Wrapper interface for ITickLabels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITickLabels WithComCleanup(this Microsoft.Office.Interop.Excel.ITickLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITickLabels, Excel.Contrib.Interfaces.IITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for IPlotArea which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPlotArea WithComCleanup(this Microsoft.Office.Interop.Excel.IPlotArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPlotArea, Excel.Contrib.Interfaces.IIPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for IChartArea which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartArea WithComCleanup(this Microsoft.Office.Interop.Excel.IChartArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartArea, Excel.Contrib.Interfaces.IIChartArea>();
		}

		/// <summary>
		/// Wrapper interface for ILegend which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILegend WithComCleanup(this Microsoft.Office.Interop.Excel.ILegend resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILegend, Excel.Contrib.Interfaces.IILegend>();
		}

		/// <summary>
		/// Wrapper interface for IErrorBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIErrorBars WithComCleanup(this Microsoft.Office.Interop.Excel.IErrorBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IErrorBars, Excel.Contrib.Interfaces.IIErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for IDataTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDataTable WithComCleanup(this Microsoft.Office.Interop.Excel.IDataTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDataTable, Excel.Contrib.Interfaces.IIDataTable>();
		}

		/// <summary>
		/// Wrapper interface for IPhonetic which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPhonetic WithComCleanup(this Microsoft.Office.Interop.Excel.IPhonetic resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPhonetic, Excel.Contrib.Interfaces.IIPhonetic>();
		}

		/// <summary>
		/// Wrapper interface for IShape which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIShape WithComCleanup(this Microsoft.Office.Interop.Excel.IShape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IShape, Excel.Contrib.Interfaces.IIShape>();
		}

		/// <summary>
		/// Wrapper interface for IShapes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIShapes WithComCleanup(this Microsoft.Office.Interop.Excel.IShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IShapes, Excel.Contrib.Interfaces.IIShapes>();
		}

		/// <summary>
		/// Wrapper interface for IShapeRange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIShapeRange WithComCleanup(this Microsoft.Office.Interop.Excel.IShapeRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IShapeRange, Excel.Contrib.Interfaces.IIShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for IGroupShapes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGroupShapes WithComCleanup(this Microsoft.Office.Interop.Excel.IGroupShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGroupShapes, Excel.Contrib.Interfaces.IIGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for ITextFrame which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITextFrame WithComCleanup(this Microsoft.Office.Interop.Excel.ITextFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITextFrame, Excel.Contrib.Interfaces.IITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for IConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIConnectorFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IConnectorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IConnectorFormat, Excel.Contrib.Interfaces.IIConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for IFreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFreeformBuilder WithComCleanup(this Microsoft.Office.Interop.Excel.IFreeformBuilder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFreeformBuilder, Excel.Contrib.Interfaces.IIFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for IControlFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIControlFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IControlFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IControlFormat, Excel.Contrib.Interfaces.IIControlFormat>();
		}

		/// <summary>
		/// Wrapper interface for IOLEFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOLEFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IOLEFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOLEFormat, Excel.Contrib.Interfaces.IIOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for ILinkFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILinkFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ILinkFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILinkFormat, Excel.Contrib.Interfaces.IILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for IPublishObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPublishObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IPublishObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPublishObjects, Excel.Contrib.Interfaces.IIPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for PublishObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPublishObject WithComCleanup(this Microsoft.Office.Interop.Excel.PublishObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PublishObject, Excel.Contrib.Interfaces.IPublishObject>();
		}

		/// <summary>
		/// Wrapper interface for IOLEDBError which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOLEDBError WithComCleanup(this Microsoft.Office.Interop.Excel.IOLEDBError resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOLEDBError, Excel.Contrib.Interfaces.IIOLEDBError>();
		}

		/// <summary>
		/// Wrapper interface for IOLEDBErrors which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOLEDBErrors WithComCleanup(this Microsoft.Office.Interop.Excel.IOLEDBErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOLEDBErrors, Excel.Contrib.Interfaces.IIOLEDBErrors>();
		}

		/// <summary>
		/// Wrapper interface for IPhonetics which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPhonetics WithComCleanup(this Microsoft.Office.Interop.Excel.IPhonetics resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPhonetics, Excel.Contrib.Interfaces.IIPhonetics>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDefaultWebOptions WithComCleanup(this Microsoft.Office.Interop.Excel.DefaultWebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DefaultWebOptions, Excel.Contrib.Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWebOptions WithComCleanup(this Microsoft.Office.Interop.Excel.WebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.WebOptions, Excel.Contrib.Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLayout which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotLayout WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotLayout resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotLayout, Excel.Contrib.Interfaces.IIPivotLayout>();
		}

		/// <summary>
		/// Wrapper interface for TreeviewControl which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITreeviewControl WithComCleanup(this Microsoft.Office.Interop.Excel.TreeviewControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TreeviewControl, Excel.Contrib.Interfaces.ITreeviewControl>();
		}

		/// <summary>
		/// Wrapper interface for CubeField which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICubeField WithComCleanup(this Microsoft.Office.Interop.Excel.CubeField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CubeField, Excel.Contrib.Interfaces.ICubeField>();
		}

		/// <summary>
		/// Wrapper interface for CubeFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICubeFields WithComCleanup(this Microsoft.Office.Interop.Excel.CubeFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CubeFields, Excel.Contrib.Interfaces.ICubeFields>();
		}

		/// <summary>
		/// Wrapper interface for IDisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDisplayUnitLabel WithComCleanup(this Microsoft.Office.Interop.Excel.IDisplayUnitLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDisplayUnitLabel, Excel.Contrib.Interfaces.IIDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for ICellFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICellFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ICellFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICellFormat, Excel.Contrib.Interfaces.IICellFormat>();
		}

		/// <summary>
		/// Wrapper interface for IUsedObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIUsedObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IUsedObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IUsedObjects, Excel.Contrib.Interfaces.IIUsedObjects>();
		}

		/// <summary>
		/// Wrapper interface for ICustomProperties which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICustomProperties WithComCleanup(this Microsoft.Office.Interop.Excel.ICustomProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICustomProperties, Excel.Contrib.Interfaces.IICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for ICustomProperty which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICustomProperty WithComCleanup(this Microsoft.Office.Interop.Excel.ICustomProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICustomProperty, Excel.Contrib.Interfaces.IICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedMembers which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICalculatedMembers WithComCleanup(this Microsoft.Office.Interop.Excel.ICalculatedMembers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICalculatedMembers, Excel.Contrib.Interfaces.IICalculatedMembers>();
		}

		/// <summary>
		/// Wrapper interface for ICalculatedMember which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICalculatedMember WithComCleanup(this Microsoft.Office.Interop.Excel.ICalculatedMember resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICalculatedMember, Excel.Contrib.Interfaces.IICalculatedMember>();
		}

		/// <summary>
		/// Wrapper interface for IWatches which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWatches WithComCleanup(this Microsoft.Office.Interop.Excel.IWatches resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWatches, Excel.Contrib.Interfaces.IIWatches>();
		}

		/// <summary>
		/// Wrapper interface for IWatch which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWatch WithComCleanup(this Microsoft.Office.Interop.Excel.IWatch resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWatch, Excel.Contrib.Interfaces.IIWatch>();
		}

		/// <summary>
		/// Wrapper interface for IPivotCell which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotCell WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotCell resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotCell, Excel.Contrib.Interfaces.IIPivotCell>();
		}

		/// <summary>
		/// Wrapper interface for IGraphic which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIGraphic WithComCleanup(this Microsoft.Office.Interop.Excel.IGraphic resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IGraphic, Excel.Contrib.Interfaces.IIGraphic>();
		}

		/// <summary>
		/// Wrapper interface for IAutoRecover which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAutoRecover WithComCleanup(this Microsoft.Office.Interop.Excel.IAutoRecover resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAutoRecover, Excel.Contrib.Interfaces.IIAutoRecover>();
		}

		/// <summary>
		/// Wrapper interface for IErrorCheckingOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIErrorCheckingOptions WithComCleanup(this Microsoft.Office.Interop.Excel.IErrorCheckingOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IErrorCheckingOptions, Excel.Contrib.Interfaces.IIErrorCheckingOptions>();
		}

		/// <summary>
		/// Wrapper interface for IErrors which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIErrors WithComCleanup(this Microsoft.Office.Interop.Excel.IErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IErrors, Excel.Contrib.Interfaces.IIErrors>();
		}

		/// <summary>
		/// Wrapper interface for IError which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIError WithComCleanup(this Microsoft.Office.Interop.Excel.IError resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IError, Excel.Contrib.Interfaces.IIError>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTagAction WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTagAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTagAction, Excel.Contrib.Interfaces.IISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTagActions WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTagActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTagActions, Excel.Contrib.Interfaces.IISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTag which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTag WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTag resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTag, Excel.Contrib.Interfaces.IISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTags which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTags WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTags resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTags, Excel.Contrib.Interfaces.IISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTagRecognizer WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTagRecognizer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTagRecognizer, Excel.Contrib.Interfaces.IISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTagRecognizers WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTagRecognizers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTagRecognizers, Excel.Contrib.Interfaces.IISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for ISmartTagOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISmartTagOptions WithComCleanup(this Microsoft.Office.Interop.Excel.ISmartTagOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISmartTagOptions, Excel.Contrib.Interfaces.IISmartTagOptions>();
		}

		/// <summary>
		/// Wrapper interface for ISpellingOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISpellingOptions WithComCleanup(this Microsoft.Office.Interop.Excel.ISpellingOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISpellingOptions, Excel.Contrib.Interfaces.IISpellingOptions>();
		}

		/// <summary>
		/// Wrapper interface for ISpeech which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISpeech WithComCleanup(this Microsoft.Office.Interop.Excel.ISpeech resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISpeech, Excel.Contrib.Interfaces.IISpeech>();
		}

		/// <summary>
		/// Wrapper interface for IProtection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIProtection WithComCleanup(this Microsoft.Office.Interop.Excel.IProtection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IProtection, Excel.Contrib.Interfaces.IIProtection>();
		}

		/// <summary>
		/// Wrapper interface for IPivotItemList which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotItemList WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotItemList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotItemList, Excel.Contrib.Interfaces.IIPivotItemList>();
		}

		/// <summary>
		/// Wrapper interface for ITab which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITab WithComCleanup(this Microsoft.Office.Interop.Excel.ITab resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITab, Excel.Contrib.Interfaces.IITab>();
		}

		/// <summary>
		/// Wrapper interface for IAllowEditRanges which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAllowEditRanges WithComCleanup(this Microsoft.Office.Interop.Excel.IAllowEditRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAllowEditRanges, Excel.Contrib.Interfaces.IIAllowEditRanges>();
		}

		/// <summary>
		/// Wrapper interface for IAllowEditRange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAllowEditRange WithComCleanup(this Microsoft.Office.Interop.Excel.IAllowEditRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAllowEditRange, Excel.Contrib.Interfaces.IIAllowEditRange>();
		}

		/// <summary>
		/// Wrapper interface for IUserAccessList which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIUserAccessList WithComCleanup(this Microsoft.Office.Interop.Excel.IUserAccessList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IUserAccessList, Excel.Contrib.Interfaces.IIUserAccessList>();
		}

		/// <summary>
		/// Wrapper interface for IUserAccess which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIUserAccess WithComCleanup(this Microsoft.Office.Interop.Excel.IUserAccess resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IUserAccess, Excel.Contrib.Interfaces.IIUserAccess>();
		}

		/// <summary>
		/// Wrapper interface for IRTD which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRTD WithComCleanup(this Microsoft.Office.Interop.Excel.IRTD resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRTD, Excel.Contrib.Interfaces.IIRTD>();
		}

		/// <summary>
		/// Wrapper interface for IDiagram which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDiagram WithComCleanup(this Microsoft.Office.Interop.Excel.IDiagram resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDiagram, Excel.Contrib.Interfaces.IIDiagram>();
		}

		/// <summary>
		/// Wrapper interface for IListObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListObjects WithComCleanup(this Microsoft.Office.Interop.Excel.IListObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListObjects, Excel.Contrib.Interfaces.IIListObjects>();
		}

		/// <summary>
		/// Wrapper interface for IListObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListObject WithComCleanup(this Microsoft.Office.Interop.Excel.IListObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListObject, Excel.Contrib.Interfaces.IIListObject>();
		}

		/// <summary>
		/// Wrapper interface for IListColumns which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListColumns WithComCleanup(this Microsoft.Office.Interop.Excel.IListColumns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListColumns, Excel.Contrib.Interfaces.IIListColumns>();
		}

		/// <summary>
		/// Wrapper interface for IListColumn which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListColumn WithComCleanup(this Microsoft.Office.Interop.Excel.IListColumn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListColumn, Excel.Contrib.Interfaces.IIListColumn>();
		}

		/// <summary>
		/// Wrapper interface for IListRows which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListRows WithComCleanup(this Microsoft.Office.Interop.Excel.IListRows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListRows, Excel.Contrib.Interfaces.IIListRows>();
		}

		/// <summary>
		/// Wrapper interface for IListRow which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListRow WithComCleanup(this Microsoft.Office.Interop.Excel.IListRow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListRow, Excel.Contrib.Interfaces.IIListRow>();
		}

		/// <summary>
		/// Wrapper interface for IXmlNamespace which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlNamespace WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlNamespace resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlNamespace, Excel.Contrib.Interfaces.IIXmlNamespace>();
		}

		/// <summary>
		/// Wrapper interface for IXmlNamespaces which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlNamespaces WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlNamespaces resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlNamespaces, Excel.Contrib.Interfaces.IIXmlNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for IXmlDataBinding which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlDataBinding WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlDataBinding resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlDataBinding, Excel.Contrib.Interfaces.IIXmlDataBinding>();
		}

		/// <summary>
		/// Wrapper interface for IXmlSchema which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlSchema WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlSchema resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlSchema, Excel.Contrib.Interfaces.IIXmlSchema>();
		}

		/// <summary>
		/// Wrapper interface for IXmlSchemas which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlSchemas WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlSchemas resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlSchemas, Excel.Contrib.Interfaces.IIXmlSchemas>();
		}

		/// <summary>
		/// Wrapper interface for IXmlMap which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlMap WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlMap resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlMap, Excel.Contrib.Interfaces.IIXmlMap>();
		}

		/// <summary>
		/// Wrapper interface for IXmlMaps which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXmlMaps WithComCleanup(this Microsoft.Office.Interop.Excel.IXmlMaps resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXmlMaps, Excel.Contrib.Interfaces.IIXmlMaps>();
		}

		/// <summary>
		/// Wrapper interface for IListDataFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIListDataFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IListDataFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IListDataFormat, Excel.Contrib.Interfaces.IIListDataFormat>();
		}

		/// <summary>
		/// Wrapper interface for IXPath which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIXPath WithComCleanup(this Microsoft.Office.Interop.Excel.IXPath resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IXPath, Excel.Contrib.Interfaces.IIXPath>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLineCells which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotLineCells WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotLineCells resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotLineCells, Excel.Contrib.Interfaces.IIPivotLineCells>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLine which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotLine WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotLine resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotLine, Excel.Contrib.Interfaces.IIPivotLine>();
		}

		/// <summary>
		/// Wrapper interface for IPivotLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotLines WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotLines, Excel.Contrib.Interfaces.IIPivotLines>();
		}

		/// <summary>
		/// Wrapper interface for IPivotAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotAxis WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotAxis, Excel.Contrib.Interfaces.IIPivotAxis>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFilter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotFilter WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotFilter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotFilter, Excel.Contrib.Interfaces.IIPivotFilter>();
		}

		/// <summary>
		/// Wrapper interface for IPivotFilters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotFilters WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotFilters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotFilters, Excel.Contrib.Interfaces.IIPivotFilters>();
		}

		/// <summary>
		/// Wrapper interface for IWorkbookConnection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWorkbookConnection WithComCleanup(this Microsoft.Office.Interop.Excel.IWorkbookConnection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWorkbookConnection, Excel.Contrib.Interfaces.IIWorkbookConnection>();
		}

		/// <summary>
		/// Wrapper interface for IConnections which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIConnections WithComCleanup(this Microsoft.Office.Interop.Excel.IConnections resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IConnections, Excel.Contrib.Interfaces.IIConnections>();
		}

		/// <summary>
		/// Wrapper interface for IWorksheetView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIWorksheetView WithComCleanup(this Microsoft.Office.Interop.Excel.IWorksheetView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IWorksheetView, Excel.Contrib.Interfaces.IIWorksheetView>();
		}

		/// <summary>
		/// Wrapper interface for IChartView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartView WithComCleanup(this Microsoft.Office.Interop.Excel.IChartView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartView, Excel.Contrib.Interfaces.IIChartView>();
		}

		/// <summary>
		/// Wrapper interface for IModuleView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIModuleView WithComCleanup(this Microsoft.Office.Interop.Excel.IModuleView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IModuleView, Excel.Contrib.Interfaces.IIModuleView>();
		}

		/// <summary>
		/// Wrapper interface for IDialogSheetView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDialogSheetView WithComCleanup(this Microsoft.Office.Interop.Excel.IDialogSheetView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDialogSheetView, Excel.Contrib.Interfaces.IIDialogSheetView>();
		}

		/// <summary>
		/// Wrapper interface for ISheetViews which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISheetViews WithComCleanup(this Microsoft.Office.Interop.Excel.ISheetViews resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISheetViews, Excel.Contrib.Interfaces.IISheetViews>();
		}

		/// <summary>
		/// Wrapper interface for IOLEDBConnection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIOLEDBConnection WithComCleanup(this Microsoft.Office.Interop.Excel.IOLEDBConnection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IOLEDBConnection, Excel.Contrib.Interfaces.IIOLEDBConnection>();
		}

		/// <summary>
		/// Wrapper interface for IODBCConnection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIODBCConnection WithComCleanup(this Microsoft.Office.Interop.Excel.IODBCConnection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IODBCConnection, Excel.Contrib.Interfaces.IIODBCConnection>();
		}

		/// <summary>
		/// Wrapper interface for IAction which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAction WithComCleanup(this Microsoft.Office.Interop.Excel.IAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAction, Excel.Contrib.Interfaces.IIAction>();
		}

		/// <summary>
		/// Wrapper interface for IActions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIActions WithComCleanup(this Microsoft.Office.Interop.Excel.IActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IActions, Excel.Contrib.Interfaces.IIActions>();
		}

		/// <summary>
		/// Wrapper interface for IFormatColor which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFormatColor WithComCleanup(this Microsoft.Office.Interop.Excel.IFormatColor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFormatColor, Excel.Contrib.Interfaces.IIFormatColor>();
		}

		/// <summary>
		/// Wrapper interface for IConditionValue which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIConditionValue WithComCleanup(this Microsoft.Office.Interop.Excel.IConditionValue resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IConditionValue, Excel.Contrib.Interfaces.IIConditionValue>();
		}

		/// <summary>
		/// Wrapper interface for IColorScale which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIColorScale WithComCleanup(this Microsoft.Office.Interop.Excel.IColorScale resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IColorScale, Excel.Contrib.Interfaces.IIColorScale>();
		}

		/// <summary>
		/// Wrapper interface for IColorScaleCriteria which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIColorScaleCriteria WithComCleanup(this Microsoft.Office.Interop.Excel.IColorScaleCriteria resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IColorScaleCriteria, Excel.Contrib.Interfaces.IIColorScaleCriteria>();
		}

		/// <summary>
		/// Wrapper interface for IColorScaleCriterion which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIColorScaleCriterion WithComCleanup(this Microsoft.Office.Interop.Excel.IColorScaleCriterion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IColorScaleCriterion, Excel.Contrib.Interfaces.IIColorScaleCriterion>();
		}

		/// <summary>
		/// Wrapper interface for IDatabar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDatabar WithComCleanup(this Microsoft.Office.Interop.Excel.IDatabar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDatabar, Excel.Contrib.Interfaces.IIDatabar>();
		}

		/// <summary>
		/// Wrapper interface for IIconSetCondition which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIIconSetCondition WithComCleanup(this Microsoft.Office.Interop.Excel.IIconSetCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IIconSetCondition, Excel.Contrib.Interfaces.IIIconSetCondition>();
		}

		/// <summary>
		/// Wrapper interface for IIconCriteria which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIIconCriteria WithComCleanup(this Microsoft.Office.Interop.Excel.IIconCriteria resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IIconCriteria, Excel.Contrib.Interfaces.IIIconCriteria>();
		}

		/// <summary>
		/// Wrapper interface for IIconCriterion which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIIconCriterion WithComCleanup(this Microsoft.Office.Interop.Excel.IIconCriterion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IIconCriterion, Excel.Contrib.Interfaces.IIIconCriterion>();
		}

		/// <summary>
		/// Wrapper interface for IIcon which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIIcon WithComCleanup(this Microsoft.Office.Interop.Excel.IIcon resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IIcon, Excel.Contrib.Interfaces.IIIcon>();
		}

		/// <summary>
		/// Wrapper interface for IIconSet which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIIconSet WithComCleanup(this Microsoft.Office.Interop.Excel.IIconSet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IIconSet, Excel.Contrib.Interfaces.IIIconSet>();
		}

		/// <summary>
		/// Wrapper interface for IIconSets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIIconSets WithComCleanup(this Microsoft.Office.Interop.Excel.IIconSets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IIconSets, Excel.Contrib.Interfaces.IIIconSets>();
		}

		/// <summary>
		/// Wrapper interface for ITop10 which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITop10 WithComCleanup(this Microsoft.Office.Interop.Excel.ITop10 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITop10, Excel.Contrib.Interfaces.IITop10>();
		}

		/// <summary>
		/// Wrapper interface for IAboveAverage which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAboveAverage WithComCleanup(this Microsoft.Office.Interop.Excel.IAboveAverage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAboveAverage, Excel.Contrib.Interfaces.IIAboveAverage>();
		}

		/// <summary>
		/// Wrapper interface for IUniqueValues which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIUniqueValues WithComCleanup(this Microsoft.Office.Interop.Excel.IUniqueValues resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IUniqueValues, Excel.Contrib.Interfaces.IIUniqueValues>();
		}

		/// <summary>
		/// Wrapper interface for IRanges which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRanges WithComCleanup(this Microsoft.Office.Interop.Excel.IRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRanges, Excel.Contrib.Interfaces.IIRanges>();
		}

		/// <summary>
		/// Wrapper interface for IHeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIHeaderFooter WithComCleanup(this Microsoft.Office.Interop.Excel.IHeaderFooter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IHeaderFooter, Excel.Contrib.Interfaces.IIHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for IPage which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPage WithComCleanup(this Microsoft.Office.Interop.Excel.IPage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPage, Excel.Contrib.Interfaces.IIPage>();
		}

		/// <summary>
		/// Wrapper interface for IPages which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPages WithComCleanup(this Microsoft.Office.Interop.Excel.IPages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPages, Excel.Contrib.Interfaces.IIPages>();
		}

		/// <summary>
		/// Wrapper interface for IServerViewableItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIServerViewableItems WithComCleanup(this Microsoft.Office.Interop.Excel.IServerViewableItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IServerViewableItems, Excel.Contrib.Interfaces.IIServerViewableItems>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyleElement which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITableStyleElement WithComCleanup(this Microsoft.Office.Interop.Excel.ITableStyleElement resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITableStyleElement, Excel.Contrib.Interfaces.IITableStyleElement>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyleElements which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITableStyleElements WithComCleanup(this Microsoft.Office.Interop.Excel.ITableStyleElements resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITableStyleElements, Excel.Contrib.Interfaces.IITableStyleElements>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITableStyle WithComCleanup(this Microsoft.Office.Interop.Excel.ITableStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITableStyle, Excel.Contrib.Interfaces.IITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for ITableStyles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IITableStyles WithComCleanup(this Microsoft.Office.Interop.Excel.ITableStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ITableStyles, Excel.Contrib.Interfaces.IITableStyles>();
		}

		/// <summary>
		/// Wrapper interface for ISortField which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISortField WithComCleanup(this Microsoft.Office.Interop.Excel.ISortField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISortField, Excel.Contrib.Interfaces.IISortField>();
		}

		/// <summary>
		/// Wrapper interface for ISortFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISortFields WithComCleanup(this Microsoft.Office.Interop.Excel.ISortFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISortFields, Excel.Contrib.Interfaces.IISortFields>();
		}

		/// <summary>
		/// Wrapper interface for ISort which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISort WithComCleanup(this Microsoft.Office.Interop.Excel.ISort resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISort, Excel.Contrib.Interfaces.IISort>();
		}

		/// <summary>
		/// Wrapper interface for IResearch which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIResearch WithComCleanup(this Microsoft.Office.Interop.Excel.IResearch resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IResearch, Excel.Contrib.Interfaces.IIResearch>();
		}

		/// <summary>
		/// Wrapper interface for IColorStop which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIColorStop WithComCleanup(this Microsoft.Office.Interop.Excel.IColorStop resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IColorStop, Excel.Contrib.Interfaces.IIColorStop>();
		}

		/// <summary>
		/// Wrapper interface for IColorStops which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIColorStops WithComCleanup(this Microsoft.Office.Interop.Excel.IColorStops resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IColorStops, Excel.Contrib.Interfaces.IIColorStops>();
		}

		/// <summary>
		/// Wrapper interface for ILinearGradient which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IILinearGradient WithComCleanup(this Microsoft.Office.Interop.Excel.ILinearGradient resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ILinearGradient, Excel.Contrib.Interfaces.IILinearGradient>();
		}

		/// <summary>
		/// Wrapper interface for IRectangularGradient which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIRectangularGradient WithComCleanup(this Microsoft.Office.Interop.Excel.IRectangularGradient resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IRectangularGradient, Excel.Contrib.Interfaces.IIRectangularGradient>();
		}

		/// <summary>
		/// Wrapper interface for IMultiThreadedCalculation which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIMultiThreadedCalculation WithComCleanup(this Microsoft.Office.Interop.Excel.IMultiThreadedCalculation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IMultiThreadedCalculation, Excel.Contrib.Interfaces.IIMultiThreadedCalculation>();
		}

		/// <summary>
		/// Wrapper interface for IChartFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIChartFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IChartFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IChartFormat, Excel.Contrib.Interfaces.IIChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for IFileExportConverter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFileExportConverter WithComCleanup(this Microsoft.Office.Interop.Excel.IFileExportConverter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFileExportConverter, Excel.Contrib.Interfaces.IIFileExportConverter>();
		}

		/// <summary>
		/// Wrapper interface for IFileExportConverters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIFileExportConverters WithComCleanup(this Microsoft.Office.Interop.Excel.IFileExportConverters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IFileExportConverters, Excel.Contrib.Interfaces.IIFileExportConverters>();
		}

		/// <summary>
		/// Wrapper interface for IAddIns2 which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIAddIns2 WithComCleanup(this Microsoft.Office.Interop.Excel.IAddIns2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IAddIns2, Excel.Contrib.Interfaces.IIAddIns2>();
		}

		/// <summary>
		/// Wrapper interface for ISparklineGroups which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparklineGroups WithComCleanup(this Microsoft.Office.Interop.Excel.ISparklineGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparklineGroups, Excel.Contrib.Interfaces.IISparklineGroups>();
		}

		/// <summary>
		/// Wrapper interface for ISparklineGroup which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparklineGroup WithComCleanup(this Microsoft.Office.Interop.Excel.ISparklineGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparklineGroup, Excel.Contrib.Interfaces.IISparklineGroup>();
		}

		/// <summary>
		/// Wrapper interface for ISparkPoints which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparkPoints WithComCleanup(this Microsoft.Office.Interop.Excel.ISparkPoints resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparkPoints, Excel.Contrib.Interfaces.IISparkPoints>();
		}

		/// <summary>
		/// Wrapper interface for ISparkline which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparkline WithComCleanup(this Microsoft.Office.Interop.Excel.ISparkline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparkline, Excel.Contrib.Interfaces.IISparkline>();
		}

		/// <summary>
		/// Wrapper interface for ISparkAxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparkAxes WithComCleanup(this Microsoft.Office.Interop.Excel.ISparkAxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparkAxes, Excel.Contrib.Interfaces.IISparkAxes>();
		}

		/// <summary>
		/// Wrapper interface for ISparkHorizontalAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparkHorizontalAxis WithComCleanup(this Microsoft.Office.Interop.Excel.ISparkHorizontalAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparkHorizontalAxis, Excel.Contrib.Interfaces.IISparkHorizontalAxis>();
		}

		/// <summary>
		/// Wrapper interface for ISparkVerticalAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparkVerticalAxis WithComCleanup(this Microsoft.Office.Interop.Excel.ISparkVerticalAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparkVerticalAxis, Excel.Contrib.Interfaces.IISparkVerticalAxis>();
		}

		/// <summary>
		/// Wrapper interface for ISparkColor which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISparkColor WithComCleanup(this Microsoft.Office.Interop.Excel.ISparkColor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISparkColor, Excel.Contrib.Interfaces.IISparkColor>();
		}

		/// <summary>
		/// Wrapper interface for IDataBarBorder which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDataBarBorder WithComCleanup(this Microsoft.Office.Interop.Excel.IDataBarBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDataBarBorder, Excel.Contrib.Interfaces.IIDataBarBorder>();
		}

		/// <summary>
		/// Wrapper interface for INegativeBarFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IINegativeBarFormat WithComCleanup(this Microsoft.Office.Interop.Excel.INegativeBarFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.INegativeBarFormat, Excel.Contrib.Interfaces.IINegativeBarFormat>();
		}

		/// <summary>
		/// Wrapper interface for IValueChange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIValueChange WithComCleanup(this Microsoft.Office.Interop.Excel.IValueChange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IValueChange, Excel.Contrib.Interfaces.IIValueChange>();
		}

		/// <summary>
		/// Wrapper interface for IPivotTableChangeList which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIPivotTableChangeList WithComCleanup(this Microsoft.Office.Interop.Excel.IPivotTableChangeList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IPivotTableChangeList, Excel.Contrib.Interfaces.IIPivotTableChangeList>();
		}

		/// <summary>
		/// Wrapper interface for IDisplayFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDisplayFormat WithComCleanup(this Microsoft.Office.Interop.Excel.IDisplayFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDisplayFormat, Excel.Contrib.Interfaces.IIDisplayFormat>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCaches which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerCaches WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerCaches resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerCaches, Excel.Contrib.Interfaces.IISlicerCaches>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCache which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerCache WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerCache resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerCache, Excel.Contrib.Interfaces.IISlicerCache>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCacheLevels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerCacheLevels WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerCacheLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerCacheLevels, Excel.Contrib.Interfaces.IISlicerCacheLevels>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerCacheLevel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerCacheLevel WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerCacheLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerCacheLevel, Excel.Contrib.Interfaces.IISlicerCacheLevel>();
		}

		/// <summary>
		/// Wrapper interface for ISlicers which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicers WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicers, Excel.Contrib.Interfaces.IISlicers>();
		}

		/// <summary>
		/// Wrapper interface for ISlicer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicer WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicer, Excel.Contrib.Interfaces.IISlicer>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerItem which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerItem WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerItem, Excel.Contrib.Interfaces.IISlicerItem>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerItems WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerItems, Excel.Contrib.Interfaces.IISlicerItems>();
		}

		/// <summary>
		/// Wrapper interface for ISlicerPivotTables which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IISlicerPivotTables WithComCleanup(this Microsoft.Office.Interop.Excel.ISlicerPivotTables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ISlicerPivotTables, Excel.Contrib.Interfaces.IISlicerPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for IProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIProtectedViewWindows WithComCleanup(this Microsoft.Office.Interop.Excel.IProtectedViewWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IProtectedViewWindows, Excel.Contrib.Interfaces.IIProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for IProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIProtectedViewWindow WithComCleanup(this Microsoft.Office.Interop.Excel.IProtectedViewWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IProtectedViewWindow, Excel.Contrib.Interfaces.IIProtectedViewWindow>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFont WithComCleanup(this Microsoft.Office.Interop.Excel.Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Font, Excel.Contrib.Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for Window which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWindow WithComCleanup(this Microsoft.Office.Interop.Excel.Window resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Window, Excel.Contrib.Interfaces.IWindow>();
		}

		/// <summary>
		/// Wrapper interface for Windows which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWindows WithComCleanup(this Microsoft.Office.Interop.Excel.Windows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Windows, Excel.Contrib.Interfaces.IWindows>();
		}

		/// <summary>
		/// Wrapper interface for AppEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAppEvents WithComCleanup(this Microsoft.Office.Interop.Excel.AppEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AppEvents, Excel.Contrib.Interfaces.IAppEvents>();
		}

		/// <summary>
		/// Wrapper interface for WorksheetFunction which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorksheetFunction WithComCleanup(this Microsoft.Office.Interop.Excel.WorksheetFunction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.WorksheetFunction, Excel.Contrib.Interfaces.IWorksheetFunction>();
		}

		/// <summary>
		/// Wrapper interface for Range which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRange WithComCleanup(this Microsoft.Office.Interop.Excel.Range resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Range, Excel.Contrib.Interfaces.IRange>();
		}

		/// <summary>
		/// Wrapper interface for ChartEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartEvents WithComCleanup(this Microsoft.Office.Interop.Excel.ChartEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartEvents, Excel.Contrib.Interfaces.IChartEvents>();
		}

		/// <summary>
		/// Wrapper interface for VPageBreak which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IVPageBreak WithComCleanup(this Microsoft.Office.Interop.Excel.VPageBreak resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.VPageBreak, Excel.Contrib.Interfaces.IVPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for HPageBreak which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IHPageBreak WithComCleanup(this Microsoft.Office.Interop.Excel.HPageBreak resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.HPageBreak, Excel.Contrib.Interfaces.IHPageBreak>();
		}

		/// <summary>
		/// Wrapper interface for HPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IHPageBreaks WithComCleanup(this Microsoft.Office.Interop.Excel.HPageBreaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.HPageBreaks, Excel.Contrib.Interfaces.IHPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for VPageBreaks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IVPageBreaks WithComCleanup(this Microsoft.Office.Interop.Excel.VPageBreaks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.VPageBreaks, Excel.Contrib.Interfaces.IVPageBreaks>();
		}

		/// <summary>
		/// Wrapper interface for RecentFile which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRecentFile WithComCleanup(this Microsoft.Office.Interop.Excel.RecentFile resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RecentFile, Excel.Contrib.Interfaces.IRecentFile>();
		}

		/// <summary>
		/// Wrapper interface for RecentFiles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRecentFiles WithComCleanup(this Microsoft.Office.Interop.Excel.RecentFiles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RecentFiles, Excel.Contrib.Interfaces.IRecentFiles>();
		}

		/// <summary>
		/// Wrapper interface for DocEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDocEvents WithComCleanup(this Microsoft.Office.Interop.Excel.DocEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DocEvents, Excel.Contrib.Interfaces.IDocEvents>();
		}

		/// <summary>
		/// Wrapper interface for Style which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IStyle WithComCleanup(this Microsoft.Office.Interop.Excel.Style resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Style, Excel.Contrib.Interfaces.IStyle>();
		}

		/// <summary>
		/// Wrapper interface for Styles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IStyles WithComCleanup(this Microsoft.Office.Interop.Excel.Styles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Styles, Excel.Contrib.Interfaces.IStyles>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IBorders WithComCleanup(this Microsoft.Office.Interop.Excel.Borders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Borders, Excel.Contrib.Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAddIn WithComCleanup(this Microsoft.Office.Interop.Excel.AddIn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AddIn, Excel.Contrib.Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAddIns WithComCleanup(this Microsoft.Office.Interop.Excel.AddIns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AddIns, Excel.Contrib.Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for Toolbar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IToolbar WithComCleanup(this Microsoft.Office.Interop.Excel.Toolbar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Toolbar, Excel.Contrib.Interfaces.IToolbar>();
		}

		/// <summary>
		/// Wrapper interface for Toolbars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IToolbars WithComCleanup(this Microsoft.Office.Interop.Excel.Toolbars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Toolbars, Excel.Contrib.Interfaces.IToolbars>();
		}

		/// <summary>
		/// Wrapper interface for ToolbarButton which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IToolbarButton WithComCleanup(this Microsoft.Office.Interop.Excel.ToolbarButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ToolbarButton, Excel.Contrib.Interfaces.IToolbarButton>();
		}

		/// <summary>
		/// Wrapper interface for ToolbarButtons which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IToolbarButtons WithComCleanup(this Microsoft.Office.Interop.Excel.ToolbarButtons resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ToolbarButtons, Excel.Contrib.Interfaces.IToolbarButtons>();
		}

		/// <summary>
		/// Wrapper interface for Areas which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAreas WithComCleanup(this Microsoft.Office.Interop.Excel.Areas resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Areas, Excel.Contrib.Interfaces.IAreas>();
		}

		/// <summary>
		/// Wrapper interface for WorkbookEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorkbookEvents WithComCleanup(this Microsoft.Office.Interop.Excel.WorkbookEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.WorkbookEvents, Excel.Contrib.Interfaces.IWorkbookEvents>();
		}

		/// <summary>
		/// Wrapper interface for MenuBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMenuBars WithComCleanup(this Microsoft.Office.Interop.Excel.MenuBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.MenuBars, Excel.Contrib.Interfaces.IMenuBars>();
		}

		/// <summary>
		/// Wrapper interface for MenuBar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMenuBar WithComCleanup(this Microsoft.Office.Interop.Excel.MenuBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.MenuBar, Excel.Contrib.Interfaces.IMenuBar>();
		}

		/// <summary>
		/// Wrapper interface for Menus which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMenus WithComCleanup(this Microsoft.Office.Interop.Excel.Menus resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Menus, Excel.Contrib.Interfaces.IMenus>();
		}

		/// <summary>
		/// Wrapper interface for Menu which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMenu WithComCleanup(this Microsoft.Office.Interop.Excel.Menu resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Menu, Excel.Contrib.Interfaces.IMenu>();
		}

		/// <summary>
		/// Wrapper interface for MenuItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMenuItems WithComCleanup(this Microsoft.Office.Interop.Excel.MenuItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.MenuItems, Excel.Contrib.Interfaces.IMenuItems>();
		}

		/// <summary>
		/// Wrapper interface for MenuItem which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMenuItem WithComCleanup(this Microsoft.Office.Interop.Excel.MenuItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.MenuItem, Excel.Contrib.Interfaces.IMenuItem>();
		}

		/// <summary>
		/// Wrapper interface for Charts which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICharts WithComCleanup(this Microsoft.Office.Interop.Excel.Charts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Charts, Excel.Contrib.Interfaces.ICharts>();
		}

		/// <summary>
		/// Wrapper interface for DrawingObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDrawingObjects WithComCleanup(this Microsoft.Office.Interop.Excel.DrawingObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DrawingObjects, Excel.Contrib.Interfaces.IDrawingObjects>();
		}

		/// <summary>
		/// Wrapper interface for PivotCache which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotCache WithComCleanup(this Microsoft.Office.Interop.Excel.PivotCache resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotCache, Excel.Contrib.Interfaces.IPivotCache>();
		}

		/// <summary>
		/// Wrapper interface for PivotCaches which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotCaches WithComCleanup(this Microsoft.Office.Interop.Excel.PivotCaches resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotCaches, Excel.Contrib.Interfaces.IPivotCaches>();
		}

		/// <summary>
		/// Wrapper interface for PivotFormula which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotFormula WithComCleanup(this Microsoft.Office.Interop.Excel.PivotFormula resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotFormula, Excel.Contrib.Interfaces.IPivotFormula>();
		}

		/// <summary>
		/// Wrapper interface for PivotFormulas which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotFormulas WithComCleanup(this Microsoft.Office.Interop.Excel.PivotFormulas resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotFormulas, Excel.Contrib.Interfaces.IPivotFormulas>();
		}

		/// <summary>
		/// Wrapper interface for PivotTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotTable WithComCleanup(this Microsoft.Office.Interop.Excel.PivotTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotTable, Excel.Contrib.Interfaces.IPivotTable>();
		}

		/// <summary>
		/// Wrapper interface for PivotTables which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotTables WithComCleanup(this Microsoft.Office.Interop.Excel.PivotTables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotTables, Excel.Contrib.Interfaces.IPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for PivotField which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotField WithComCleanup(this Microsoft.Office.Interop.Excel.PivotField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotField, Excel.Contrib.Interfaces.IPivotField>();
		}

		/// <summary>
		/// Wrapper interface for PivotFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotFields WithComCleanup(this Microsoft.Office.Interop.Excel.PivotFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotFields, Excel.Contrib.Interfaces.IPivotFields>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICalculatedFields WithComCleanup(this Microsoft.Office.Interop.Excel.CalculatedFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CalculatedFields, Excel.Contrib.Interfaces.ICalculatedFields>();
		}

		/// <summary>
		/// Wrapper interface for PivotItem which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotItem WithComCleanup(this Microsoft.Office.Interop.Excel.PivotItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotItem, Excel.Contrib.Interfaces.IPivotItem>();
		}

		/// <summary>
		/// Wrapper interface for PivotItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotItems WithComCleanup(this Microsoft.Office.Interop.Excel.PivotItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotItems, Excel.Contrib.Interfaces.IPivotItems>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICalculatedItems WithComCleanup(this Microsoft.Office.Interop.Excel.CalculatedItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CalculatedItems, Excel.Contrib.Interfaces.ICalculatedItems>();
		}

		/// <summary>
		/// Wrapper interface for Characters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICharacters WithComCleanup(this Microsoft.Office.Interop.Excel.Characters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Characters, Excel.Contrib.Interfaces.ICharacters>();
		}

		/// <summary>
		/// Wrapper interface for Dialogs which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDialogs WithComCleanup(this Microsoft.Office.Interop.Excel.Dialogs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Dialogs, Excel.Contrib.Interfaces.IDialogs>();
		}

		/// <summary>
		/// Wrapper interface for Dialog which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDialog WithComCleanup(this Microsoft.Office.Interop.Excel.Dialog resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Dialog, Excel.Contrib.Interfaces.IDialog>();
		}

		/// <summary>
		/// Wrapper interface for SoundNote which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISoundNote WithComCleanup(this Microsoft.Office.Interop.Excel.SoundNote resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SoundNote, Excel.Contrib.Interfaces.ISoundNote>();
		}

		/// <summary>
		/// Wrapper interface for Button which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IButton WithComCleanup(this Microsoft.Office.Interop.Excel.Button resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Button, Excel.Contrib.Interfaces.IButton>();
		}

		/// <summary>
		/// Wrapper interface for Buttons which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IButtons WithComCleanup(this Microsoft.Office.Interop.Excel.Buttons resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Buttons, Excel.Contrib.Interfaces.IButtons>();
		}

		/// <summary>
		/// Wrapper interface for CheckBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICheckBox WithComCleanup(this Microsoft.Office.Interop.Excel.CheckBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CheckBox, Excel.Contrib.Interfaces.ICheckBox>();
		}

		/// <summary>
		/// Wrapper interface for CheckBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICheckBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.CheckBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CheckBoxes, Excel.Contrib.Interfaces.ICheckBoxes>();
		}

		/// <summary>
		/// Wrapper interface for OptionButton which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOptionButton WithComCleanup(this Microsoft.Office.Interop.Excel.OptionButton resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OptionButton, Excel.Contrib.Interfaces.IOptionButton>();
		}

		/// <summary>
		/// Wrapper interface for OptionButtons which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOptionButtons WithComCleanup(this Microsoft.Office.Interop.Excel.OptionButtons resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OptionButtons, Excel.Contrib.Interfaces.IOptionButtons>();
		}

		/// <summary>
		/// Wrapper interface for EditBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IEditBox WithComCleanup(this Microsoft.Office.Interop.Excel.EditBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.EditBox, Excel.Contrib.Interfaces.IEditBox>();
		}

		/// <summary>
		/// Wrapper interface for EditBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IEditBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.EditBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.EditBoxes, Excel.Contrib.Interfaces.IEditBoxes>();
		}

		/// <summary>
		/// Wrapper interface for ScrollBar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IScrollBar WithComCleanup(this Microsoft.Office.Interop.Excel.ScrollBar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ScrollBar, Excel.Contrib.Interfaces.IScrollBar>();
		}

		/// <summary>
		/// Wrapper interface for ScrollBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IScrollBars WithComCleanup(this Microsoft.Office.Interop.Excel.ScrollBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ScrollBars, Excel.Contrib.Interfaces.IScrollBars>();
		}

		/// <summary>
		/// Wrapper interface for ListBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListBox WithComCleanup(this Microsoft.Office.Interop.Excel.ListBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListBox, Excel.Contrib.Interfaces.IListBox>();
		}

		/// <summary>
		/// Wrapper interface for ListBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.ListBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListBoxes, Excel.Contrib.Interfaces.IListBoxes>();
		}

		/// <summary>
		/// Wrapper interface for GroupBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGroupBox WithComCleanup(this Microsoft.Office.Interop.Excel.GroupBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.GroupBox, Excel.Contrib.Interfaces.IGroupBox>();
		}

		/// <summary>
		/// Wrapper interface for GroupBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGroupBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.GroupBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.GroupBoxes, Excel.Contrib.Interfaces.IGroupBoxes>();
		}

		/// <summary>
		/// Wrapper interface for DropDown which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDropDown WithComCleanup(this Microsoft.Office.Interop.Excel.DropDown resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DropDown, Excel.Contrib.Interfaces.IDropDown>();
		}

		/// <summary>
		/// Wrapper interface for DropDowns which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDropDowns WithComCleanup(this Microsoft.Office.Interop.Excel.DropDowns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DropDowns, Excel.Contrib.Interfaces.IDropDowns>();
		}

		/// <summary>
		/// Wrapper interface for Spinner which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISpinner WithComCleanup(this Microsoft.Office.Interop.Excel.Spinner resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Spinner, Excel.Contrib.Interfaces.ISpinner>();
		}

		/// <summary>
		/// Wrapper interface for Spinners which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISpinners WithComCleanup(this Microsoft.Office.Interop.Excel.Spinners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Spinners, Excel.Contrib.Interfaces.ISpinners>();
		}

		/// <summary>
		/// Wrapper interface for DialogFrame which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDialogFrame WithComCleanup(this Microsoft.Office.Interop.Excel.DialogFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DialogFrame, Excel.Contrib.Interfaces.IDialogFrame>();
		}

		/// <summary>
		/// Wrapper interface for Label which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILabel WithComCleanup(this Microsoft.Office.Interop.Excel.Label resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Label, Excel.Contrib.Interfaces.ILabel>();
		}

		/// <summary>
		/// Wrapper interface for Labels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILabels WithComCleanup(this Microsoft.Office.Interop.Excel.Labels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Labels, Excel.Contrib.Interfaces.ILabels>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPanes WithComCleanup(this Microsoft.Office.Interop.Excel.Panes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Panes, Excel.Contrib.Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPane WithComCleanup(this Microsoft.Office.Interop.Excel.Pane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Pane, Excel.Contrib.Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for Scenarios which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IScenarios WithComCleanup(this Microsoft.Office.Interop.Excel.Scenarios resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Scenarios, Excel.Contrib.Interfaces.IScenarios>();
		}

		/// <summary>
		/// Wrapper interface for Scenario which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IScenario WithComCleanup(this Microsoft.Office.Interop.Excel.Scenario resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Scenario, Excel.Contrib.Interfaces.IScenario>();
		}

		/// <summary>
		/// Wrapper interface for GroupObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGroupObject WithComCleanup(this Microsoft.Office.Interop.Excel.GroupObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.GroupObject, Excel.Contrib.Interfaces.IGroupObject>();
		}

		/// <summary>
		/// Wrapper interface for GroupObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGroupObjects WithComCleanup(this Microsoft.Office.Interop.Excel.GroupObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.GroupObjects, Excel.Contrib.Interfaces.IGroupObjects>();
		}

		/// <summary>
		/// Wrapper interface for Line which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILine WithComCleanup(this Microsoft.Office.Interop.Excel.Line resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Line, Excel.Contrib.Interfaces.ILine>();
		}

		/// <summary>
		/// Wrapper interface for Lines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILines WithComCleanup(this Microsoft.Office.Interop.Excel.Lines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Lines, Excel.Contrib.Interfaces.ILines>();
		}

		/// <summary>
		/// Wrapper interface for Rectangle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRectangle WithComCleanup(this Microsoft.Office.Interop.Excel.Rectangle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Rectangle, Excel.Contrib.Interfaces.IRectangle>();
		}

		/// <summary>
		/// Wrapper interface for Rectangles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRectangles WithComCleanup(this Microsoft.Office.Interop.Excel.Rectangles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Rectangles, Excel.Contrib.Interfaces.IRectangles>();
		}

		/// <summary>
		/// Wrapper interface for Oval which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOval WithComCleanup(this Microsoft.Office.Interop.Excel.Oval resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Oval, Excel.Contrib.Interfaces.IOval>();
		}

		/// <summary>
		/// Wrapper interface for Ovals which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOvals WithComCleanup(this Microsoft.Office.Interop.Excel.Ovals resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Ovals, Excel.Contrib.Interfaces.IOvals>();
		}

		/// <summary>
		/// Wrapper interface for Arc which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IArc WithComCleanup(this Microsoft.Office.Interop.Excel.Arc resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Arc, Excel.Contrib.Interfaces.IArc>();
		}

		/// <summary>
		/// Wrapper interface for Arcs which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IArcs WithComCleanup(this Microsoft.Office.Interop.Excel.Arcs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Arcs, Excel.Contrib.Interfaces.IArcs>();
		}

		/// <summary>
		/// Wrapper interface for OLEObjectEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEObjectEvents WithComCleanup(this Microsoft.Office.Interop.Excel.OLEObjectEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEObjectEvents, Excel.Contrib.Interfaces.IOLEObjectEvents>();
		}

		/// <summary>
		/// Wrapper interface for _OLEObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_OLEObject WithComCleanup(this Microsoft.Office.Interop.Excel._OLEObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._OLEObject, Excel.Contrib.Interfaces.I_OLEObject>();
		}

		/// <summary>
		/// Wrapper interface for OLEObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEObjects WithComCleanup(this Microsoft.Office.Interop.Excel.OLEObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEObjects, Excel.Contrib.Interfaces.IOLEObjects>();
		}

		/// <summary>
		/// Wrapper interface for TextBox which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITextBox WithComCleanup(this Microsoft.Office.Interop.Excel.TextBox resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TextBox, Excel.Contrib.Interfaces.ITextBox>();
		}

		/// <summary>
		/// Wrapper interface for TextBoxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITextBoxes WithComCleanup(this Microsoft.Office.Interop.Excel.TextBoxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TextBoxes, Excel.Contrib.Interfaces.ITextBoxes>();
		}

		/// <summary>
		/// Wrapper interface for Picture which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPicture WithComCleanup(this Microsoft.Office.Interop.Excel.Picture resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Picture, Excel.Contrib.Interfaces.IPicture>();
		}

		/// <summary>
		/// Wrapper interface for Pictures which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPictures WithComCleanup(this Microsoft.Office.Interop.Excel.Pictures resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Pictures, Excel.Contrib.Interfaces.IPictures>();
		}

		/// <summary>
		/// Wrapper interface for Drawing which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDrawing WithComCleanup(this Microsoft.Office.Interop.Excel.Drawing resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Drawing, Excel.Contrib.Interfaces.IDrawing>();
		}

		/// <summary>
		/// Wrapper interface for Drawings which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDrawings WithComCleanup(this Microsoft.Office.Interop.Excel.Drawings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Drawings, Excel.Contrib.Interfaces.IDrawings>();
		}

		/// <summary>
		/// Wrapper interface for RoutingSlip which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRoutingSlip WithComCleanup(this Microsoft.Office.Interop.Excel.RoutingSlip resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RoutingSlip, Excel.Contrib.Interfaces.IRoutingSlip>();
		}

		/// <summary>
		/// Wrapper interface for Outline which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOutline WithComCleanup(this Microsoft.Office.Interop.Excel.Outline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Outline, Excel.Contrib.Interfaces.IOutline>();
		}

		/// <summary>
		/// Wrapper interface for Module which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IModule WithComCleanup(this Microsoft.Office.Interop.Excel.Module resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Module, Excel.Contrib.Interfaces.IModule>();
		}

		/// <summary>
		/// Wrapper interface for Modules which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IModules WithComCleanup(this Microsoft.Office.Interop.Excel.Modules resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Modules, Excel.Contrib.Interfaces.IModules>();
		}

		/// <summary>
		/// Wrapper interface for DialogSheet which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDialogSheet WithComCleanup(this Microsoft.Office.Interop.Excel.DialogSheet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DialogSheet, Excel.Contrib.Interfaces.IDialogSheet>();
		}

		/// <summary>
		/// Wrapper interface for DialogSheets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDialogSheets WithComCleanup(this Microsoft.Office.Interop.Excel.DialogSheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DialogSheets, Excel.Contrib.Interfaces.IDialogSheets>();
		}

		/// <summary>
		/// Wrapper interface for Worksheets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorksheets WithComCleanup(this Microsoft.Office.Interop.Excel.Worksheets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Worksheets, Excel.Contrib.Interfaces.IWorksheets>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPageSetup WithComCleanup(this Microsoft.Office.Interop.Excel.PageSetup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PageSetup, Excel.Contrib.Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for Names which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.INames WithComCleanup(this Microsoft.Office.Interop.Excel.Names resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Names, Excel.Contrib.Interfaces.INames>();
		}

		/// <summary>
		/// Wrapper interface for Name which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IName WithComCleanup(this Microsoft.Office.Interop.Excel.Name resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Name, Excel.Contrib.Interfaces.IName>();
		}

		/// <summary>
		/// Wrapper interface for ChartObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartObject WithComCleanup(this Microsoft.Office.Interop.Excel.ChartObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartObject, Excel.Contrib.Interfaces.IChartObject>();
		}

		/// <summary>
		/// Wrapper interface for ChartObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartObjects WithComCleanup(this Microsoft.Office.Interop.Excel.ChartObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartObjects, Excel.Contrib.Interfaces.IChartObjects>();
		}

		/// <summary>
		/// Wrapper interface for Mailer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMailer WithComCleanup(this Microsoft.Office.Interop.Excel.Mailer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Mailer, Excel.Contrib.Interfaces.IMailer>();
		}

		/// <summary>
		/// Wrapper interface for CustomViews which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICustomViews WithComCleanup(this Microsoft.Office.Interop.Excel.CustomViews resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CustomViews, Excel.Contrib.Interfaces.ICustomViews>();
		}

		/// <summary>
		/// Wrapper interface for CustomView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICustomView WithComCleanup(this Microsoft.Office.Interop.Excel.CustomView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CustomView, Excel.Contrib.Interfaces.ICustomView>();
		}

		/// <summary>
		/// Wrapper interface for FormatConditions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFormatConditions WithComCleanup(this Microsoft.Office.Interop.Excel.FormatConditions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FormatConditions, Excel.Contrib.Interfaces.IFormatConditions>();
		}

		/// <summary>
		/// Wrapper interface for FormatCondition which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFormatCondition WithComCleanup(this Microsoft.Office.Interop.Excel.FormatCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FormatCondition, Excel.Contrib.Interfaces.IFormatCondition>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IComments WithComCleanup(this Microsoft.Office.Interop.Excel.Comments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Comments, Excel.Contrib.Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IComment WithComCleanup(this Microsoft.Office.Interop.Excel.Comment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Comment, Excel.Contrib.Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for RefreshEvents which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRefreshEvents WithComCleanup(this Microsoft.Office.Interop.Excel.RefreshEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RefreshEvents, Excel.Contrib.Interfaces.IRefreshEvents>();
		}

		/// <summary>
		/// Wrapper interface for _QueryTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.I_QueryTable WithComCleanup(this Microsoft.Office.Interop.Excel._QueryTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel._QueryTable, Excel.Contrib.Interfaces.I_QueryTable>();
		}

		/// <summary>
		/// Wrapper interface for QueryTables which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IQueryTables WithComCleanup(this Microsoft.Office.Interop.Excel.QueryTables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.QueryTables, Excel.Contrib.Interfaces.IQueryTables>();
		}

		/// <summary>
		/// Wrapper interface for Parameter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IParameter WithComCleanup(this Microsoft.Office.Interop.Excel.Parameter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Parameter, Excel.Contrib.Interfaces.IParameter>();
		}

		/// <summary>
		/// Wrapper interface for Parameters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IParameters WithComCleanup(this Microsoft.Office.Interop.Excel.Parameters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Parameters, Excel.Contrib.Interfaces.IParameters>();
		}

		/// <summary>
		/// Wrapper interface for ODBCError which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IODBCError WithComCleanup(this Microsoft.Office.Interop.Excel.ODBCError resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ODBCError, Excel.Contrib.Interfaces.IODBCError>();
		}

		/// <summary>
		/// Wrapper interface for ODBCErrors which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IODBCErrors WithComCleanup(this Microsoft.Office.Interop.Excel.ODBCErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ODBCErrors, Excel.Contrib.Interfaces.IODBCErrors>();
		}

		/// <summary>
		/// Wrapper interface for Validation which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IValidation WithComCleanup(this Microsoft.Office.Interop.Excel.Validation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Validation, Excel.Contrib.Interfaces.IValidation>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IHyperlinks WithComCleanup(this Microsoft.Office.Interop.Excel.Hyperlinks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Hyperlinks, Excel.Contrib.Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IHyperlink WithComCleanup(this Microsoft.Office.Interop.Excel.Hyperlink resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Hyperlink, Excel.Contrib.Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for AutoFilter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAutoFilter WithComCleanup(this Microsoft.Office.Interop.Excel.AutoFilter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AutoFilter, Excel.Contrib.Interfaces.IAutoFilter>();
		}

		/// <summary>
		/// Wrapper interface for Filters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFilters WithComCleanup(this Microsoft.Office.Interop.Excel.Filters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Filters, Excel.Contrib.Interfaces.IFilters>();
		}

		/// <summary>
		/// Wrapper interface for Filter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFilter WithComCleanup(this Microsoft.Office.Interop.Excel.Filter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Filter, Excel.Contrib.Interfaces.IFilter>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAutoCorrect WithComCleanup(this Microsoft.Office.Interop.Excel.AutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AutoCorrect, Excel.Contrib.Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for Border which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IBorder WithComCleanup(this Microsoft.Office.Interop.Excel.Border resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Border, Excel.Contrib.Interfaces.IBorder>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IInterior WithComCleanup(this Microsoft.Office.Interop.Excel.Interior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Interior, Excel.Contrib.Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartFillFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ChartFillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartFillFormat, Excel.Contrib.Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartColorFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ChartColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartColorFormat, Excel.Contrib.Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAxis WithComCleanup(this Microsoft.Office.Interop.Excel.Axis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Axis, Excel.Contrib.Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartTitle WithComCleanup(this Microsoft.Office.Interop.Excel.ChartTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartTitle, Excel.Contrib.Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAxisTitle WithComCleanup(this Microsoft.Office.Interop.Excel.AxisTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AxisTitle, Excel.Contrib.Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartGroup WithComCleanup(this Microsoft.Office.Interop.Excel.ChartGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartGroup, Excel.Contrib.Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartGroups WithComCleanup(this Microsoft.Office.Interop.Excel.ChartGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartGroups, Excel.Contrib.Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAxes WithComCleanup(this Microsoft.Office.Interop.Excel.Axes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Axes, Excel.Contrib.Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPoints WithComCleanup(this Microsoft.Office.Interop.Excel.Points resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Points, Excel.Contrib.Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPoint WithComCleanup(this Microsoft.Office.Interop.Excel.Point resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Point, Excel.Contrib.Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISeries WithComCleanup(this Microsoft.Office.Interop.Excel.Series resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Series, Excel.Contrib.Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISeriesCollection WithComCleanup(this Microsoft.Office.Interop.Excel.SeriesCollection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SeriesCollection, Excel.Contrib.Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDataLabel WithComCleanup(this Microsoft.Office.Interop.Excel.DataLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DataLabel, Excel.Contrib.Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDataLabels WithComCleanup(this Microsoft.Office.Interop.Excel.DataLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DataLabels, Excel.Contrib.Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILegendEntry WithComCleanup(this Microsoft.Office.Interop.Excel.LegendEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LegendEntry, Excel.Contrib.Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILegendEntries WithComCleanup(this Microsoft.Office.Interop.Excel.LegendEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LegendEntries, Excel.Contrib.Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILegendKey WithComCleanup(this Microsoft.Office.Interop.Excel.LegendKey resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LegendKey, Excel.Contrib.Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITrendlines WithComCleanup(this Microsoft.Office.Interop.Excel.Trendlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Trendlines, Excel.Contrib.Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITrendline WithComCleanup(this Microsoft.Office.Interop.Excel.Trendline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Trendline, Excel.Contrib.Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICorners WithComCleanup(this Microsoft.Office.Interop.Excel.Corners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Corners, Excel.Contrib.Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISeriesLines WithComCleanup(this Microsoft.Office.Interop.Excel.SeriesLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SeriesLines, Excel.Contrib.Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IHiLoLines WithComCleanup(this Microsoft.Office.Interop.Excel.HiLoLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.HiLoLines, Excel.Contrib.Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGridlines WithComCleanup(this Microsoft.Office.Interop.Excel.Gridlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Gridlines, Excel.Contrib.Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDropLines WithComCleanup(this Microsoft.Office.Interop.Excel.DropLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DropLines, Excel.Contrib.Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILeaderLines WithComCleanup(this Microsoft.Office.Interop.Excel.LeaderLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LeaderLines, Excel.Contrib.Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IUpBars WithComCleanup(this Microsoft.Office.Interop.Excel.UpBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.UpBars, Excel.Contrib.Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDownBars WithComCleanup(this Microsoft.Office.Interop.Excel.DownBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DownBars, Excel.Contrib.Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFloor WithComCleanup(this Microsoft.Office.Interop.Excel.Floor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Floor, Excel.Contrib.Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWalls WithComCleanup(this Microsoft.Office.Interop.Excel.Walls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Walls, Excel.Contrib.Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITickLabels WithComCleanup(this Microsoft.Office.Interop.Excel.TickLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TickLabels, Excel.Contrib.Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPlotArea WithComCleanup(this Microsoft.Office.Interop.Excel.PlotArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PlotArea, Excel.Contrib.Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartArea WithComCleanup(this Microsoft.Office.Interop.Excel.ChartArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartArea, Excel.Contrib.Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILegend WithComCleanup(this Microsoft.Office.Interop.Excel.Legend resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Legend, Excel.Contrib.Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IErrorBars WithComCleanup(this Microsoft.Office.Interop.Excel.ErrorBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ErrorBars, Excel.Contrib.Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDataTable WithComCleanup(this Microsoft.Office.Interop.Excel.DataTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DataTable, Excel.Contrib.Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for Phonetic which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPhonetic WithComCleanup(this Microsoft.Office.Interop.Excel.Phonetic resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Phonetic, Excel.Contrib.Interfaces.IPhonetic>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IShape WithComCleanup(this Microsoft.Office.Interop.Excel.Shape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Shape, Excel.Contrib.Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IShapes WithComCleanup(this Microsoft.Office.Interop.Excel.Shapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Shapes, Excel.Contrib.Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IShapeRange WithComCleanup(this Microsoft.Office.Interop.Excel.ShapeRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ShapeRange, Excel.Contrib.Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGroupShapes WithComCleanup(this Microsoft.Office.Interop.Excel.GroupShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.GroupShapes, Excel.Contrib.Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITextFrame WithComCleanup(this Microsoft.Office.Interop.Excel.TextFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TextFrame, Excel.Contrib.Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IConnectorFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ConnectorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ConnectorFormat, Excel.Contrib.Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFreeformBuilder WithComCleanup(this Microsoft.Office.Interop.Excel.FreeformBuilder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FreeformBuilder, Excel.Contrib.Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for ControlFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IControlFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ControlFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ControlFormat, Excel.Contrib.Interfaces.IControlFormat>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEFormat WithComCleanup(this Microsoft.Office.Interop.Excel.OLEFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEFormat, Excel.Contrib.Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILinkFormat WithComCleanup(this Microsoft.Office.Interop.Excel.LinkFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LinkFormat, Excel.Contrib.Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for PublishObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPublishObjects WithComCleanup(this Microsoft.Office.Interop.Excel.PublishObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PublishObjects, Excel.Contrib.Interfaces.IPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for OLEDBError which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEDBError WithComCleanup(this Microsoft.Office.Interop.Excel.OLEDBError resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEDBError, Excel.Contrib.Interfaces.IOLEDBError>();
		}

		/// <summary>
		/// Wrapper interface for OLEDBErrors which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEDBErrors WithComCleanup(this Microsoft.Office.Interop.Excel.OLEDBErrors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEDBErrors, Excel.Contrib.Interfaces.IOLEDBErrors>();
		}

		/// <summary>
		/// Wrapper interface for Phonetics which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPhonetics WithComCleanup(this Microsoft.Office.Interop.Excel.Phonetics resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Phonetics, Excel.Contrib.Interfaces.IPhonetics>();
		}

		/// <summary>
		/// Wrapper interface for PivotLayout which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotLayout WithComCleanup(this Microsoft.Office.Interop.Excel.PivotLayout resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotLayout, Excel.Contrib.Interfaces.IPivotLayout>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDisplayUnitLabel WithComCleanup(this Microsoft.Office.Interop.Excel.DisplayUnitLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DisplayUnitLabel, Excel.Contrib.Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for CellFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICellFormat WithComCleanup(this Microsoft.Office.Interop.Excel.CellFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CellFormat, Excel.Contrib.Interfaces.ICellFormat>();
		}

		/// <summary>
		/// Wrapper interface for UsedObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IUsedObjects WithComCleanup(this Microsoft.Office.Interop.Excel.UsedObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.UsedObjects, Excel.Contrib.Interfaces.IUsedObjects>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperties which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICustomProperties WithComCleanup(this Microsoft.Office.Interop.Excel.CustomProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CustomProperties, Excel.Contrib.Interfaces.ICustomProperties>();
		}

		/// <summary>
		/// Wrapper interface for CustomProperty which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICustomProperty WithComCleanup(this Microsoft.Office.Interop.Excel.CustomProperty resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CustomProperty, Excel.Contrib.Interfaces.ICustomProperty>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedMembers which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICalculatedMembers WithComCleanup(this Microsoft.Office.Interop.Excel.CalculatedMembers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CalculatedMembers, Excel.Contrib.Interfaces.ICalculatedMembers>();
		}

		/// <summary>
		/// Wrapper interface for CalculatedMember which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ICalculatedMember WithComCleanup(this Microsoft.Office.Interop.Excel.CalculatedMember resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.CalculatedMember, Excel.Contrib.Interfaces.ICalculatedMember>();
		}

		/// <summary>
		/// Wrapper interface for Watches which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWatches WithComCleanup(this Microsoft.Office.Interop.Excel.Watches resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Watches, Excel.Contrib.Interfaces.IWatches>();
		}

		/// <summary>
		/// Wrapper interface for Watch which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWatch WithComCleanup(this Microsoft.Office.Interop.Excel.Watch resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Watch, Excel.Contrib.Interfaces.IWatch>();
		}

		/// <summary>
		/// Wrapper interface for PivotCell which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotCell WithComCleanup(this Microsoft.Office.Interop.Excel.PivotCell resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotCell, Excel.Contrib.Interfaces.IPivotCell>();
		}

		/// <summary>
		/// Wrapper interface for Graphic which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGraphic WithComCleanup(this Microsoft.Office.Interop.Excel.Graphic resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Graphic, Excel.Contrib.Interfaces.IGraphic>();
		}

		/// <summary>
		/// Wrapper interface for AutoRecover which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAutoRecover WithComCleanup(this Microsoft.Office.Interop.Excel.AutoRecover resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AutoRecover, Excel.Contrib.Interfaces.IAutoRecover>();
		}

		/// <summary>
		/// Wrapper interface for ErrorCheckingOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IErrorCheckingOptions WithComCleanup(this Microsoft.Office.Interop.Excel.ErrorCheckingOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ErrorCheckingOptions, Excel.Contrib.Interfaces.IErrorCheckingOptions>();
		}

		/// <summary>
		/// Wrapper interface for Errors which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IErrors WithComCleanup(this Microsoft.Office.Interop.Excel.Errors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Errors, Excel.Contrib.Interfaces.IErrors>();
		}

		/// <summary>
		/// Wrapper interface for Error which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IError WithComCleanup(this Microsoft.Office.Interop.Excel.Error resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Error, Excel.Contrib.Interfaces.IError>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagAction which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTagAction WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTagAction resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTagAction, Excel.Contrib.Interfaces.ISmartTagAction>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagActions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTagActions WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTagActions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTagActions, Excel.Contrib.Interfaces.ISmartTagActions>();
		}

		/// <summary>
		/// Wrapper interface for SmartTag which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTag WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTag resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTag, Excel.Contrib.Interfaces.ISmartTag>();
		}

		/// <summary>
		/// Wrapper interface for SmartTags which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTags WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTags resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTags, Excel.Contrib.Interfaces.ISmartTags>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTagRecognizer WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTagRecognizer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTagRecognizer, Excel.Contrib.Interfaces.ISmartTagRecognizer>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagRecognizers which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTagRecognizers WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTagRecognizers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTagRecognizers, Excel.Contrib.Interfaces.ISmartTagRecognizers>();
		}

		/// <summary>
		/// Wrapper interface for SmartTagOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISmartTagOptions WithComCleanup(this Microsoft.Office.Interop.Excel.SmartTagOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SmartTagOptions, Excel.Contrib.Interfaces.ISmartTagOptions>();
		}

		/// <summary>
		/// Wrapper interface for SpellingOptions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISpellingOptions WithComCleanup(this Microsoft.Office.Interop.Excel.SpellingOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SpellingOptions, Excel.Contrib.Interfaces.ISpellingOptions>();
		}

		/// <summary>
		/// Wrapper interface for Speech which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISpeech WithComCleanup(this Microsoft.Office.Interop.Excel.Speech resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Speech, Excel.Contrib.Interfaces.ISpeech>();
		}

		/// <summary>
		/// Wrapper interface for Protection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IProtection WithComCleanup(this Microsoft.Office.Interop.Excel.Protection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Protection, Excel.Contrib.Interfaces.IProtection>();
		}

		/// <summary>
		/// Wrapper interface for PivotItemList which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotItemList WithComCleanup(this Microsoft.Office.Interop.Excel.PivotItemList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotItemList, Excel.Contrib.Interfaces.IPivotItemList>();
		}

		/// <summary>
		/// Wrapper interface for Tab which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITab WithComCleanup(this Microsoft.Office.Interop.Excel.Tab resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Tab, Excel.Contrib.Interfaces.ITab>();
		}

		/// <summary>
		/// Wrapper interface for AllowEditRanges which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAllowEditRanges WithComCleanup(this Microsoft.Office.Interop.Excel.AllowEditRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AllowEditRanges, Excel.Contrib.Interfaces.IAllowEditRanges>();
		}

		/// <summary>
		/// Wrapper interface for AllowEditRange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAllowEditRange WithComCleanup(this Microsoft.Office.Interop.Excel.AllowEditRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AllowEditRange, Excel.Contrib.Interfaces.IAllowEditRange>();
		}

		/// <summary>
		/// Wrapper interface for UserAccessList which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IUserAccessList WithComCleanup(this Microsoft.Office.Interop.Excel.UserAccessList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.UserAccessList, Excel.Contrib.Interfaces.IUserAccessList>();
		}

		/// <summary>
		/// Wrapper interface for UserAccess which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IUserAccess WithComCleanup(this Microsoft.Office.Interop.Excel.UserAccess resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.UserAccess, Excel.Contrib.Interfaces.IUserAccess>();
		}

		/// <summary>
		/// Wrapper interface for RTD which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRTD WithComCleanup(this Microsoft.Office.Interop.Excel.RTD resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RTD, Excel.Contrib.Interfaces.IRTD>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDiagram WithComCleanup(this Microsoft.Office.Interop.Excel.Diagram resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Diagram, Excel.Contrib.Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for ListObjects which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListObjects WithComCleanup(this Microsoft.Office.Interop.Excel.ListObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListObjects, Excel.Contrib.Interfaces.IListObjects>();
		}

		/// <summary>
		/// Wrapper interface for ListObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListObject WithComCleanup(this Microsoft.Office.Interop.Excel.ListObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListObject, Excel.Contrib.Interfaces.IListObject>();
		}

		/// <summary>
		/// Wrapper interface for ListColumns which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListColumns WithComCleanup(this Microsoft.Office.Interop.Excel.ListColumns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListColumns, Excel.Contrib.Interfaces.IListColumns>();
		}

		/// <summary>
		/// Wrapper interface for ListColumn which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListColumn WithComCleanup(this Microsoft.Office.Interop.Excel.ListColumn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListColumn, Excel.Contrib.Interfaces.IListColumn>();
		}

		/// <summary>
		/// Wrapper interface for ListRows which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListRows WithComCleanup(this Microsoft.Office.Interop.Excel.ListRows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListRows, Excel.Contrib.Interfaces.IListRows>();
		}

		/// <summary>
		/// Wrapper interface for ListRow which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListRow WithComCleanup(this Microsoft.Office.Interop.Excel.ListRow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListRow, Excel.Contrib.Interfaces.IListRow>();
		}

		/// <summary>
		/// Wrapper interface for XmlNamespace which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlNamespace WithComCleanup(this Microsoft.Office.Interop.Excel.XmlNamespace resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlNamespace, Excel.Contrib.Interfaces.IXmlNamespace>();
		}

		/// <summary>
		/// Wrapper interface for XmlNamespaces which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlNamespaces WithComCleanup(this Microsoft.Office.Interop.Excel.XmlNamespaces resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlNamespaces, Excel.Contrib.Interfaces.IXmlNamespaces>();
		}

		/// <summary>
		/// Wrapper interface for XmlDataBinding which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlDataBinding WithComCleanup(this Microsoft.Office.Interop.Excel.XmlDataBinding resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlDataBinding, Excel.Contrib.Interfaces.IXmlDataBinding>();
		}

		/// <summary>
		/// Wrapper interface for XmlSchema which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlSchema WithComCleanup(this Microsoft.Office.Interop.Excel.XmlSchema resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlSchema, Excel.Contrib.Interfaces.IXmlSchema>();
		}

		/// <summary>
		/// Wrapper interface for XmlSchemas which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlSchemas WithComCleanup(this Microsoft.Office.Interop.Excel.XmlSchemas resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlSchemas, Excel.Contrib.Interfaces.IXmlSchemas>();
		}

		/// <summary>
		/// Wrapper interface for XmlMap which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlMap WithComCleanup(this Microsoft.Office.Interop.Excel.XmlMap resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlMap, Excel.Contrib.Interfaces.IXmlMap>();
		}

		/// <summary>
		/// Wrapper interface for XmlMaps which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXmlMaps WithComCleanup(this Microsoft.Office.Interop.Excel.XmlMaps resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XmlMaps, Excel.Contrib.Interfaces.IXmlMaps>();
		}

		/// <summary>
		/// Wrapper interface for ListDataFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IListDataFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ListDataFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ListDataFormat, Excel.Contrib.Interfaces.IListDataFormat>();
		}

		/// <summary>
		/// Wrapper interface for XPath which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IXPath WithComCleanup(this Microsoft.Office.Interop.Excel.XPath resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.XPath, Excel.Contrib.Interfaces.IXPath>();
		}

		/// <summary>
		/// Wrapper interface for PivotLineCells which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotLineCells WithComCleanup(this Microsoft.Office.Interop.Excel.PivotLineCells resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotLineCells, Excel.Contrib.Interfaces.IPivotLineCells>();
		}

		/// <summary>
		/// Wrapper interface for PivotLine which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotLine WithComCleanup(this Microsoft.Office.Interop.Excel.PivotLine resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotLine, Excel.Contrib.Interfaces.IPivotLine>();
		}

		/// <summary>
		/// Wrapper interface for PivotLines which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotLines WithComCleanup(this Microsoft.Office.Interop.Excel.PivotLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotLines, Excel.Contrib.Interfaces.IPivotLines>();
		}

		/// <summary>
		/// Wrapper interface for PivotAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotAxis WithComCleanup(this Microsoft.Office.Interop.Excel.PivotAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotAxis, Excel.Contrib.Interfaces.IPivotAxis>();
		}

		/// <summary>
		/// Wrapper interface for PivotFilter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotFilter WithComCleanup(this Microsoft.Office.Interop.Excel.PivotFilter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotFilter, Excel.Contrib.Interfaces.IPivotFilter>();
		}

		/// <summary>
		/// Wrapper interface for PivotFilters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotFilters WithComCleanup(this Microsoft.Office.Interop.Excel.PivotFilters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotFilters, Excel.Contrib.Interfaces.IPivotFilters>();
		}

		/// <summary>
		/// Wrapper interface for WorkbookConnection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorkbookConnection WithComCleanup(this Microsoft.Office.Interop.Excel.WorkbookConnection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.WorkbookConnection, Excel.Contrib.Interfaces.IWorkbookConnection>();
		}

		/// <summary>
		/// Wrapper interface for Connections which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IConnections WithComCleanup(this Microsoft.Office.Interop.Excel.Connections resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Connections, Excel.Contrib.Interfaces.IConnections>();
		}

		/// <summary>
		/// Wrapper interface for WorksheetView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorksheetView WithComCleanup(this Microsoft.Office.Interop.Excel.WorksheetView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.WorksheetView, Excel.Contrib.Interfaces.IWorksheetView>();
		}

		/// <summary>
		/// Wrapper interface for ChartView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartView WithComCleanup(this Microsoft.Office.Interop.Excel.ChartView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartView, Excel.Contrib.Interfaces.IChartView>();
		}

		/// <summary>
		/// Wrapper interface for ModuleView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IModuleView WithComCleanup(this Microsoft.Office.Interop.Excel.ModuleView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ModuleView, Excel.Contrib.Interfaces.IModuleView>();
		}

		/// <summary>
		/// Wrapper interface for DialogSheetView which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDialogSheetView WithComCleanup(this Microsoft.Office.Interop.Excel.DialogSheetView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DialogSheetView, Excel.Contrib.Interfaces.IDialogSheetView>();
		}

		/// <summary>
		/// Wrapper interface for SheetViews which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISheetViews WithComCleanup(this Microsoft.Office.Interop.Excel.SheetViews resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SheetViews, Excel.Contrib.Interfaces.ISheetViews>();
		}

		/// <summary>
		/// Wrapper interface for OLEDBConnection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEDBConnection WithComCleanup(this Microsoft.Office.Interop.Excel.OLEDBConnection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEDBConnection, Excel.Contrib.Interfaces.IOLEDBConnection>();
		}

		/// <summary>
		/// Wrapper interface for ODBCConnection which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IODBCConnection WithComCleanup(this Microsoft.Office.Interop.Excel.ODBCConnection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ODBCConnection, Excel.Contrib.Interfaces.IODBCConnection>();
		}

		/// <summary>
		/// Wrapper interface for Action which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAction WithComCleanup(this Microsoft.Office.Interop.Excel.Action resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Action, Excel.Contrib.Interfaces.IAction>();
		}

		/// <summary>
		/// Wrapper interface for Actions which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IActions WithComCleanup(this Microsoft.Office.Interop.Excel.Actions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Actions, Excel.Contrib.Interfaces.IActions>();
		}

		/// <summary>
		/// Wrapper interface for FormatColor which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFormatColor WithComCleanup(this Microsoft.Office.Interop.Excel.FormatColor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FormatColor, Excel.Contrib.Interfaces.IFormatColor>();
		}

		/// <summary>
		/// Wrapper interface for ConditionValue which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IConditionValue WithComCleanup(this Microsoft.Office.Interop.Excel.ConditionValue resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ConditionValue, Excel.Contrib.Interfaces.IConditionValue>();
		}

		/// <summary>
		/// Wrapper interface for ColorScale which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IColorScale WithComCleanup(this Microsoft.Office.Interop.Excel.ColorScale resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ColorScale, Excel.Contrib.Interfaces.IColorScale>();
		}

		/// <summary>
		/// Wrapper interface for ColorScaleCriteria which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IColorScaleCriteria WithComCleanup(this Microsoft.Office.Interop.Excel.ColorScaleCriteria resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ColorScaleCriteria, Excel.Contrib.Interfaces.IColorScaleCriteria>();
		}

		/// <summary>
		/// Wrapper interface for ColorScaleCriterion which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IColorScaleCriterion WithComCleanup(this Microsoft.Office.Interop.Excel.ColorScaleCriterion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ColorScaleCriterion, Excel.Contrib.Interfaces.IColorScaleCriterion>();
		}

		/// <summary>
		/// Wrapper interface for Databar which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDatabar WithComCleanup(this Microsoft.Office.Interop.Excel.Databar resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Databar, Excel.Contrib.Interfaces.IDatabar>();
		}

		/// <summary>
		/// Wrapper interface for IconSetCondition which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIconSetCondition WithComCleanup(this Microsoft.Office.Interop.Excel.IconSetCondition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IconSetCondition, Excel.Contrib.Interfaces.IIconSetCondition>();
		}

		/// <summary>
		/// Wrapper interface for IconCriteria which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIconCriteria WithComCleanup(this Microsoft.Office.Interop.Excel.IconCriteria resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IconCriteria, Excel.Contrib.Interfaces.IIconCriteria>();
		}

		/// <summary>
		/// Wrapper interface for IconCriterion which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIconCriterion WithComCleanup(this Microsoft.Office.Interop.Excel.IconCriterion resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IconCriterion, Excel.Contrib.Interfaces.IIconCriterion>();
		}

		/// <summary>
		/// Wrapper interface for Icon which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIcon WithComCleanup(this Microsoft.Office.Interop.Excel.Icon resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Icon, Excel.Contrib.Interfaces.IIcon>();
		}

		/// <summary>
		/// Wrapper interface for IconSet which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIconSet WithComCleanup(this Microsoft.Office.Interop.Excel.IconSet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IconSet, Excel.Contrib.Interfaces.IIconSet>();
		}

		/// <summary>
		/// Wrapper interface for IconSets which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIconSets WithComCleanup(this Microsoft.Office.Interop.Excel.IconSets resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IconSets, Excel.Contrib.Interfaces.IIconSets>();
		}

		/// <summary>
		/// Wrapper interface for Top10 which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITop10 WithComCleanup(this Microsoft.Office.Interop.Excel.Top10 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Top10, Excel.Contrib.Interfaces.ITop10>();
		}

		/// <summary>
		/// Wrapper interface for AboveAverage which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAboveAverage WithComCleanup(this Microsoft.Office.Interop.Excel.AboveAverage resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AboveAverage, Excel.Contrib.Interfaces.IAboveAverage>();
		}

		/// <summary>
		/// Wrapper interface for UniqueValues which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IUniqueValues WithComCleanup(this Microsoft.Office.Interop.Excel.UniqueValues resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.UniqueValues, Excel.Contrib.Interfaces.IUniqueValues>();
		}

		/// <summary>
		/// Wrapper interface for Ranges which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRanges WithComCleanup(this Microsoft.Office.Interop.Excel.Ranges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Ranges, Excel.Contrib.Interfaces.IRanges>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IHeaderFooter WithComCleanup(this Microsoft.Office.Interop.Excel.HeaderFooter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.HeaderFooter, Excel.Contrib.Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for Page which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPage WithComCleanup(this Microsoft.Office.Interop.Excel.Page resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Page, Excel.Contrib.Interfaces.IPage>();
		}

		/// <summary>
		/// Wrapper interface for Pages which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPages WithComCleanup(this Microsoft.Office.Interop.Excel.Pages resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Pages, Excel.Contrib.Interfaces.IPages>();
		}

		/// <summary>
		/// Wrapper interface for ServerViewableItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IServerViewableItems WithComCleanup(this Microsoft.Office.Interop.Excel.ServerViewableItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ServerViewableItems, Excel.Contrib.Interfaces.IServerViewableItems>();
		}

		/// <summary>
		/// Wrapper interface for TableStyleElement which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITableStyleElement WithComCleanup(this Microsoft.Office.Interop.Excel.TableStyleElement resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TableStyleElement, Excel.Contrib.Interfaces.ITableStyleElement>();
		}

		/// <summary>
		/// Wrapper interface for TableStyleElements which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITableStyleElements WithComCleanup(this Microsoft.Office.Interop.Excel.TableStyleElements resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TableStyleElements, Excel.Contrib.Interfaces.ITableStyleElements>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITableStyle WithComCleanup(this Microsoft.Office.Interop.Excel.TableStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TableStyle, Excel.Contrib.Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for TableStyles which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ITableStyles WithComCleanup(this Microsoft.Office.Interop.Excel.TableStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.TableStyles, Excel.Contrib.Interfaces.ITableStyles>();
		}

		/// <summary>
		/// Wrapper interface for SortField which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISortField WithComCleanup(this Microsoft.Office.Interop.Excel.SortField resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SortField, Excel.Contrib.Interfaces.ISortField>();
		}

		/// <summary>
		/// Wrapper interface for SortFields which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISortFields WithComCleanup(this Microsoft.Office.Interop.Excel.SortFields resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SortFields, Excel.Contrib.Interfaces.ISortFields>();
		}

		/// <summary>
		/// Wrapper interface for Sort which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISort WithComCleanup(this Microsoft.Office.Interop.Excel.Sort resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Sort, Excel.Contrib.Interfaces.ISort>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IResearch WithComCleanup(this Microsoft.Office.Interop.Excel.Research resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Research, Excel.Contrib.Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for ColorStop which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IColorStop WithComCleanup(this Microsoft.Office.Interop.Excel.ColorStop resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ColorStop, Excel.Contrib.Interfaces.IColorStop>();
		}

		/// <summary>
		/// Wrapper interface for ColorStops which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IColorStops WithComCleanup(this Microsoft.Office.Interop.Excel.ColorStops resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ColorStops, Excel.Contrib.Interfaces.IColorStops>();
		}

		/// <summary>
		/// Wrapper interface for LinearGradient which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ILinearGradient WithComCleanup(this Microsoft.Office.Interop.Excel.LinearGradient resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.LinearGradient, Excel.Contrib.Interfaces.ILinearGradient>();
		}

		/// <summary>
		/// Wrapper interface for RectangularGradient which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRectangularGradient WithComCleanup(this Microsoft.Office.Interop.Excel.RectangularGradient resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RectangularGradient, Excel.Contrib.Interfaces.IRectangularGradient>();
		}

		/// <summary>
		/// Wrapper interface for MultiThreadedCalculation which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IMultiThreadedCalculation WithComCleanup(this Microsoft.Office.Interop.Excel.MultiThreadedCalculation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.MultiThreadedCalculation, Excel.Contrib.Interfaces.IMultiThreadedCalculation>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartFormat WithComCleanup(this Microsoft.Office.Interop.Excel.ChartFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartFormat, Excel.Contrib.Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for FileExportConverter which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFileExportConverter WithComCleanup(this Microsoft.Office.Interop.Excel.FileExportConverter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FileExportConverter, Excel.Contrib.Interfaces.IFileExportConverter>();
		}

		/// <summary>
		/// Wrapper interface for FileExportConverters which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IFileExportConverters WithComCleanup(this Microsoft.Office.Interop.Excel.FileExportConverters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.FileExportConverters, Excel.Contrib.Interfaces.IFileExportConverters>();
		}

		/// <summary>
		/// Wrapper interface for AddIns2 which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAddIns2 WithComCleanup(this Microsoft.Office.Interop.Excel.AddIns2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AddIns2, Excel.Contrib.Interfaces.IAddIns2>();
		}

		/// <summary>
		/// Wrapper interface for SparklineGroups which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparklineGroups WithComCleanup(this Microsoft.Office.Interop.Excel.SparklineGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparklineGroups, Excel.Contrib.Interfaces.ISparklineGroups>();
		}

		/// <summary>
		/// Wrapper interface for SparklineGroup which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparklineGroup WithComCleanup(this Microsoft.Office.Interop.Excel.SparklineGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparklineGroup, Excel.Contrib.Interfaces.ISparklineGroup>();
		}

		/// <summary>
		/// Wrapper interface for SparkPoints which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparkPoints WithComCleanup(this Microsoft.Office.Interop.Excel.SparkPoints resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparkPoints, Excel.Contrib.Interfaces.ISparkPoints>();
		}

		/// <summary>
		/// Wrapper interface for Sparkline which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparkline WithComCleanup(this Microsoft.Office.Interop.Excel.Sparkline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Sparkline, Excel.Contrib.Interfaces.ISparkline>();
		}

		/// <summary>
		/// Wrapper interface for SparkAxes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparkAxes WithComCleanup(this Microsoft.Office.Interop.Excel.SparkAxes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparkAxes, Excel.Contrib.Interfaces.ISparkAxes>();
		}

		/// <summary>
		/// Wrapper interface for SparkHorizontalAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparkHorizontalAxis WithComCleanup(this Microsoft.Office.Interop.Excel.SparkHorizontalAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparkHorizontalAxis, Excel.Contrib.Interfaces.ISparkHorizontalAxis>();
		}

		/// <summary>
		/// Wrapper interface for SparkVerticalAxis which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparkVerticalAxis WithComCleanup(this Microsoft.Office.Interop.Excel.SparkVerticalAxis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparkVerticalAxis, Excel.Contrib.Interfaces.ISparkVerticalAxis>();
		}

		/// <summary>
		/// Wrapper interface for SparkColor which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISparkColor WithComCleanup(this Microsoft.Office.Interop.Excel.SparkColor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SparkColor, Excel.Contrib.Interfaces.ISparkColor>();
		}

		/// <summary>
		/// Wrapper interface for DataBarBorder which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDataBarBorder WithComCleanup(this Microsoft.Office.Interop.Excel.DataBarBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DataBarBorder, Excel.Contrib.Interfaces.IDataBarBorder>();
		}

		/// <summary>
		/// Wrapper interface for NegativeBarFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.INegativeBarFormat WithComCleanup(this Microsoft.Office.Interop.Excel.NegativeBarFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.NegativeBarFormat, Excel.Contrib.Interfaces.INegativeBarFormat>();
		}

		/// <summary>
		/// Wrapper interface for ValueChange which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IValueChange WithComCleanup(this Microsoft.Office.Interop.Excel.ValueChange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ValueChange, Excel.Contrib.Interfaces.IValueChange>();
		}

		/// <summary>
		/// Wrapper interface for PivotTableChangeList which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IPivotTableChangeList WithComCleanup(this Microsoft.Office.Interop.Excel.PivotTableChangeList resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.PivotTableChangeList, Excel.Contrib.Interfaces.IPivotTableChangeList>();
		}

		/// <summary>
		/// Wrapper interface for DisplayFormat which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDisplayFormat WithComCleanup(this Microsoft.Office.Interop.Excel.DisplayFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DisplayFormat, Excel.Contrib.Interfaces.IDisplayFormat>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCaches which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerCaches WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerCaches resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerCaches, Excel.Contrib.Interfaces.ISlicerCaches>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCache which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerCache WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerCache resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerCache, Excel.Contrib.Interfaces.ISlicerCache>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCacheLevels which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerCacheLevels WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerCacheLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerCacheLevels, Excel.Contrib.Interfaces.ISlicerCacheLevels>();
		}

		/// <summary>
		/// Wrapper interface for SlicerCacheLevel which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerCacheLevel WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerCacheLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerCacheLevel, Excel.Contrib.Interfaces.ISlicerCacheLevel>();
		}

		/// <summary>
		/// Wrapper interface for Slicers which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicers WithComCleanup(this Microsoft.Office.Interop.Excel.Slicers resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Slicers, Excel.Contrib.Interfaces.ISlicers>();
		}

		/// <summary>
		/// Wrapper interface for Slicer which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicer WithComCleanup(this Microsoft.Office.Interop.Excel.Slicer resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Slicer, Excel.Contrib.Interfaces.ISlicer>();
		}

		/// <summary>
		/// Wrapper interface for SlicerItem which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerItem WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerItem resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerItem, Excel.Contrib.Interfaces.ISlicerItem>();
		}

		/// <summary>
		/// Wrapper interface for SlicerItems which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerItems WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerItems resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerItems, Excel.Contrib.Interfaces.ISlicerItems>();
		}

		/// <summary>
		/// Wrapper interface for SlicerPivotTables which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.ISlicerPivotTables WithComCleanup(this Microsoft.Office.Interop.Excel.SlicerPivotTables resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.SlicerPivotTables, Excel.Contrib.Interfaces.ISlicerPivotTables>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IProtectedViewWindows WithComCleanup(this Microsoft.Office.Interop.Excel.ProtectedViewWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ProtectedViewWindows, Excel.Contrib.Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IProtectedViewWindow WithComCleanup(this Microsoft.Office.Interop.Excel.ProtectedViewWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ProtectedViewWindow, Excel.Contrib.Interfaces.IProtectedViewWindow>();
		}

		/// <summary>
		/// Wrapper interface for IDummy which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IIDummy WithComCleanup(this Microsoft.Office.Interop.Excel.IDummy resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.IDummy, Excel.Contrib.Interfaces.IIDummy>();
		}

		/// <summary>
		/// Wrapper interface for ICanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IICanvasShapes WithComCleanup(this Microsoft.Office.Interop.Excel.ICanvasShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ICanvasShapes, Excel.Contrib.Interfaces.IICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for RefreshEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IRefreshEvents_Event WithComCleanup(this Microsoft.Office.Interop.Excel.RefreshEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.RefreshEvents_Event, Excel.Contrib.Interfaces.IRefreshEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for QueryTable which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IQueryTable WithComCleanup(this Microsoft.Office.Interop.Excel.QueryTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.QueryTable, Excel.Contrib.Interfaces.IQueryTable>();
		}

		/// <summary>
		/// Wrapper interface for AppEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IAppEvents_Event WithComCleanup(this Microsoft.Office.Interop.Excel.AppEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.AppEvents_Event, Excel.Contrib.Interfaces.IAppEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IApplication WithComCleanup(this Microsoft.Office.Interop.Excel.Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Application, Excel.Contrib.Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for ChartEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChartEvents_Event WithComCleanup(this Microsoft.Office.Interop.Excel.ChartEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.ChartEvents_Event, Excel.Contrib.Interfaces.IChartEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IChart WithComCleanup(this Microsoft.Office.Interop.Excel.Chart resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Chart, Excel.Contrib.Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for DocEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IDocEvents_Event WithComCleanup(this Microsoft.Office.Interop.Excel.DocEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.DocEvents_Event, Excel.Contrib.Interfaces.IDocEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Worksheet which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorksheet WithComCleanup(this Microsoft.Office.Interop.Excel.Worksheet resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Worksheet, Excel.Contrib.Interfaces.IWorksheet>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IGlobal WithComCleanup(this Microsoft.Office.Interop.Excel.Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Global, Excel.Contrib.Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for WorkbookEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorkbookEvents_Event WithComCleanup(this Microsoft.Office.Interop.Excel.WorkbookEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.WorkbookEvents_Event, Excel.Contrib.Interfaces.IWorkbookEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Workbook which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IWorkbook WithComCleanup(this Microsoft.Office.Interop.Excel.Workbook resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.Workbook, Excel.Contrib.Interfaces.IWorkbook>();
		}

		/// <summary>
		/// Wrapper interface for OLEObjectEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEObjectEvents_Event WithComCleanup(this Microsoft.Office.Interop.Excel.OLEObjectEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEObjectEvents_Event, Excel.Contrib.Interfaces.IOLEObjectEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEObject which adds IDispose to the interface
		/// </summary>
		public static Excel.Contrib.Interfaces.IOLEObject WithComCleanup(this Microsoft.Office.Interop.Excel.OLEObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.Excel.OLEObject, Excel.Contrib.Interfaces.IOLEObject>();
		}

	}
}