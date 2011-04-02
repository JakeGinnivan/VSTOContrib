using Office.Contrib.Extensions;

namespace Office.Contrib.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for Collection which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICollection WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Collection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Collection, PowerPoint.Contrib.Interfaces.ICollection>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.I_Application WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Application, PowerPoint.Contrib.Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.I_Global WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Global, PowerPoint.Contrib.Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for EApplication_Event which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IEApplication_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EApplication_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EApplication_Event, PowerPoint.Contrib.Interfaces.IEApplication_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IApplication WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Application, PowerPoint.Contrib.Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IGlobal WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Global, PowerPoint.Contrib.Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IColorFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorFormat, PowerPoint.Contrib.Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowWindow which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlideShowWindow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowWindow, PowerPoint.Contrib.Interfaces.ISlideShowWindow>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISelection WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Selection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Selection, PowerPoint.Contrib.Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for DocumentWindows which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDocumentWindows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DocumentWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DocumentWindows, PowerPoint.Contrib.Interfaces.IDocumentWindows>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowWindows which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlideShowWindows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowWindows, PowerPoint.Contrib.Interfaces.ISlideShowWindows>();
		}

		/// <summary>
		/// Wrapper interface for DocumentWindow which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDocumentWindow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DocumentWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DocumentWindow, PowerPoint.Contrib.Interfaces.IDocumentWindow>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IView WithComCleanup(this Microsoft.Office.Interop.PowerPoint.View resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.View, PowerPoint.Contrib.Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowView which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlideShowView WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowView, PowerPoint.Contrib.Interfaces.ISlideShowView>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowSettings which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlideShowSettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowSettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowSettings, PowerPoint.Contrib.Interfaces.ISlideShowSettings>();
		}

		/// <summary>
		/// Wrapper interface for NamedSlideShows which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.INamedSlideShows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.NamedSlideShows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.NamedSlideShows, PowerPoint.Contrib.Interfaces.INamedSlideShows>();
		}

		/// <summary>
		/// Wrapper interface for NamedSlideShow which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.INamedSlideShow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.NamedSlideShow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.NamedSlideShow, PowerPoint.Contrib.Interfaces.INamedSlideShow>();
		}

		/// <summary>
		/// Wrapper interface for PrintOptions which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPrintOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PrintOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PrintOptions, PowerPoint.Contrib.Interfaces.IPrintOptions>();
		}

		/// <summary>
		/// Wrapper interface for PrintRanges which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPrintRanges WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PrintRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PrintRanges, PowerPoint.Contrib.Interfaces.IPrintRanges>();
		}

		/// <summary>
		/// Wrapper interface for PrintRange which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPrintRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PrintRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PrintRange, PowerPoint.Contrib.Interfaces.IPrintRange>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAddIns WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AddIns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AddIns, PowerPoint.Contrib.Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAddIn WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AddIn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AddIn, PowerPoint.Contrib.Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for Presentations which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPresentations WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Presentations resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Presentations, PowerPoint.Contrib.Interfaces.IPresentations>();
		}

		/// <summary>
		/// Wrapper interface for PresEvents which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPresEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PresEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PresEvents, PowerPoint.Contrib.Interfaces.IPresEvents>();
		}

		/// <summary>
		/// Wrapper interface for PresEvents_Event which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPresEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PresEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PresEvents_Event, PowerPoint.Contrib.Interfaces.IPresEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Presentation which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPresentation WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Presentation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Presentation, PowerPoint.Contrib.Interfaces.IPresentation>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IHyperlinks WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Hyperlinks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Hyperlinks, PowerPoint.Contrib.Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IHyperlink WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Hyperlink resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Hyperlink, PowerPoint.Contrib.Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPageSetup WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PageSetup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PageSetup, PowerPoint.Contrib.Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for Fonts which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFonts WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Fonts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Fonts, PowerPoint.Contrib.Interfaces.IFonts>();
		}

		/// <summary>
		/// Wrapper interface for ExtraColors which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IExtraColors WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ExtraColors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ExtraColors, PowerPoint.Contrib.Interfaces.IExtraColors>();
		}

		/// <summary>
		/// Wrapper interface for Slides which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlides WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Slides resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Slides, PowerPoint.Contrib.Interfaces.ISlides>();
		}

		/// <summary>
		/// Wrapper interface for _Slide which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.I_Slide WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Slide resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Slide, PowerPoint.Contrib.Interfaces.I_Slide>();
		}

		/// <summary>
		/// Wrapper interface for SlideRange which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlideRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideRange, PowerPoint.Contrib.Interfaces.ISlideRange>();
		}

		/// <summary>
		/// Wrapper interface for _Master which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.I_Master WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Master resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Master, PowerPoint.Contrib.Interfaces.I_Master>();
		}

		/// <summary>
		/// Wrapper interface for SldEvents which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISldEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SldEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SldEvents, PowerPoint.Contrib.Interfaces.ISldEvents>();
		}

		/// <summary>
		/// Wrapper interface for SldEvents_Event which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISldEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SldEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SldEvents_Event, PowerPoint.Contrib.Interfaces.ISldEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Slide which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlide WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Slide resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Slide, PowerPoint.Contrib.Interfaces.ISlide>();
		}

		/// <summary>
		/// Wrapper interface for ColorSchemes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IColorSchemes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorSchemes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorSchemes, PowerPoint.Contrib.Interfaces.IColorSchemes>();
		}

		/// <summary>
		/// Wrapper interface for ColorScheme which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IColorScheme WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorScheme resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorScheme, PowerPoint.Contrib.Interfaces.IColorScheme>();
		}

		/// <summary>
		/// Wrapper interface for RGBColor which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRGBColor WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RGBColor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RGBColor, PowerPoint.Contrib.Interfaces.IRGBColor>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowTransition which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISlideShowTransition WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowTransition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowTransition, PowerPoint.Contrib.Interfaces.ISlideShowTransition>();
		}

		/// <summary>
		/// Wrapper interface for SoundEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISoundEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SoundEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SoundEffect, PowerPoint.Contrib.Interfaces.ISoundEffect>();
		}

		/// <summary>
		/// Wrapper interface for SoundFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISoundFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SoundFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SoundFormat, PowerPoint.Contrib.Interfaces.ISoundFormat>();
		}

		/// <summary>
		/// Wrapper interface for HeadersFooters which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IHeadersFooters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.HeadersFooters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.HeadersFooters, PowerPoint.Contrib.Interfaces.IHeadersFooters>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IShapes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Shapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Shapes, PowerPoint.Contrib.Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for Placeholders which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPlaceholders WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Placeholders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Placeholders, PowerPoint.Contrib.Interfaces.IPlaceholders>();
		}

		/// <summary>
		/// Wrapper interface for PlaceholderFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPlaceholderFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PlaceholderFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PlaceholderFormat, PowerPoint.Contrib.Interfaces.IPlaceholderFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFreeformBuilder WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FreeformBuilder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FreeformBuilder, PowerPoint.Contrib.Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IShape WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Shape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Shape, PowerPoint.Contrib.Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IShapeRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShapeRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShapeRange, PowerPoint.Contrib.Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IGroupShapes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.GroupShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.GroupShapes, PowerPoint.Contrib.Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAdjustments WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Adjustments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Adjustments, PowerPoint.Contrib.Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPictureFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PictureFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PictureFormat, PowerPoint.Contrib.Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFillFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FillFormat, PowerPoint.Contrib.Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILineFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LineFormat, PowerPoint.Contrib.Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IShadowFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShadowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShadowFormat, PowerPoint.Contrib.Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IConnectorFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ConnectorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ConnectorFormat, PowerPoint.Contrib.Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextEffectFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextEffectFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextEffectFormat, PowerPoint.Contrib.Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IThreeDFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ThreeDFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ThreeDFormat, PowerPoint.Contrib.Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextFrame WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextFrame, PowerPoint.Contrib.Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICalloutFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CalloutFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CalloutFormat, PowerPoint.Contrib.Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IShapeNodes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShapeNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShapeNodes, PowerPoint.Contrib.Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IShapeNode WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShapeNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShapeNode, PowerPoint.Contrib.Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IOLEFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OLEFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OLEFormat, PowerPoint.Contrib.Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILinkFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LinkFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LinkFormat, PowerPoint.Contrib.Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for ObjectVerbs which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IObjectVerbs WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ObjectVerbs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ObjectVerbs, PowerPoint.Contrib.Interfaces.IObjectVerbs>();
		}

		/// <summary>
		/// Wrapper interface for AnimationSettings which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAnimationSettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationSettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationSettings, PowerPoint.Contrib.Interfaces.IAnimationSettings>();
		}

		/// <summary>
		/// Wrapper interface for ActionSettings which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IActionSettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ActionSettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ActionSettings, PowerPoint.Contrib.Interfaces.IActionSettings>();
		}

		/// <summary>
		/// Wrapper interface for ActionSetting which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IActionSetting WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ActionSetting resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ActionSetting, PowerPoint.Contrib.Interfaces.IActionSetting>();
		}

		/// <summary>
		/// Wrapper interface for PlaySettings which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPlaySettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PlaySettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PlaySettings, PowerPoint.Contrib.Interfaces.IPlaySettings>();
		}

		/// <summary>
		/// Wrapper interface for TextRange which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextRange, PowerPoint.Contrib.Interfaces.ITextRange>();
		}

		/// <summary>
		/// Wrapper interface for Ruler which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRuler WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Ruler resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Ruler, PowerPoint.Contrib.Interfaces.IRuler>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevels which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRulerLevels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RulerLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RulerLevels, PowerPoint.Contrib.Interfaces.IRulerLevels>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevel which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRulerLevel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RulerLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RulerLevel, PowerPoint.Contrib.Interfaces.IRulerLevel>();
		}

		/// <summary>
		/// Wrapper interface for TabStops which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITabStops WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TabStops resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TabStops, PowerPoint.Contrib.Interfaces.ITabStops>();
		}

		/// <summary>
		/// Wrapper interface for TabStop which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITabStop WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TabStop resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TabStop, PowerPoint.Contrib.Interfaces.ITabStop>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFont WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Font, PowerPoint.Contrib.Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IParagraphFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ParagraphFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ParagraphFormat, PowerPoint.Contrib.Interfaces.IParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for BulletFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IBulletFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.BulletFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.BulletFormat, PowerPoint.Contrib.Interfaces.IBulletFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextStyles which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextStyles WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyles, PowerPoint.Contrib.Interfaces.ITextStyles>();
		}

		/// <summary>
		/// Wrapper interface for TextStyle which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextStyle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyle, PowerPoint.Contrib.Interfaces.ITextStyle>();
		}

		/// <summary>
		/// Wrapper interface for TextStyleLevels which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextStyleLevels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyleLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyleLevels, PowerPoint.Contrib.Interfaces.ITextStyleLevels>();
		}

		/// <summary>
		/// Wrapper interface for TextStyleLevel which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextStyleLevel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyleLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyleLevel, PowerPoint.Contrib.Interfaces.ITextStyleLevel>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IHeaderFooter WithComCleanup(this Microsoft.Office.Interop.PowerPoint.HeaderFooter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.HeaderFooter, PowerPoint.Contrib.Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for _Presentation which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.I_Presentation WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Presentation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Presentation, PowerPoint.Contrib.Interfaces.I_Presentation>();
		}

		/// <summary>
		/// Wrapper interface for Tags which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITags WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Tags resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Tags, PowerPoint.Contrib.Interfaces.ITags>();
		}

		/// <summary>
		/// Wrapper interface for MouseTracker which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMouseTracker WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MouseTracker resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MouseTracker, PowerPoint.Contrib.Interfaces.IMouseTracker>();
		}

		/// <summary>
		/// Wrapper interface for MouseDownHandler which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMouseDownHandler WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MouseDownHandler resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MouseDownHandler, PowerPoint.Contrib.Interfaces.IMouseDownHandler>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtender which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IOCXExtender WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OCXExtender resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OCXExtender, PowerPoint.Contrib.Interfaces.IOCXExtender>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtenderEvents which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IOCXExtenderEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents, PowerPoint.Contrib.Interfaces.IOCXExtenderEvents>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtenderEvents_Event which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IOCXExtenderEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event, PowerPoint.Contrib.Interfaces.IOCXExtenderEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEControl which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IOLEControl WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OLEControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OLEControl, PowerPoint.Contrib.Interfaces.IOLEControl>();
		}

		/// <summary>
		/// Wrapper interface for EApplication which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IEApplication WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EApplication resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EApplication, PowerPoint.Contrib.Interfaces.IEApplication>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITable WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Table resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Table, PowerPoint.Contrib.Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IColumns WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Columns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Columns, PowerPoint.Contrib.Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IColumn WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Column resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Column, PowerPoint.Contrib.Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for Rows which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Rows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Rows, PowerPoint.Contrib.Interfaces.IRows>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Row resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Row, PowerPoint.Contrib.Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for CellRange which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICellRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CellRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CellRange, PowerPoint.Contrib.Interfaces.ICellRange>();
		}

		/// <summary>
		/// Wrapper interface for Cell which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICell WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Cell resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Cell, PowerPoint.Contrib.Interfaces.ICell>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IBorders WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Borders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Borders, PowerPoint.Contrib.Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPanes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Panes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Panes, PowerPoint.Contrib.Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPane WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Pane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Pane, PowerPoint.Contrib.Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDefaultWebOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DefaultWebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DefaultWebOptions, PowerPoint.Contrib.Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IWebOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.WebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.WebOptions, PowerPoint.Contrib.Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for PublishObjects which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPublishObjects WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PublishObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PublishObjects, PowerPoint.Contrib.Interfaces.IPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for PublishObject which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPublishObject WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PublishObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PublishObject, PowerPoint.Contrib.Interfaces.IPublishObject>();
		}

		/// <summary>
		/// Wrapper interface for MasterEvents which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMasterEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MasterEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MasterEvents, PowerPoint.Contrib.Interfaces.IMasterEvents>();
		}

		/// <summary>
		/// Wrapper interface for MasterEvents_Event which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMasterEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MasterEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MasterEvents_Event, PowerPoint.Contrib.Interfaces.IMasterEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Master which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMaster WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Master resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Master, PowerPoint.Contrib.Interfaces.IMaster>();
		}

		/// <summary>
		/// Wrapper interface for _PowerRex which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.I_PowerRex WithComCleanup(this Microsoft.Office.Interop.PowerPoint._PowerRex resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._PowerRex, PowerPoint.Contrib.Interfaces.I_PowerRex>();
		}

		/// <summary>
		/// Wrapper interface for PowerRex which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPowerRex WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PowerRex resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PowerRex, PowerPoint.Contrib.Interfaces.IPowerRex>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IComments WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Comments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Comments, PowerPoint.Contrib.Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IComment WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Comment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Comment, PowerPoint.Contrib.Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for Designs which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDesigns WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Designs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Designs, PowerPoint.Contrib.Interfaces.IDesigns>();
		}

		/// <summary>
		/// Wrapper interface for Design which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDesign WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Design resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Design, PowerPoint.Contrib.Interfaces.IDesign>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDiagramNode WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DiagramNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DiagramNode, PowerPoint.Contrib.Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDiagramNodeChildren WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren, PowerPoint.Contrib.Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDiagramNodes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DiagramNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DiagramNodes, PowerPoint.Contrib.Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDiagram WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Diagram resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Diagram, PowerPoint.Contrib.Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for TimeLine which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITimeLine WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TimeLine resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TimeLine, PowerPoint.Contrib.Interfaces.ITimeLine>();
		}

		/// <summary>
		/// Wrapper interface for Sequences which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISequences WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Sequences resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Sequences, PowerPoint.Contrib.Interfaces.ISequences>();
		}

		/// <summary>
		/// Wrapper interface for Sequence which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISequence WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Sequence resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Sequence, PowerPoint.Contrib.Interfaces.ISequence>();
		}

		/// <summary>
		/// Wrapper interface for Effect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Effect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Effect, PowerPoint.Contrib.Interfaces.IEffect>();
		}

		/// <summary>
		/// Wrapper interface for Timing which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITiming WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Timing resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Timing, PowerPoint.Contrib.Interfaces.ITiming>();
		}

		/// <summary>
		/// Wrapper interface for EffectParameters which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IEffectParameters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EffectParameters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EffectParameters, PowerPoint.Contrib.Interfaces.IEffectParameters>();
		}

		/// <summary>
		/// Wrapper interface for EffectInformation which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IEffectInformation WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EffectInformation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EffectInformation, PowerPoint.Contrib.Interfaces.IEffectInformation>();
		}

		/// <summary>
		/// Wrapper interface for AnimationBehaviors which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAnimationBehaviors WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationBehaviors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationBehaviors, PowerPoint.Contrib.Interfaces.IAnimationBehaviors>();
		}

		/// <summary>
		/// Wrapper interface for AnimationBehavior which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAnimationBehavior WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationBehavior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationBehavior, PowerPoint.Contrib.Interfaces.IAnimationBehavior>();
		}

		/// <summary>
		/// Wrapper interface for MotionEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMotionEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MotionEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MotionEffect, PowerPoint.Contrib.Interfaces.IMotionEffect>();
		}

		/// <summary>
		/// Wrapper interface for ColorEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IColorEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorEffect, PowerPoint.Contrib.Interfaces.IColorEffect>();
		}

		/// <summary>
		/// Wrapper interface for ScaleEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IScaleEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ScaleEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ScaleEffect, PowerPoint.Contrib.Interfaces.IScaleEffect>();
		}

		/// <summary>
		/// Wrapper interface for RotationEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IRotationEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RotationEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RotationEffect, PowerPoint.Contrib.Interfaces.IRotationEffect>();
		}

		/// <summary>
		/// Wrapper interface for PropertyEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPropertyEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PropertyEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PropertyEffect, PowerPoint.Contrib.Interfaces.IPropertyEffect>();
		}

		/// <summary>
		/// Wrapper interface for AnimationPoints which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAnimationPoints WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationPoints resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationPoints, PowerPoint.Contrib.Interfaces.IAnimationPoints>();
		}

		/// <summary>
		/// Wrapper interface for AnimationPoint which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAnimationPoint WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationPoint resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationPoint, PowerPoint.Contrib.Interfaces.IAnimationPoint>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICanvasShapes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CanvasShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CanvasShapes, PowerPoint.Contrib.Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAutoCorrect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AutoCorrect, PowerPoint.Contrib.Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for Options which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Options resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Options, PowerPoint.Contrib.Interfaces.IOptions>();
		}

		/// <summary>
		/// Wrapper interface for CommandEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICommandEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CommandEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CommandEffect, PowerPoint.Contrib.Interfaces.ICommandEffect>();
		}

		/// <summary>
		/// Wrapper interface for FilterEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFilterEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FilterEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FilterEffect, PowerPoint.Contrib.Interfaces.IFilterEffect>();
		}

		/// <summary>
		/// Wrapper interface for SetEffect which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISetEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SetEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SetEffect, PowerPoint.Contrib.Interfaces.ISetEffect>();
		}

		/// <summary>
		/// Wrapper interface for CustomLayouts which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICustomLayouts WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CustomLayouts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CustomLayouts, PowerPoint.Contrib.Interfaces.ICustomLayouts>();
		}

		/// <summary>
		/// Wrapper interface for CustomLayout which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICustomLayout WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CustomLayout resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CustomLayout, PowerPoint.Contrib.Interfaces.ICustomLayout>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITableStyle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TableStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TableStyle, PowerPoint.Contrib.Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for CustomerData which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICustomerData WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CustomerData resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CustomerData, PowerPoint.Contrib.Interfaces.ICustomerData>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IResearch WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Research resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Research, PowerPoint.Contrib.Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for TableBackground which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITableBackground WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TableBackground resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TableBackground, PowerPoint.Contrib.Interfaces.ITableBackground>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITextFrame2 WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextFrame2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextFrame2, PowerPoint.Contrib.Interfaces.ITextFrame2>();
		}

		/// <summary>
		/// Wrapper interface for FileConverters which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFileConverters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FileConverters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FileConverters, PowerPoint.Contrib.Interfaces.IFileConverters>();
		}

		/// <summary>
		/// Wrapper interface for FileConverter which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFileConverter WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FileConverter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FileConverter, PowerPoint.Contrib.Interfaces.IFileConverter>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAxes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Axes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Axes, PowerPoint.Contrib.Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAxis WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Axis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Axis, PowerPoint.Contrib.Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IAxisTitle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AxisTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AxisTitle, PowerPoint.Contrib.Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChart WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Chart resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Chart, PowerPoint.Contrib.Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for ChartBorder which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartBorder WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartBorder, PowerPoint.Contrib.Interfaces.IChartBorder>();
		}

		/// <summary>
		/// Wrapper interface for ChartCharacters which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartCharacters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartCharacters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartCharacters, PowerPoint.Contrib.Interfaces.IChartCharacters>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartArea WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartArea, PowerPoint.Contrib.Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartColorFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartColorFormat, PowerPoint.Contrib.Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartData which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartData WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartData resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartData, PowerPoint.Contrib.Interfaces.IChartData>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartFillFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartFillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartFillFormat, PowerPoint.Contrib.Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartFormat, PowerPoint.Contrib.Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartGroup WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartGroup, PowerPoint.Contrib.Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartGroups WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartGroups, PowerPoint.Contrib.Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartTitle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartTitle, PowerPoint.Contrib.Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICorners WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Corners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Corners, PowerPoint.Contrib.Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDataLabel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DataLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DataLabel, PowerPoint.Contrib.Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDataLabels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DataLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DataLabels, PowerPoint.Contrib.Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDataTable WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DataTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DataTable, PowerPoint.Contrib.Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDisplayUnitLabel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel, PowerPoint.Contrib.Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDownBars WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DownBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DownBars, PowerPoint.Contrib.Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IDropLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DropLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DropLines, PowerPoint.Contrib.Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IErrorBars WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ErrorBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ErrorBars, PowerPoint.Contrib.Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IFloor WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Floor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Floor, PowerPoint.Contrib.Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IChartFont WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartFont, PowerPoint.Contrib.Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IGridlines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Gridlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Gridlines, PowerPoint.Contrib.Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IHiLoLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.HiLoLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.HiLoLines, PowerPoint.Contrib.Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IInterior WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Interior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Interior, PowerPoint.Contrib.Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILeaderLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LeaderLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LeaderLines, PowerPoint.Contrib.Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILegend WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Legend resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Legend, PowerPoint.Contrib.Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILegendEntries WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LegendEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LegendEntries, PowerPoint.Contrib.Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILegendEntry WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LegendEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LegendEntry, PowerPoint.Contrib.Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ILegendKey WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LegendKey resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LegendKey, PowerPoint.Contrib.Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPlotArea WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PlotArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PlotArea, PowerPoint.Contrib.Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPoint WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Point resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Point, PowerPoint.Contrib.Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPoints WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Points resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Points, PowerPoint.Contrib.Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISeries WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Series resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Series, PowerPoint.Contrib.Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISeriesCollection WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SeriesCollection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SeriesCollection, PowerPoint.Contrib.Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISeriesLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SeriesLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SeriesLines, PowerPoint.Contrib.Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITickLabels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TickLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TickLabels, PowerPoint.Contrib.Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITrendline WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Trendline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Trendline, PowerPoint.Contrib.Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ITrendlines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Trendlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Trendlines, PowerPoint.Contrib.Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IUpBars WithComCleanup(this Microsoft.Office.Interop.PowerPoint.UpBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.UpBars, PowerPoint.Contrib.Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IWalls WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Walls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Walls, PowerPoint.Contrib.Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for MediaFormat which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMediaFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MediaFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MediaFormat, PowerPoint.Contrib.Interfaces.IMediaFormat>();
		}

		/// <summary>
		/// Wrapper interface for SectionProperties which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ISectionProperties WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SectionProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SectionProperties, PowerPoint.Contrib.Interfaces.ISectionProperties>();
		}

		/// <summary>
		/// Wrapper interface for Player which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IPlayer WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Player resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Player, PowerPoint.Contrib.Interfaces.IPlayer>();
		}

		/// <summary>
		/// Wrapper interface for ResampleMediaTask which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IResampleMediaTask WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ResampleMediaTask resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ResampleMediaTask, PowerPoint.Contrib.Interfaces.IResampleMediaTask>();
		}

		/// <summary>
		/// Wrapper interface for ResampleMediaTasks which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IResampleMediaTasks WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks, PowerPoint.Contrib.Interfaces.IResampleMediaTasks>();
		}

		/// <summary>
		/// Wrapper interface for MediaBookmark which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMediaBookmark WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MediaBookmark resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MediaBookmark, PowerPoint.Contrib.Interfaces.IMediaBookmark>();
		}

		/// <summary>
		/// Wrapper interface for MediaBookmarks which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IMediaBookmarks WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MediaBookmarks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MediaBookmarks, PowerPoint.Contrib.Interfaces.IMediaBookmarks>();
		}

		/// <summary>
		/// Wrapper interface for Coauthoring which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.ICoauthoring WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Coauthoring resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Coauthoring, PowerPoint.Contrib.Interfaces.ICoauthoring>();
		}

		/// <summary>
		/// Wrapper interface for Broadcast which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IBroadcast WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Broadcast resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Broadcast, PowerPoint.Contrib.Interfaces.IBroadcast>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IProtectedViewWindows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows, PowerPoint.Contrib.Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static PowerPoint.Contrib.Interfaces.IProtectedViewWindow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow, PowerPoint.Contrib.Interfaces.IProtectedViewWindow>();
		}

	}
}