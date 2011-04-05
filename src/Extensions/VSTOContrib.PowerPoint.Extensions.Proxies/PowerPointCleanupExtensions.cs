using VSTOContrib.Extensions.Proxies;

//Microsoft.Office.Interop.PowerPoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.PowerPoint.Extensions.Proxies
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for Collection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICollection WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Collection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Collection, Interfaces.ICollection>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Application WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint._Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint._Application, Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Global WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint._Global resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint._Global, Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for EApplication_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEApplication_Event WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.EApplication_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.EApplication_Event, Interfaces.IEApplication_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplication WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Application resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Application, Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlobal WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Global resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Global, Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ColorFormat, Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowWindow WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SlideShowWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SlideShowWindow, Interfaces.ISlideShowWindow>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISelection WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Selection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Selection, Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for DocumentWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentWindows WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DocumentWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DocumentWindows, Interfaces.IDocumentWindows>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowWindows WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SlideShowWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SlideShowWindows, Interfaces.ISlideShowWindows>();
		}

		/// <summary>
		/// Wrapper interface for DocumentWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentWindow WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DocumentWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DocumentWindow, Interfaces.IDocumentWindow>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IView WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.View resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.View, Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowView WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SlideShowView resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SlideShowView, Interfaces.ISlideShowView>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowSettings WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SlideShowSettings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SlideShowSettings, Interfaces.ISlideShowSettings>();
		}

		/// <summary>
		/// Wrapper interface for NamedSlideShows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INamedSlideShows WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.NamedSlideShows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.NamedSlideShows, Interfaces.INamedSlideShows>();
		}

		/// <summary>
		/// Wrapper interface for NamedSlideShow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INamedSlideShow WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.NamedSlideShow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.NamedSlideShow, Interfaces.INamedSlideShow>();
		}

		/// <summary>
		/// Wrapper interface for PrintOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPrintOptions WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PrintOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PrintOptions, Interfaces.IPrintOptions>();
		}

		/// <summary>
		/// Wrapper interface for PrintRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPrintRanges WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PrintRanges resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PrintRanges, Interfaces.IPrintRanges>();
		}

		/// <summary>
		/// Wrapper interface for PrintRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPrintRange WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PrintRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PrintRange, Interfaces.IPrintRange>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIns WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AddIns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AddIns, Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIn WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AddIn resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AddIn, Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for Presentations which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresentations WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Presentations resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Presentations, Interfaces.IPresentations>();
		}

		/// <summary>
		/// Wrapper interface for PresEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresEvents WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PresEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PresEvents, Interfaces.IPresEvents>();
		}

		/// <summary>
		/// Wrapper interface for PresEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PresEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PresEvents_Event, Interfaces.IPresEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Presentation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresentation WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Presentation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Presentation, Interfaces.IPresentation>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlinks WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Hyperlinks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Hyperlinks, Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlink WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Hyperlink resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Hyperlink, Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageSetup WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PageSetup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PageSetup, Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for Fonts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFonts WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Fonts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Fonts, Interfaces.IFonts>();
		}

		/// <summary>
		/// Wrapper interface for ExtraColors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExtraColors WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ExtraColors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ExtraColors, Interfaces.IExtraColors>();
		}

		/// <summary>
		/// Wrapper interface for Slides which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlides WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Slides resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Slides, Interfaces.ISlides>();
		}

		/// <summary>
		/// Wrapper interface for _Slide which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Slide WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint._Slide resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint._Slide, Interfaces.I_Slide>();
		}

		/// <summary>
		/// Wrapper interface for SlideRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideRange WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SlideRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SlideRange, Interfaces.ISlideRange>();
		}

		/// <summary>
		/// Wrapper interface for _Master which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Master WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint._Master resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint._Master, Interfaces.I_Master>();
		}

		/// <summary>
		/// Wrapper interface for SldEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISldEvents WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SldEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SldEvents, Interfaces.ISldEvents>();
		}

		/// <summary>
		/// Wrapper interface for SldEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISldEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SldEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SldEvents_Event, Interfaces.ISldEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Slide which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlide WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Slide resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Slide, Interfaces.ISlide>();
		}

		/// <summary>
		/// Wrapper interface for ColorSchemes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorSchemes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ColorSchemes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ColorSchemes, Interfaces.IColorSchemes>();
		}

		/// <summary>
		/// Wrapper interface for ColorScheme which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorScheme WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ColorScheme resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ColorScheme, Interfaces.IColorScheme>();
		}

		/// <summary>
		/// Wrapper interface for RGBColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRGBColor WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.RGBColor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.RGBColor, Interfaces.IRGBColor>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowTransition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowTransition WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SlideShowTransition resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SlideShowTransition, Interfaces.ISlideShowTransition>();
		}

		/// <summary>
		/// Wrapper interface for SoundEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoundEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SoundEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SoundEffect, Interfaces.ISoundEffect>();
		}

		/// <summary>
		/// Wrapper interface for SoundFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoundFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SoundFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SoundFormat, Interfaces.ISoundFormat>();
		}

		/// <summary>
		/// Wrapper interface for HeadersFooters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadersFooters WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.HeadersFooters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.HeadersFooters, Interfaces.IHeadersFooters>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Shapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Shapes, Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for Placeholders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaceholders WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Placeholders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Placeholders, Interfaces.IPlaceholders>();
		}

		/// <summary>
		/// Wrapper interface for PlaceholderFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaceholderFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PlaceholderFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PlaceholderFormat, Interfaces.IPlaceholderFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFreeformBuilder WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.FreeformBuilder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.FreeformBuilder, Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShape WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Shape resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Shape, Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeRange WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ShapeRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ShapeRange, Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupShapes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.GroupShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.GroupShapes, Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAdjustments WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Adjustments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Adjustments, Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PictureFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PictureFormat, Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.FillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.FillFormat, Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.LineFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.LineFormat, Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShadowFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ShadowFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ShadowFormat, Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnectorFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ConnectorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ConnectorFormat, Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextEffectFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextEffectFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextEffectFormat, Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThreeDFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ThreeDFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ThreeDFormat, Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextFrame resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextFrame, Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalloutFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CalloutFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CalloutFormat, Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNodes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ShapeNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ShapeNodes, Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNode WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ShapeNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ShapeNode, Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.OLEFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.OLEFormat, Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinkFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.LinkFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.LinkFormat, Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for ObjectVerbs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IObjectVerbs WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ObjectVerbs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ObjectVerbs, Interfaces.IObjectVerbs>();
		}

		/// <summary>
		/// Wrapper interface for AnimationSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationSettings WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AnimationSettings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AnimationSettings, Interfaces.IAnimationSettings>();
		}

		/// <summary>
		/// Wrapper interface for ActionSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IActionSettings WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ActionSettings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ActionSettings, Interfaces.IActionSettings>();
		}

		/// <summary>
		/// Wrapper interface for ActionSetting which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IActionSetting WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ActionSetting resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ActionSetting, Interfaces.IActionSetting>();
		}

		/// <summary>
		/// Wrapper interface for PlaySettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaySettings WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PlaySettings resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PlaySettings, Interfaces.IPlaySettings>();
		}

		/// <summary>
		/// Wrapper interface for TextRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextRange WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextRange, Interfaces.ITextRange>();
		}

		/// <summary>
		/// Wrapper interface for Ruler which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuler WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Ruler resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Ruler, Interfaces.IRuler>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRulerLevels WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.RulerLevels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.RulerLevels, Interfaces.IRulerLevels>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRulerLevel WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.RulerLevel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.RulerLevel, Interfaces.IRulerLevel>();
		}

		/// <summary>
		/// Wrapper interface for TabStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStops WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TabStops resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TabStops, Interfaces.ITabStops>();
		}

		/// <summary>
		/// Wrapper interface for TabStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStop WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TabStop resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TabStop, Interfaces.ITabStop>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFont WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Font resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Font, Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ParagraphFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ParagraphFormat, Interfaces.IParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for BulletFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBulletFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.BulletFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.BulletFormat, Interfaces.IBulletFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyles WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextStyles resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextStyles, Interfaces.ITextStyles>();
		}

		/// <summary>
		/// Wrapper interface for TextStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyle WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextStyle, Interfaces.ITextStyle>();
		}

		/// <summary>
		/// Wrapper interface for TextStyleLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyleLevels WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextStyleLevels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextStyleLevels, Interfaces.ITextStyleLevels>();
		}

		/// <summary>
		/// Wrapper interface for TextStyleLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyleLevel WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextStyleLevel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextStyleLevel, Interfaces.ITextStyleLevel>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeaderFooter WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.HeaderFooter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.HeaderFooter, Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for _Presentation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Presentation WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint._Presentation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint._Presentation, Interfaces.I_Presentation>();
		}

		/// <summary>
		/// Wrapper interface for Tags which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITags WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Tags resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Tags, Interfaces.ITags>();
		}

		/// <summary>
		/// Wrapper interface for MouseTracker which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMouseTracker WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MouseTracker resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MouseTracker, Interfaces.IMouseTracker>();
		}

		/// <summary>
		/// Wrapper interface for MouseDownHandler which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMouseDownHandler WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MouseDownHandler resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MouseDownHandler, Interfaces.IMouseDownHandler>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtender which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXExtender WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.OCXExtender resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.OCXExtender, Interfaces.IOCXExtender>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtenderEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXExtenderEvents WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents, Interfaces.IOCXExtenderEvents>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtenderEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXExtenderEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event, Interfaces.IOCXExtenderEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEControl WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.OLEControl resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.OLEControl, Interfaces.IOLEControl>();
		}

		/// <summary>
		/// Wrapper interface for EApplication which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEApplication WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.EApplication resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.EApplication, Interfaces.IEApplication>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITable WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Table resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Table, Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumns WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Columns resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Columns, Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumn WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Column resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Column, Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for Rows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRows WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Rows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Rows, Interfaces.IRows>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRow WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Row resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Row, Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for CellRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICellRange WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CellRange resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CellRange, Interfaces.ICellRange>();
		}

		/// <summary>
		/// Wrapper interface for Cell which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICell WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Cell resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Cell, Interfaces.ICell>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorders WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Borders resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Borders, Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPanes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Panes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Panes, Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPane WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Pane resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Pane, Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDefaultWebOptions WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DefaultWebOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DefaultWebOptions, Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebOptions WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.WebOptions resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.WebOptions, Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for PublishObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPublishObjects WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PublishObjects resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PublishObjects, Interfaces.IPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for PublishObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPublishObject WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PublishObject resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PublishObject, Interfaces.IPublishObject>();
		}

		/// <summary>
		/// Wrapper interface for MasterEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMasterEvents WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MasterEvents resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MasterEvents, Interfaces.IMasterEvents>();
		}

		/// <summary>
		/// Wrapper interface for MasterEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMasterEvents_Event WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MasterEvents_Event resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MasterEvents_Event, Interfaces.IMasterEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Master which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMaster WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Master resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Master, Interfaces.IMaster>();
		}

		/// <summary>
		/// Wrapper interface for _PowerRex which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_PowerRex WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint._PowerRex resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint._PowerRex, Interfaces.I_PowerRex>();
		}

		/// <summary>
		/// Wrapper interface for PowerRex which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPowerRex WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PowerRex resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PowerRex, Interfaces.IPowerRex>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComments WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Comments resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Comments, Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComment WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Comment resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Comment, Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for Designs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDesigns WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Designs resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Designs, Interfaces.IDesigns>();
		}

		/// <summary>
		/// Wrapper interface for Design which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDesign WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Design resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Design, Interfaces.IDesign>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNode WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DiagramNode resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DiagramNode, Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodeChildren WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren, Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DiagramNodes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DiagramNodes, Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagram WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Diagram resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Diagram, Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for TimeLine which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITimeLine WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TimeLine resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TimeLine, Interfaces.ITimeLine>();
		}

		/// <summary>
		/// Wrapper interface for Sequences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISequences WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Sequences resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Sequences, Interfaces.ISequences>();
		}

		/// <summary>
		/// Wrapper interface for Sequence which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISequence WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Sequence resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Sequence, Interfaces.ISequence>();
		}

		/// <summary>
		/// Wrapper interface for Effect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Effect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Effect, Interfaces.IEffect>();
		}

		/// <summary>
		/// Wrapper interface for Timing which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITiming WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Timing resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Timing, Interfaces.ITiming>();
		}

		/// <summary>
		/// Wrapper interface for EffectParameters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffectParameters WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.EffectParameters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.EffectParameters, Interfaces.IEffectParameters>();
		}

		/// <summary>
		/// Wrapper interface for EffectInformation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffectInformation WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.EffectInformation resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.EffectInformation, Interfaces.IEffectInformation>();
		}

		/// <summary>
		/// Wrapper interface for AnimationBehaviors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationBehaviors WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AnimationBehaviors resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AnimationBehaviors, Interfaces.IAnimationBehaviors>();
		}

		/// <summary>
		/// Wrapper interface for AnimationBehavior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationBehavior WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AnimationBehavior resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AnimationBehavior, Interfaces.IAnimationBehavior>();
		}

		/// <summary>
		/// Wrapper interface for MotionEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMotionEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MotionEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MotionEffect, Interfaces.IMotionEffect>();
		}

		/// <summary>
		/// Wrapper interface for ColorEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ColorEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ColorEffect, Interfaces.IColorEffect>();
		}

		/// <summary>
		/// Wrapper interface for ScaleEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScaleEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ScaleEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ScaleEffect, Interfaces.IScaleEffect>();
		}

		/// <summary>
		/// Wrapper interface for RotationEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRotationEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.RotationEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.RotationEffect, Interfaces.IRotationEffect>();
		}

		/// <summary>
		/// Wrapper interface for PropertyEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PropertyEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PropertyEffect, Interfaces.IPropertyEffect>();
		}

		/// <summary>
		/// Wrapper interface for AnimationPoints which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationPoints WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AnimationPoints resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AnimationPoints, Interfaces.IAnimationPoints>();
		}

		/// <summary>
		/// Wrapper interface for AnimationPoint which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationPoint WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AnimationPoint resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AnimationPoint, Interfaces.IAnimationPoint>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICanvasShapes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CanvasShapes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CanvasShapes, Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AutoCorrect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AutoCorrect, Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for Options which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOptions WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Options resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Options, Interfaces.IOptions>();
		}

		/// <summary>
		/// Wrapper interface for CommandEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CommandEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CommandEffect, Interfaces.ICommandEffect>();
		}

		/// <summary>
		/// Wrapper interface for FilterEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFilterEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.FilterEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.FilterEffect, Interfaces.IFilterEffect>();
		}

		/// <summary>
		/// Wrapper interface for SetEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISetEffect WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SetEffect resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SetEffect, Interfaces.ISetEffect>();
		}

		/// <summary>
		/// Wrapper interface for CustomLayouts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLayouts WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CustomLayouts resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CustomLayouts, Interfaces.ICustomLayouts>();
		}

		/// <summary>
		/// Wrapper interface for CustomLayout which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLayout WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CustomLayout resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CustomLayout, Interfaces.ICustomLayout>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyle WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TableStyle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TableStyle, Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for CustomerData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomerData WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.CustomerData resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.CustomerData, Interfaces.ICustomerData>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResearch WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Research resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Research, Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for TableBackground which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableBackground WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TableBackground resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TableBackground, Interfaces.ITableBackground>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame2 WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TextFrame2 resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TextFrame2, Interfaces.ITextFrame2>();
		}

		/// <summary>
		/// Wrapper interface for FileConverters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverters WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.FileConverters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.FileConverters, Interfaces.IFileConverters>();
		}

		/// <summary>
		/// Wrapper interface for FileConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverter WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.FileConverter resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.FileConverter, Interfaces.IFileConverter>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxes WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Axes resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Axes, Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxis WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Axis resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Axis, Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxisTitle WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.AxisTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.AxisTitle, Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChart WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Chart resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Chart, Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for ChartBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartBorder WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartBorder resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartBorder, Interfaces.IChartBorder>();
		}

		/// <summary>
		/// Wrapper interface for ChartCharacters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartCharacters WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartCharacters resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartCharacters, Interfaces.IChartCharacters>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartArea WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartArea, Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartColorFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartColorFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartColorFormat, Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartData WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartData resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartData, Interfaces.IChartData>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFillFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartFillFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartFillFormat, Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartFormat, Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroup WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartGroup resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartGroup, Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroups WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartGroups resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartGroups, Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartTitle WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartTitle resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartTitle, Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICorners WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Corners resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Corners, Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabel WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DataLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DataLabel, Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabels WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DataLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DataLabels, Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataTable WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DataTable resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DataTable, Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDisplayUnitLabel WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel, Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDownBars WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DownBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DownBars, Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropLines WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.DropLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.DropLines, Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrorBars WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ErrorBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ErrorBars, Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFloor WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Floor resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Floor, Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFont WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ChartFont resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ChartFont, Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGridlines WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Gridlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Gridlines, Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHiLoLines WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.HiLoLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.HiLoLines, Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInterior WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Interior resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Interior, Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILeaderLines WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.LeaderLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.LeaderLines, Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegend WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Legend resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Legend, Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntries WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.LegendEntries resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.LegendEntries, Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntry WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.LegendEntry resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.LegendEntry, Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendKey WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.LegendKey resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.LegendKey, Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlotArea WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.PlotArea resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.PlotArea, Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoint WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Point resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Point, Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoints WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Points resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Points, Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeries WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Series resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Series, Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesCollection WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SeriesCollection resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SeriesCollection, Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesLines WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SeriesLines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SeriesLines, Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITickLabels WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.TickLabels resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.TickLabels, Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendline WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Trendline resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Trendline, Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendlines WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Trendlines resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Trendlines, Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUpBars WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.UpBars resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.UpBars, Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWalls WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Walls resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Walls, Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for MediaFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMediaFormat WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MediaFormat resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MediaFormat, Interfaces.IMediaFormat>();
		}

		/// <summary>
		/// Wrapper interface for SectionProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISectionProperties WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.SectionProperties resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.SectionProperties, Interfaces.ISectionProperties>();
		}

		/// <summary>
		/// Wrapper interface for Player which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlayer WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Player resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Player, Interfaces.IPlayer>();
		}

		/// <summary>
		/// Wrapper interface for ResampleMediaTask which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResampleMediaTask WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ResampleMediaTask resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ResampleMediaTask, Interfaces.IResampleMediaTask>();
		}

		/// <summary>
		/// Wrapper interface for ResampleMediaTasks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResampleMediaTasks WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks, Interfaces.IResampleMediaTasks>();
		}

		/// <summary>
		/// Wrapper interface for MediaBookmark which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMediaBookmark WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MediaBookmark resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MediaBookmark, Interfaces.IMediaBookmark>();
		}

		/// <summary>
		/// Wrapper interface for MediaBookmarks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMediaBookmarks WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.MediaBookmarks resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.MediaBookmarks, Interfaces.IMediaBookmarks>();
		}

		/// <summary>
		/// Wrapper interface for Coauthoring which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoauthoring WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Coauthoring resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Coauthoring, Interfaces.ICoauthoring>();
		}

		/// <summary>
		/// Wrapper interface for Broadcast which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBroadcast WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.Broadcast resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.Broadcast, Interfaces.IBroadcast>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindows WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows, Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindow WithComCleanupProxy(this Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow resource)
		{
			return resource.WithComCleanupProxy<Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow, Interfaces.IProtectedViewWindow>();
		}

	}
}