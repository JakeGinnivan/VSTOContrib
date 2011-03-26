using Office.Contrib.Extensions;

namespace Office.PowerPoint.Contrib.Extensions
{
	/// <summary>
	/// Provides cleanup extension methods for interfaces exposed by Office.dll
	/// </summary>
	public static class OfficeCleanupExtensions
	{
		/// <summary>
		/// Wrapper interface for Collection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICollection WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Collection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Collection, Interfaces.ICollection>();
		}

		/// <summary>
		/// Wrapper interface for _Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Application WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Application, Interfaces.I_Application>();
		}

		/// <summary>
		/// Wrapper interface for _Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Global WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Global, Interfaces.I_Global>();
		}

		/// <summary>
		/// Wrapper interface for EApplication_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEApplication_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EApplication_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EApplication_Event, Interfaces.IEApplication_Event>();
		}

		/// <summary>
		/// Wrapper interface for Application which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IApplication WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Application resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Application, Interfaces.IApplication>();
		}

		/// <summary>
		/// Wrapper interface for Global which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGlobal WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Global resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Global, Interfaces.IGlobal>();
		}

		/// <summary>
		/// Wrapper interface for ColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorFormat, Interfaces.IColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowWindow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowWindow, Interfaces.ISlideShowWindow>();
		}

		/// <summary>
		/// Wrapper interface for Selection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISelection WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Selection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Selection, Interfaces.ISelection>();
		}

		/// <summary>
		/// Wrapper interface for DocumentWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentWindows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DocumentWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DocumentWindows, Interfaces.IDocumentWindows>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowWindows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowWindows, Interfaces.ISlideShowWindows>();
		}

		/// <summary>
		/// Wrapper interface for DocumentWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDocumentWindow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DocumentWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DocumentWindow, Interfaces.IDocumentWindow>();
		}

		/// <summary>
		/// Wrapper interface for View which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IView WithComCleanup(this Microsoft.Office.Interop.PowerPoint.View resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.View, Interfaces.IView>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowView which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowView WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowView resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowView, Interfaces.ISlideShowView>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowSettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowSettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowSettings, Interfaces.ISlideShowSettings>();
		}

		/// <summary>
		/// Wrapper interface for NamedSlideShows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INamedSlideShows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.NamedSlideShows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.NamedSlideShows, Interfaces.INamedSlideShows>();
		}

		/// <summary>
		/// Wrapper interface for NamedSlideShow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.INamedSlideShow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.NamedSlideShow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.NamedSlideShow, Interfaces.INamedSlideShow>();
		}

		/// <summary>
		/// Wrapper interface for PrintOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPrintOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PrintOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PrintOptions, Interfaces.IPrintOptions>();
		}

		/// <summary>
		/// Wrapper interface for PrintRanges which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPrintRanges WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PrintRanges resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PrintRanges, Interfaces.IPrintRanges>();
		}

		/// <summary>
		/// Wrapper interface for PrintRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPrintRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PrintRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PrintRange, Interfaces.IPrintRange>();
		}

		/// <summary>
		/// Wrapper interface for AddIns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIns WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AddIns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AddIns, Interfaces.IAddIns>();
		}

		/// <summary>
		/// Wrapper interface for AddIn which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAddIn WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AddIn resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AddIn, Interfaces.IAddIn>();
		}

		/// <summary>
		/// Wrapper interface for Presentations which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresentations WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Presentations resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Presentations, Interfaces.IPresentations>();
		}

		/// <summary>
		/// Wrapper interface for PresEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PresEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PresEvents, Interfaces.IPresEvents>();
		}

		/// <summary>
		/// Wrapper interface for PresEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PresEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PresEvents_Event, Interfaces.IPresEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Presentation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPresentation WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Presentation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Presentation, Interfaces.IPresentation>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlinks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlinks WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Hyperlinks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Hyperlinks, Interfaces.IHyperlinks>();
		}

		/// <summary>
		/// Wrapper interface for Hyperlink which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHyperlink WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Hyperlink resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Hyperlink, Interfaces.IHyperlink>();
		}

		/// <summary>
		/// Wrapper interface for PageSetup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPageSetup WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PageSetup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PageSetup, Interfaces.IPageSetup>();
		}

		/// <summary>
		/// Wrapper interface for Fonts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFonts WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Fonts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Fonts, Interfaces.IFonts>();
		}

		/// <summary>
		/// Wrapper interface for ExtraColors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IExtraColors WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ExtraColors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ExtraColors, Interfaces.IExtraColors>();
		}

		/// <summary>
		/// Wrapper interface for Slides which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlides WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Slides resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Slides, Interfaces.ISlides>();
		}

		/// <summary>
		/// Wrapper interface for _Slide which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Slide WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Slide resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Slide, Interfaces.I_Slide>();
		}

		/// <summary>
		/// Wrapper interface for SlideRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideRange, Interfaces.ISlideRange>();
		}

		/// <summary>
		/// Wrapper interface for _Master which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Master WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Master resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Master, Interfaces.I_Master>();
		}

		/// <summary>
		/// Wrapper interface for SldEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISldEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SldEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SldEvents, Interfaces.ISldEvents>();
		}

		/// <summary>
		/// Wrapper interface for SldEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISldEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SldEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SldEvents_Event, Interfaces.ISldEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Slide which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlide WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Slide resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Slide, Interfaces.ISlide>();
		}

		/// <summary>
		/// Wrapper interface for ColorSchemes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorSchemes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorSchemes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorSchemes, Interfaces.IColorSchemes>();
		}

		/// <summary>
		/// Wrapper interface for ColorScheme which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorScheme WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorScheme resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorScheme, Interfaces.IColorScheme>();
		}

		/// <summary>
		/// Wrapper interface for RGBColor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRGBColor WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RGBColor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RGBColor, Interfaces.IRGBColor>();
		}

		/// <summary>
		/// Wrapper interface for SlideShowTransition which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISlideShowTransition WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SlideShowTransition resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SlideShowTransition, Interfaces.ISlideShowTransition>();
		}

		/// <summary>
		/// Wrapper interface for SoundEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoundEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SoundEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SoundEffect, Interfaces.ISoundEffect>();
		}

		/// <summary>
		/// Wrapper interface for SoundFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISoundFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SoundFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SoundFormat, Interfaces.ISoundFormat>();
		}

		/// <summary>
		/// Wrapper interface for HeadersFooters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeadersFooters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.HeadersFooters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.HeadersFooters, Interfaces.IHeadersFooters>();
		}

		/// <summary>
		/// Wrapper interface for Shapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Shapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Shapes, Interfaces.IShapes>();
		}

		/// <summary>
		/// Wrapper interface for Placeholders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaceholders WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Placeholders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Placeholders, Interfaces.IPlaceholders>();
		}

		/// <summary>
		/// Wrapper interface for PlaceholderFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaceholderFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PlaceholderFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PlaceholderFormat, Interfaces.IPlaceholderFormat>();
		}

		/// <summary>
		/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFreeformBuilder WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FreeformBuilder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FreeformBuilder, Interfaces.IFreeformBuilder>();
		}

		/// <summary>
		/// Wrapper interface for Shape which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShape WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Shape resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Shape, Interfaces.IShape>();
		}

		/// <summary>
		/// Wrapper interface for ShapeRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShapeRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShapeRange, Interfaces.IShapeRange>();
		}

		/// <summary>
		/// Wrapper interface for GroupShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGroupShapes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.GroupShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.GroupShapes, Interfaces.IGroupShapes>();
		}

		/// <summary>
		/// Wrapper interface for Adjustments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAdjustments WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Adjustments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Adjustments, Interfaces.IAdjustments>();
		}

		/// <summary>
		/// Wrapper interface for PictureFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPictureFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PictureFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PictureFormat, Interfaces.IPictureFormat>();
		}

		/// <summary>
		/// Wrapper interface for FillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFillFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FillFormat, Interfaces.IFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for LineFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILineFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LineFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LineFormat, Interfaces.ILineFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShadowFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShadowFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShadowFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShadowFormat, Interfaces.IShadowFormat>();
		}

		/// <summary>
		/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IConnectorFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ConnectorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ConnectorFormat, Interfaces.IConnectorFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextEffectFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextEffectFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextEffectFormat, Interfaces.ITextEffectFormat>();
		}

		/// <summary>
		/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IThreeDFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ThreeDFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ThreeDFormat, Interfaces.IThreeDFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextFrame resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextFrame, Interfaces.ITextFrame>();
		}

		/// <summary>
		/// Wrapper interface for CalloutFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICalloutFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CalloutFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CalloutFormat, Interfaces.ICalloutFormat>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNodes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShapeNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShapeNodes, Interfaces.IShapeNodes>();
		}

		/// <summary>
		/// Wrapper interface for ShapeNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IShapeNode WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ShapeNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ShapeNode, Interfaces.IShapeNode>();
		}

		/// <summary>
		/// Wrapper interface for OLEFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OLEFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OLEFormat, Interfaces.IOLEFormat>();
		}

		/// <summary>
		/// Wrapper interface for LinkFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILinkFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LinkFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LinkFormat, Interfaces.ILinkFormat>();
		}

		/// <summary>
		/// Wrapper interface for ObjectVerbs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IObjectVerbs WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ObjectVerbs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ObjectVerbs, Interfaces.IObjectVerbs>();
		}

		/// <summary>
		/// Wrapper interface for AnimationSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationSettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationSettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationSettings, Interfaces.IAnimationSettings>();
		}

		/// <summary>
		/// Wrapper interface for ActionSettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IActionSettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ActionSettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ActionSettings, Interfaces.IActionSettings>();
		}

		/// <summary>
		/// Wrapper interface for ActionSetting which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IActionSetting WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ActionSetting resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ActionSetting, Interfaces.IActionSetting>();
		}

		/// <summary>
		/// Wrapper interface for PlaySettings which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlaySettings WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PlaySettings resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PlaySettings, Interfaces.IPlaySettings>();
		}

		/// <summary>
		/// Wrapper interface for TextRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextRange, Interfaces.ITextRange>();
		}

		/// <summary>
		/// Wrapper interface for Ruler which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRuler WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Ruler resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Ruler, Interfaces.IRuler>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRulerLevels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RulerLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RulerLevels, Interfaces.IRulerLevels>();
		}

		/// <summary>
		/// Wrapper interface for RulerLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRulerLevel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RulerLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RulerLevel, Interfaces.IRulerLevel>();
		}

		/// <summary>
		/// Wrapper interface for TabStops which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStops WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TabStops resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TabStops, Interfaces.ITabStops>();
		}

		/// <summary>
		/// Wrapper interface for TabStop which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITabStop WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TabStop resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TabStop, Interfaces.ITabStop>();
		}

		/// <summary>
		/// Wrapper interface for Font which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFont WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Font resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Font, Interfaces.IFont>();
		}

		/// <summary>
		/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IParagraphFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ParagraphFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ParagraphFormat, Interfaces.IParagraphFormat>();
		}

		/// <summary>
		/// Wrapper interface for BulletFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBulletFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.BulletFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.BulletFormat, Interfaces.IBulletFormat>();
		}

		/// <summary>
		/// Wrapper interface for TextStyles which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyles WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyles resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyles, Interfaces.ITextStyles>();
		}

		/// <summary>
		/// Wrapper interface for TextStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyle, Interfaces.ITextStyle>();
		}

		/// <summary>
		/// Wrapper interface for TextStyleLevels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyleLevels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyleLevels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyleLevels, Interfaces.ITextStyleLevels>();
		}

		/// <summary>
		/// Wrapper interface for TextStyleLevel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextStyleLevel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextStyleLevel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextStyleLevel, Interfaces.ITextStyleLevel>();
		}

		/// <summary>
		/// Wrapper interface for HeaderFooter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHeaderFooter WithComCleanup(this Microsoft.Office.Interop.PowerPoint.HeaderFooter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.HeaderFooter, Interfaces.IHeaderFooter>();
		}

		/// <summary>
		/// Wrapper interface for _Presentation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_Presentation WithComCleanup(this Microsoft.Office.Interop.PowerPoint._Presentation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._Presentation, Interfaces.I_Presentation>();
		}

		/// <summary>
		/// Wrapper interface for Tags which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITags WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Tags resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Tags, Interfaces.ITags>();
		}

		/// <summary>
		/// Wrapper interface for MouseTracker which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMouseTracker WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MouseTracker resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MouseTracker, Interfaces.IMouseTracker>();
		}

		/// <summary>
		/// Wrapper interface for MouseDownHandler which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMouseDownHandler WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MouseDownHandler resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MouseDownHandler, Interfaces.IMouseDownHandler>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtender which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXExtender WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OCXExtender resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OCXExtender, Interfaces.IOCXExtender>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtenderEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXExtenderEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents, Interfaces.IOCXExtenderEvents>();
		}

		/// <summary>
		/// Wrapper interface for OCXExtenderEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOCXExtenderEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event, Interfaces.IOCXExtenderEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for OLEControl which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOLEControl WithComCleanup(this Microsoft.Office.Interop.PowerPoint.OLEControl resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.OLEControl, Interfaces.IOLEControl>();
		}

		/// <summary>
		/// Wrapper interface for EApplication which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEApplication WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EApplication resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EApplication, Interfaces.IEApplication>();
		}

		/// <summary>
		/// Wrapper interface for Table which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITable WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Table resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Table, Interfaces.ITable>();
		}

		/// <summary>
		/// Wrapper interface for Columns which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumns WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Columns resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Columns, Interfaces.IColumns>();
		}

		/// <summary>
		/// Wrapper interface for Column which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColumn WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Column resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Column, Interfaces.IColumn>();
		}

		/// <summary>
		/// Wrapper interface for Rows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Rows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Rows, Interfaces.IRows>();
		}

		/// <summary>
		/// Wrapper interface for Row which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Row resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Row, Interfaces.IRow>();
		}

		/// <summary>
		/// Wrapper interface for CellRange which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICellRange WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CellRange resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CellRange, Interfaces.ICellRange>();
		}

		/// <summary>
		/// Wrapper interface for Cell which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICell WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Cell resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Cell, Interfaces.ICell>();
		}

		/// <summary>
		/// Wrapper interface for Borders which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBorders WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Borders resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Borders, Interfaces.IBorders>();
		}

		/// <summary>
		/// Wrapper interface for Panes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPanes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Panes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Panes, Interfaces.IPanes>();
		}

		/// <summary>
		/// Wrapper interface for Pane which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPane WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Pane resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Pane, Interfaces.IPane>();
		}

		/// <summary>
		/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDefaultWebOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DefaultWebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DefaultWebOptions, Interfaces.IDefaultWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for WebOptions which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWebOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.WebOptions resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.WebOptions, Interfaces.IWebOptions>();
		}

		/// <summary>
		/// Wrapper interface for PublishObjects which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPublishObjects WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PublishObjects resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PublishObjects, Interfaces.IPublishObjects>();
		}

		/// <summary>
		/// Wrapper interface for PublishObject which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPublishObject WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PublishObject resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PublishObject, Interfaces.IPublishObject>();
		}

		/// <summary>
		/// Wrapper interface for MasterEvents which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMasterEvents WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MasterEvents resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MasterEvents, Interfaces.IMasterEvents>();
		}

		/// <summary>
		/// Wrapper interface for MasterEvents_Event which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMasterEvents_Event WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MasterEvents_Event resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MasterEvents_Event, Interfaces.IMasterEvents_Event>();
		}

		/// <summary>
		/// Wrapper interface for Master which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMaster WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Master resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Master, Interfaces.IMaster>();
		}

		/// <summary>
		/// Wrapper interface for _PowerRex which adds IDispose to the interface
		/// </summary>
		public static Interfaces.I_PowerRex WithComCleanup(this Microsoft.Office.Interop.PowerPoint._PowerRex resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint._PowerRex, Interfaces.I_PowerRex>();
		}

		/// <summary>
		/// Wrapper interface for PowerRex which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPowerRex WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PowerRex resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PowerRex, Interfaces.IPowerRex>();
		}

		/// <summary>
		/// Wrapper interface for Comments which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComments WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Comments resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Comments, Interfaces.IComments>();
		}

		/// <summary>
		/// Wrapper interface for Comment which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IComment WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Comment resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Comment, Interfaces.IComment>();
		}

		/// <summary>
		/// Wrapper interface for Designs which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDesigns WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Designs resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Designs, Interfaces.IDesigns>();
		}

		/// <summary>
		/// Wrapper interface for Design which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDesign WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Design resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Design, Interfaces.IDesign>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNode which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNode WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DiagramNode resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DiagramNode, Interfaces.IDiagramNode>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodeChildren WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren, Interfaces.IDiagramNodeChildren>();
		}

		/// <summary>
		/// Wrapper interface for DiagramNodes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagramNodes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DiagramNodes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DiagramNodes, Interfaces.IDiagramNodes>();
		}

		/// <summary>
		/// Wrapper interface for Diagram which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDiagram WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Diagram resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Diagram, Interfaces.IDiagram>();
		}

		/// <summary>
		/// Wrapper interface for TimeLine which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITimeLine WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TimeLine resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TimeLine, Interfaces.ITimeLine>();
		}

		/// <summary>
		/// Wrapper interface for Sequences which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISequences WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Sequences resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Sequences, Interfaces.ISequences>();
		}

		/// <summary>
		/// Wrapper interface for Sequence which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISequence WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Sequence resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Sequence, Interfaces.ISequence>();
		}

		/// <summary>
		/// Wrapper interface for Effect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Effect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Effect, Interfaces.IEffect>();
		}

		/// <summary>
		/// Wrapper interface for Timing which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITiming WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Timing resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Timing, Interfaces.ITiming>();
		}

		/// <summary>
		/// Wrapper interface for EffectParameters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffectParameters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EffectParameters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EffectParameters, Interfaces.IEffectParameters>();
		}

		/// <summary>
		/// Wrapper interface for EffectInformation which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IEffectInformation WithComCleanup(this Microsoft.Office.Interop.PowerPoint.EffectInformation resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.EffectInformation, Interfaces.IEffectInformation>();
		}

		/// <summary>
		/// Wrapper interface for AnimationBehaviors which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationBehaviors WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationBehaviors resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationBehaviors, Interfaces.IAnimationBehaviors>();
		}

		/// <summary>
		/// Wrapper interface for AnimationBehavior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationBehavior WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationBehavior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationBehavior, Interfaces.IAnimationBehavior>();
		}

		/// <summary>
		/// Wrapper interface for MotionEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMotionEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MotionEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MotionEffect, Interfaces.IMotionEffect>();
		}

		/// <summary>
		/// Wrapper interface for ColorEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IColorEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ColorEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ColorEffect, Interfaces.IColorEffect>();
		}

		/// <summary>
		/// Wrapper interface for ScaleEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IScaleEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ScaleEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ScaleEffect, Interfaces.IScaleEffect>();
		}

		/// <summary>
		/// Wrapper interface for RotationEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IRotationEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.RotationEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.RotationEffect, Interfaces.IRotationEffect>();
		}

		/// <summary>
		/// Wrapper interface for PropertyEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPropertyEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PropertyEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PropertyEffect, Interfaces.IPropertyEffect>();
		}

		/// <summary>
		/// Wrapper interface for AnimationPoints which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationPoints WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationPoints resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationPoints, Interfaces.IAnimationPoints>();
		}

		/// <summary>
		/// Wrapper interface for AnimationPoint which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAnimationPoint WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AnimationPoint resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AnimationPoint, Interfaces.IAnimationPoint>();
		}

		/// <summary>
		/// Wrapper interface for CanvasShapes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICanvasShapes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CanvasShapes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CanvasShapes, Interfaces.ICanvasShapes>();
		}

		/// <summary>
		/// Wrapper interface for AutoCorrect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAutoCorrect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AutoCorrect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AutoCorrect, Interfaces.IAutoCorrect>();
		}

		/// <summary>
		/// Wrapper interface for Options which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IOptions WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Options resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Options, Interfaces.IOptions>();
		}

		/// <summary>
		/// Wrapper interface for CommandEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICommandEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CommandEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CommandEffect, Interfaces.ICommandEffect>();
		}

		/// <summary>
		/// Wrapper interface for FilterEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFilterEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FilterEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FilterEffect, Interfaces.IFilterEffect>();
		}

		/// <summary>
		/// Wrapper interface for SetEffect which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISetEffect WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SetEffect resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SetEffect, Interfaces.ISetEffect>();
		}

		/// <summary>
		/// Wrapper interface for CustomLayouts which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLayouts WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CustomLayouts resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CustomLayouts, Interfaces.ICustomLayouts>();
		}

		/// <summary>
		/// Wrapper interface for CustomLayout which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomLayout WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CustomLayout resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CustomLayout, Interfaces.ICustomLayout>();
		}

		/// <summary>
		/// Wrapper interface for TableStyle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableStyle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TableStyle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TableStyle, Interfaces.ITableStyle>();
		}

		/// <summary>
		/// Wrapper interface for CustomerData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICustomerData WithComCleanup(this Microsoft.Office.Interop.PowerPoint.CustomerData resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.CustomerData, Interfaces.ICustomerData>();
		}

		/// <summary>
		/// Wrapper interface for Research which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResearch WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Research resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Research, Interfaces.IResearch>();
		}

		/// <summary>
		/// Wrapper interface for TableBackground which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITableBackground WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TableBackground resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TableBackground, Interfaces.ITableBackground>();
		}

		/// <summary>
		/// Wrapper interface for TextFrame2 which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITextFrame2 WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TextFrame2 resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TextFrame2, Interfaces.ITextFrame2>();
		}

		/// <summary>
		/// Wrapper interface for FileConverters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FileConverters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FileConverters, Interfaces.IFileConverters>();
		}

		/// <summary>
		/// Wrapper interface for FileConverter which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFileConverter WithComCleanup(this Microsoft.Office.Interop.PowerPoint.FileConverter resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.FileConverter, Interfaces.IFileConverter>();
		}

		/// <summary>
		/// Wrapper interface for Axes which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxes WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Axes resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Axes, Interfaces.IAxes>();
		}

		/// <summary>
		/// Wrapper interface for Axis which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxis WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Axis resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Axis, Interfaces.IAxis>();
		}

		/// <summary>
		/// Wrapper interface for AxisTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IAxisTitle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.AxisTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.AxisTitle, Interfaces.IAxisTitle>();
		}

		/// <summary>
		/// Wrapper interface for Chart which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChart WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Chart resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Chart, Interfaces.IChart>();
		}

		/// <summary>
		/// Wrapper interface for ChartBorder which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartBorder WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartBorder resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartBorder, Interfaces.IChartBorder>();
		}

		/// <summary>
		/// Wrapper interface for ChartCharacters which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartCharacters WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartCharacters resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartCharacters, Interfaces.IChartCharacters>();
		}

		/// <summary>
		/// Wrapper interface for ChartArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartArea WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartArea, Interfaces.IChartArea>();
		}

		/// <summary>
		/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartColorFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartColorFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartColorFormat, Interfaces.IChartColorFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartData which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartData WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartData resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartData, Interfaces.IChartData>();
		}

		/// <summary>
		/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFillFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartFillFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartFillFormat, Interfaces.IChartFillFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartFormat, Interfaces.IChartFormat>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroup which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroup WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartGroup resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartGroup, Interfaces.IChartGroup>();
		}

		/// <summary>
		/// Wrapper interface for ChartGroups which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartGroups WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartGroups resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartGroups, Interfaces.IChartGroups>();
		}

		/// <summary>
		/// Wrapper interface for ChartTitle which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartTitle WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartTitle resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartTitle, Interfaces.IChartTitle>();
		}

		/// <summary>
		/// Wrapper interface for Corners which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICorners WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Corners resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Corners, Interfaces.ICorners>();
		}

		/// <summary>
		/// Wrapper interface for DataLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DataLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DataLabel, Interfaces.IDataLabel>();
		}

		/// <summary>
		/// Wrapper interface for DataLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataLabels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DataLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DataLabels, Interfaces.IDataLabels>();
		}

		/// <summary>
		/// Wrapper interface for DataTable which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDataTable WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DataTable resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DataTable, Interfaces.IDataTable>();
		}

		/// <summary>
		/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDisplayUnitLabel WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel, Interfaces.IDisplayUnitLabel>();
		}

		/// <summary>
		/// Wrapper interface for DownBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDownBars WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DownBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DownBars, Interfaces.IDownBars>();
		}

		/// <summary>
		/// Wrapper interface for DropLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IDropLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.DropLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.DropLines, Interfaces.IDropLines>();
		}

		/// <summary>
		/// Wrapper interface for ErrorBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IErrorBars WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ErrorBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ErrorBars, Interfaces.IErrorBars>();
		}

		/// <summary>
		/// Wrapper interface for Floor which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IFloor WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Floor resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Floor, Interfaces.IFloor>();
		}

		/// <summary>
		/// Wrapper interface for ChartFont which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IChartFont WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ChartFont resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ChartFont, Interfaces.IChartFont>();
		}

		/// <summary>
		/// Wrapper interface for Gridlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IGridlines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Gridlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Gridlines, Interfaces.IGridlines>();
		}

		/// <summary>
		/// Wrapper interface for HiLoLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IHiLoLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.HiLoLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.HiLoLines, Interfaces.IHiLoLines>();
		}

		/// <summary>
		/// Wrapper interface for Interior which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IInterior WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Interior resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Interior, Interfaces.IInterior>();
		}

		/// <summary>
		/// Wrapper interface for LeaderLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILeaderLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LeaderLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LeaderLines, Interfaces.ILeaderLines>();
		}

		/// <summary>
		/// Wrapper interface for Legend which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegend WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Legend resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Legend, Interfaces.ILegend>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntries which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntries WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LegendEntries resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LegendEntries, Interfaces.ILegendEntries>();
		}

		/// <summary>
		/// Wrapper interface for LegendEntry which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendEntry WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LegendEntry resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LegendEntry, Interfaces.ILegendEntry>();
		}

		/// <summary>
		/// Wrapper interface for LegendKey which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ILegendKey WithComCleanup(this Microsoft.Office.Interop.PowerPoint.LegendKey resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.LegendKey, Interfaces.ILegendKey>();
		}

		/// <summary>
		/// Wrapper interface for PlotArea which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlotArea WithComCleanup(this Microsoft.Office.Interop.PowerPoint.PlotArea resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.PlotArea, Interfaces.IPlotArea>();
		}

		/// <summary>
		/// Wrapper interface for Point which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoint WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Point resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Point, Interfaces.IPoint>();
		}

		/// <summary>
		/// Wrapper interface for Points which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPoints WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Points resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Points, Interfaces.IPoints>();
		}

		/// <summary>
		/// Wrapper interface for Series which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeries WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Series resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Series, Interfaces.ISeries>();
		}

		/// <summary>
		/// Wrapper interface for SeriesCollection which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesCollection WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SeriesCollection resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SeriesCollection, Interfaces.ISeriesCollection>();
		}

		/// <summary>
		/// Wrapper interface for SeriesLines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISeriesLines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SeriesLines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SeriesLines, Interfaces.ISeriesLines>();
		}

		/// <summary>
		/// Wrapper interface for TickLabels which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITickLabels WithComCleanup(this Microsoft.Office.Interop.PowerPoint.TickLabels resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.TickLabels, Interfaces.ITickLabels>();
		}

		/// <summary>
		/// Wrapper interface for Trendline which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendline WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Trendline resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Trendline, Interfaces.ITrendline>();
		}

		/// <summary>
		/// Wrapper interface for Trendlines which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ITrendlines WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Trendlines resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Trendlines, Interfaces.ITrendlines>();
		}

		/// <summary>
		/// Wrapper interface for UpBars which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IUpBars WithComCleanup(this Microsoft.Office.Interop.PowerPoint.UpBars resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.UpBars, Interfaces.IUpBars>();
		}

		/// <summary>
		/// Wrapper interface for Walls which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IWalls WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Walls resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Walls, Interfaces.IWalls>();
		}

		/// <summary>
		/// Wrapper interface for MediaFormat which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMediaFormat WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MediaFormat resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MediaFormat, Interfaces.IMediaFormat>();
		}

		/// <summary>
		/// Wrapper interface for SectionProperties which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ISectionProperties WithComCleanup(this Microsoft.Office.Interop.PowerPoint.SectionProperties resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.SectionProperties, Interfaces.ISectionProperties>();
		}

		/// <summary>
		/// Wrapper interface for Player which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IPlayer WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Player resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Player, Interfaces.IPlayer>();
		}

		/// <summary>
		/// Wrapper interface for ResampleMediaTask which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResampleMediaTask WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ResampleMediaTask resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ResampleMediaTask, Interfaces.IResampleMediaTask>();
		}

		/// <summary>
		/// Wrapper interface for ResampleMediaTasks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IResampleMediaTasks WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks, Interfaces.IResampleMediaTasks>();
		}

		/// <summary>
		/// Wrapper interface for MediaBookmark which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMediaBookmark WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MediaBookmark resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MediaBookmark, Interfaces.IMediaBookmark>();
		}

		/// <summary>
		/// Wrapper interface for MediaBookmarks which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IMediaBookmarks WithComCleanup(this Microsoft.Office.Interop.PowerPoint.MediaBookmarks resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.MediaBookmarks, Interfaces.IMediaBookmarks>();
		}

		/// <summary>
		/// Wrapper interface for Coauthoring which adds IDispose to the interface
		/// </summary>
		public static Interfaces.ICoauthoring WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Coauthoring resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Coauthoring, Interfaces.ICoauthoring>();
		}

		/// <summary>
		/// Wrapper interface for Broadcast which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IBroadcast WithComCleanup(this Microsoft.Office.Interop.PowerPoint.Broadcast resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.Broadcast, Interfaces.IBroadcast>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindows WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows, Interfaces.IProtectedViewWindows>();
		}

		/// <summary>
		/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
		/// </summary>
		public static Interfaces.IProtectedViewWindow WithComCleanup(this Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow resource)
		{
			return resource.WithComCleanup<Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow, Interfaces.IProtectedViewWindow>();
		}

	}
}