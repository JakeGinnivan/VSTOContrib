//Microsoft.Office.Interop.PowerPoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c
namespace VSTOContrib.PowerPoint.Extensions.Proxies.Interfaces
{
	/// <summary>
	/// Wrapper interface for Collection which adds IDispose to the interface
	/// </summary>
	public interface ICollection : Microsoft.Office.Interop.PowerPoint.Collection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Collection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Application which adds IDispose to the interface
	/// </summary>
	public interface I_Application : Microsoft.Office.Interop.PowerPoint._Application, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint._Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Global which adds IDispose to the interface
	/// </summary>
	public interface I_Global : Microsoft.Office.Interop.PowerPoint._Global, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint._Global Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EApplication_Event which adds IDispose to the interface
	/// </summary>
	public interface IEApplication_Event : Microsoft.Office.Interop.PowerPoint.EApplication_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.EApplication_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Application which adds IDispose to the interface
	/// </summary>
	public interface IApplication : Microsoft.Office.Interop.PowerPoint.Application, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Application Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Global which adds IDispose to the interface
	/// </summary>
	public interface IGlobal : Microsoft.Office.Interop.PowerPoint.Global, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Global Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IColorFormat : Microsoft.Office.Interop.PowerPoint.ColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlideShowWindow which adds IDispose to the interface
	/// </summary>
	public interface ISlideShowWindow : Microsoft.Office.Interop.PowerPoint.SlideShowWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SlideShowWindow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Selection which adds IDispose to the interface
	/// </summary>
	public interface ISelection : Microsoft.Office.Interop.PowerPoint.Selection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Selection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentWindows which adds IDispose to the interface
	/// </summary>
	public interface IDocumentWindows : Microsoft.Office.Interop.PowerPoint.DocumentWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DocumentWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlideShowWindows which adds IDispose to the interface
	/// </summary>
	public interface ISlideShowWindows : Microsoft.Office.Interop.PowerPoint.SlideShowWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SlideShowWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DocumentWindow which adds IDispose to the interface
	/// </summary>
	public interface IDocumentWindow : Microsoft.Office.Interop.PowerPoint.DocumentWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DocumentWindow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for View which adds IDispose to the interface
	/// </summary>
	public interface IView : Microsoft.Office.Interop.PowerPoint.View, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.View Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlideShowView which adds IDispose to the interface
	/// </summary>
	public interface ISlideShowView : Microsoft.Office.Interop.PowerPoint.SlideShowView, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SlideShowView Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlideShowSettings which adds IDispose to the interface
	/// </summary>
	public interface ISlideShowSettings : Microsoft.Office.Interop.PowerPoint.SlideShowSettings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SlideShowSettings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NamedSlideShows which adds IDispose to the interface
	/// </summary>
	public interface INamedSlideShows : Microsoft.Office.Interop.PowerPoint.NamedSlideShows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.NamedSlideShows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for NamedSlideShow which adds IDispose to the interface
	/// </summary>
	public interface INamedSlideShow : Microsoft.Office.Interop.PowerPoint.NamedSlideShow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.NamedSlideShow Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PrintOptions which adds IDispose to the interface
	/// </summary>
	public interface IPrintOptions : Microsoft.Office.Interop.PowerPoint.PrintOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PrintOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PrintRanges which adds IDispose to the interface
	/// </summary>
	public interface IPrintRanges : Microsoft.Office.Interop.PowerPoint.PrintRanges, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PrintRanges Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PrintRange which adds IDispose to the interface
	/// </summary>
	public interface IPrintRange : Microsoft.Office.Interop.PowerPoint.PrintRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PrintRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIns which adds IDispose to the interface
	/// </summary>
	public interface IAddIns : Microsoft.Office.Interop.PowerPoint.AddIns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AddIns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AddIn which adds IDispose to the interface
	/// </summary>
	public interface IAddIn : Microsoft.Office.Interop.PowerPoint.AddIn, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AddIn Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Presentations which adds IDispose to the interface
	/// </summary>
	public interface IPresentations : Microsoft.Office.Interop.PowerPoint.Presentations, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Presentations Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PresEvents which adds IDispose to the interface
	/// </summary>
	public interface IPresEvents : Microsoft.Office.Interop.PowerPoint.PresEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PresEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PresEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IPresEvents_Event : Microsoft.Office.Interop.PowerPoint.PresEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PresEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Presentation which adds IDispose to the interface
	/// </summary>
	public interface IPresentation : Microsoft.Office.Interop.PowerPoint.Presentation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Presentation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Hyperlinks which adds IDispose to the interface
	/// </summary>
	public interface IHyperlinks : Microsoft.Office.Interop.PowerPoint.Hyperlinks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Hyperlinks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Hyperlink which adds IDispose to the interface
	/// </summary>
	public interface IHyperlink : Microsoft.Office.Interop.PowerPoint.Hyperlink, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Hyperlink Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PageSetup which adds IDispose to the interface
	/// </summary>
	public interface IPageSetup : Microsoft.Office.Interop.PowerPoint.PageSetup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PageSetup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Fonts which adds IDispose to the interface
	/// </summary>
	public interface IFonts : Microsoft.Office.Interop.PowerPoint.Fonts, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Fonts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ExtraColors which adds IDispose to the interface
	/// </summary>
	public interface IExtraColors : Microsoft.Office.Interop.PowerPoint.ExtraColors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ExtraColors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Slides which adds IDispose to the interface
	/// </summary>
	public interface ISlides : Microsoft.Office.Interop.PowerPoint.Slides, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Slides Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Slide which adds IDispose to the interface
	/// </summary>
	public interface I_Slide : Microsoft.Office.Interop.PowerPoint._Slide, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint._Slide Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlideRange which adds IDispose to the interface
	/// </summary>
	public interface ISlideRange : Microsoft.Office.Interop.PowerPoint.SlideRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SlideRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Master which adds IDispose to the interface
	/// </summary>
	public interface I_Master : Microsoft.Office.Interop.PowerPoint._Master, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint._Master Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SldEvents which adds IDispose to the interface
	/// </summary>
	public interface ISldEvents : Microsoft.Office.Interop.PowerPoint.SldEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SldEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SldEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface ISldEvents_Event : Microsoft.Office.Interop.PowerPoint.SldEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SldEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Slide which adds IDispose to the interface
	/// </summary>
	public interface ISlide : Microsoft.Office.Interop.PowerPoint.Slide, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Slide Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorSchemes which adds IDispose to the interface
	/// </summary>
	public interface IColorSchemes : Microsoft.Office.Interop.PowerPoint.ColorSchemes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ColorSchemes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorScheme which adds IDispose to the interface
	/// </summary>
	public interface IColorScheme : Microsoft.Office.Interop.PowerPoint.ColorScheme, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ColorScheme Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RGBColor which adds IDispose to the interface
	/// </summary>
	public interface IRGBColor : Microsoft.Office.Interop.PowerPoint.RGBColor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.RGBColor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SlideShowTransition which adds IDispose to the interface
	/// </summary>
	public interface ISlideShowTransition : Microsoft.Office.Interop.PowerPoint.SlideShowTransition, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SlideShowTransition Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SoundEffect which adds IDispose to the interface
	/// </summary>
	public interface ISoundEffect : Microsoft.Office.Interop.PowerPoint.SoundEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SoundEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SoundFormat which adds IDispose to the interface
	/// </summary>
	public interface ISoundFormat : Microsoft.Office.Interop.PowerPoint.SoundFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SoundFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeadersFooters which adds IDispose to the interface
	/// </summary>
	public interface IHeadersFooters : Microsoft.Office.Interop.PowerPoint.HeadersFooters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.HeadersFooters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shapes which adds IDispose to the interface
	/// </summary>
	public interface IShapes : Microsoft.Office.Interop.PowerPoint.Shapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Shapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Placeholders which adds IDispose to the interface
	/// </summary>
	public interface IPlaceholders : Microsoft.Office.Interop.PowerPoint.Placeholders, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Placeholders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PlaceholderFormat which adds IDispose to the interface
	/// </summary>
	public interface IPlaceholderFormat : Microsoft.Office.Interop.PowerPoint.PlaceholderFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PlaceholderFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FreeformBuilder which adds IDispose to the interface
	/// </summary>
	public interface IFreeformBuilder : Microsoft.Office.Interop.PowerPoint.FreeformBuilder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.FreeformBuilder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Shape which adds IDispose to the interface
	/// </summary>
	public interface IShape : Microsoft.Office.Interop.PowerPoint.Shape, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Shape Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeRange which adds IDispose to the interface
	/// </summary>
	public interface IShapeRange : Microsoft.Office.Interop.PowerPoint.ShapeRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ShapeRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for GroupShapes which adds IDispose to the interface
	/// </summary>
	public interface IGroupShapes : Microsoft.Office.Interop.PowerPoint.GroupShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.GroupShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Adjustments which adds IDispose to the interface
	/// </summary>
	public interface IAdjustments : Microsoft.Office.Interop.PowerPoint.Adjustments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Adjustments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PictureFormat which adds IDispose to the interface
	/// </summary>
	public interface IPictureFormat : Microsoft.Office.Interop.PowerPoint.PictureFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PictureFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FillFormat which adds IDispose to the interface
	/// </summary>
	public interface IFillFormat : Microsoft.Office.Interop.PowerPoint.FillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.FillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LineFormat which adds IDispose to the interface
	/// </summary>
	public interface ILineFormat : Microsoft.Office.Interop.PowerPoint.LineFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.LineFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShadowFormat which adds IDispose to the interface
	/// </summary>
	public interface IShadowFormat : Microsoft.Office.Interop.PowerPoint.ShadowFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ShadowFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ConnectorFormat which adds IDispose to the interface
	/// </summary>
	public interface IConnectorFormat : Microsoft.Office.Interop.PowerPoint.ConnectorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ConnectorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextEffectFormat which adds IDispose to the interface
	/// </summary>
	public interface ITextEffectFormat : Microsoft.Office.Interop.PowerPoint.TextEffectFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextEffectFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ThreeDFormat which adds IDispose to the interface
	/// </summary>
	public interface IThreeDFormat : Microsoft.Office.Interop.PowerPoint.ThreeDFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ThreeDFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame : Microsoft.Office.Interop.PowerPoint.TextFrame, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextFrame Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CalloutFormat which adds IDispose to the interface
	/// </summary>
	public interface ICalloutFormat : Microsoft.Office.Interop.PowerPoint.CalloutFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CalloutFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNodes which adds IDispose to the interface
	/// </summary>
	public interface IShapeNodes : Microsoft.Office.Interop.PowerPoint.ShapeNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ShapeNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ShapeNode which adds IDispose to the interface
	/// </summary>
	public interface IShapeNode : Microsoft.Office.Interop.PowerPoint.ShapeNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ShapeNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEFormat which adds IDispose to the interface
	/// </summary>
	public interface IOLEFormat : Microsoft.Office.Interop.PowerPoint.OLEFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.OLEFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LinkFormat which adds IDispose to the interface
	/// </summary>
	public interface ILinkFormat : Microsoft.Office.Interop.PowerPoint.LinkFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.LinkFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ObjectVerbs which adds IDispose to the interface
	/// </summary>
	public interface IObjectVerbs : Microsoft.Office.Interop.PowerPoint.ObjectVerbs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ObjectVerbs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnimationSettings which adds IDispose to the interface
	/// </summary>
	public interface IAnimationSettings : Microsoft.Office.Interop.PowerPoint.AnimationSettings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AnimationSettings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ActionSettings which adds IDispose to the interface
	/// </summary>
	public interface IActionSettings : Microsoft.Office.Interop.PowerPoint.ActionSettings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ActionSettings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ActionSetting which adds IDispose to the interface
	/// </summary>
	public interface IActionSetting : Microsoft.Office.Interop.PowerPoint.ActionSetting, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ActionSetting Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PlaySettings which adds IDispose to the interface
	/// </summary>
	public interface IPlaySettings : Microsoft.Office.Interop.PowerPoint.PlaySettings, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PlaySettings Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextRange which adds IDispose to the interface
	/// </summary>
	public interface ITextRange : Microsoft.Office.Interop.PowerPoint.TextRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Ruler which adds IDispose to the interface
	/// </summary>
	public interface IRuler : Microsoft.Office.Interop.PowerPoint.Ruler, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Ruler Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RulerLevels which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevels : Microsoft.Office.Interop.PowerPoint.RulerLevels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.RulerLevels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RulerLevel which adds IDispose to the interface
	/// </summary>
	public interface IRulerLevel : Microsoft.Office.Interop.PowerPoint.RulerLevel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.RulerLevel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TabStops which adds IDispose to the interface
	/// </summary>
	public interface ITabStops : Microsoft.Office.Interop.PowerPoint.TabStops, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TabStops Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TabStop which adds IDispose to the interface
	/// </summary>
	public interface ITabStop : Microsoft.Office.Interop.PowerPoint.TabStop, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TabStop Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Font which adds IDispose to the interface
	/// </summary>
	public interface IFont : Microsoft.Office.Interop.PowerPoint.Font, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Font Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ParagraphFormat which adds IDispose to the interface
	/// </summary>
	public interface IParagraphFormat : Microsoft.Office.Interop.PowerPoint.ParagraphFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ParagraphFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for BulletFormat which adds IDispose to the interface
	/// </summary>
	public interface IBulletFormat : Microsoft.Office.Interop.PowerPoint.BulletFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.BulletFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextStyles which adds IDispose to the interface
	/// </summary>
	public interface ITextStyles : Microsoft.Office.Interop.PowerPoint.TextStyles, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextStyles Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextStyle which adds IDispose to the interface
	/// </summary>
	public interface ITextStyle : Microsoft.Office.Interop.PowerPoint.TextStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextStyleLevels which adds IDispose to the interface
	/// </summary>
	public interface ITextStyleLevels : Microsoft.Office.Interop.PowerPoint.TextStyleLevels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextStyleLevels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextStyleLevel which adds IDispose to the interface
	/// </summary>
	public interface ITextStyleLevel : Microsoft.Office.Interop.PowerPoint.TextStyleLevel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextStyleLevel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HeaderFooter which adds IDispose to the interface
	/// </summary>
	public interface IHeaderFooter : Microsoft.Office.Interop.PowerPoint.HeaderFooter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.HeaderFooter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _Presentation which adds IDispose to the interface
	/// </summary>
	public interface I_Presentation : Microsoft.Office.Interop.PowerPoint._Presentation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint._Presentation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Tags which adds IDispose to the interface
	/// </summary>
	public interface ITags : Microsoft.Office.Interop.PowerPoint.Tags, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Tags Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MouseTracker which adds IDispose to the interface
	/// </summary>
	public interface IMouseTracker : Microsoft.Office.Interop.PowerPoint.MouseTracker, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MouseTracker Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MouseDownHandler which adds IDispose to the interface
	/// </summary>
	public interface IMouseDownHandler : Microsoft.Office.Interop.PowerPoint.MouseDownHandler, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MouseDownHandler Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OCXExtender which adds IDispose to the interface
	/// </summary>
	public interface IOCXExtender : Microsoft.Office.Interop.PowerPoint.OCXExtender, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.OCXExtender Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OCXExtenderEvents which adds IDispose to the interface
	/// </summary>
	public interface IOCXExtenderEvents : Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OCXExtenderEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IOCXExtenderEvents_Event : Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.OCXExtenderEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for OLEControl which adds IDispose to the interface
	/// </summary>
	public interface IOLEControl : Microsoft.Office.Interop.PowerPoint.OLEControl, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.OLEControl Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EApplication which adds IDispose to the interface
	/// </summary>
	public interface IEApplication : Microsoft.Office.Interop.PowerPoint.EApplication, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.EApplication Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Table which adds IDispose to the interface
	/// </summary>
	public interface ITable : Microsoft.Office.Interop.PowerPoint.Table, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Table Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Columns which adds IDispose to the interface
	/// </summary>
	public interface IColumns : Microsoft.Office.Interop.PowerPoint.Columns, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Columns Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Column which adds IDispose to the interface
	/// </summary>
	public interface IColumn : Microsoft.Office.Interop.PowerPoint.Column, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Column Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Rows which adds IDispose to the interface
	/// </summary>
	public interface IRows : Microsoft.Office.Interop.PowerPoint.Rows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Rows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Row which adds IDispose to the interface
	/// </summary>
	public interface IRow : Microsoft.Office.Interop.PowerPoint.Row, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Row Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CellRange which adds IDispose to the interface
	/// </summary>
	public interface ICellRange : Microsoft.Office.Interop.PowerPoint.CellRange, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CellRange Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Cell which adds IDispose to the interface
	/// </summary>
	public interface ICell : Microsoft.Office.Interop.PowerPoint.Cell, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Cell Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Borders which adds IDispose to the interface
	/// </summary>
	public interface IBorders : Microsoft.Office.Interop.PowerPoint.Borders, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Borders Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Panes which adds IDispose to the interface
	/// </summary>
	public interface IPanes : Microsoft.Office.Interop.PowerPoint.Panes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Panes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Pane which adds IDispose to the interface
	/// </summary>
	public interface IPane : Microsoft.Office.Interop.PowerPoint.Pane, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Pane Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DefaultWebOptions which adds IDispose to the interface
	/// </summary>
	public interface IDefaultWebOptions : Microsoft.Office.Interop.PowerPoint.DefaultWebOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DefaultWebOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for WebOptions which adds IDispose to the interface
	/// </summary>
	public interface IWebOptions : Microsoft.Office.Interop.PowerPoint.WebOptions, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.WebOptions Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PublishObjects which adds IDispose to the interface
	/// </summary>
	public interface IPublishObjects : Microsoft.Office.Interop.PowerPoint.PublishObjects, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PublishObjects Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PublishObject which adds IDispose to the interface
	/// </summary>
	public interface IPublishObject : Microsoft.Office.Interop.PowerPoint.PublishObject, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PublishObject Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MasterEvents which adds IDispose to the interface
	/// </summary>
	public interface IMasterEvents : Microsoft.Office.Interop.PowerPoint.MasterEvents, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MasterEvents Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MasterEvents_Event which adds IDispose to the interface
	/// </summary>
	public interface IMasterEvents_Event : Microsoft.Office.Interop.PowerPoint.MasterEvents_Event, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MasterEvents_Event Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Master which adds IDispose to the interface
	/// </summary>
	public interface IMaster : Microsoft.Office.Interop.PowerPoint.Master, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Master Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for _PowerRex which adds IDispose to the interface
	/// </summary>
	public interface I_PowerRex : Microsoft.Office.Interop.PowerPoint._PowerRex, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint._PowerRex Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PowerRex which adds IDispose to the interface
	/// </summary>
	public interface IPowerRex : Microsoft.Office.Interop.PowerPoint.PowerRex, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PowerRex Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Comments which adds IDispose to the interface
	/// </summary>
	public interface IComments : Microsoft.Office.Interop.PowerPoint.Comments, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Comments Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Comment which adds IDispose to the interface
	/// </summary>
	public interface IComment : Microsoft.Office.Interop.PowerPoint.Comment, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Comment Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Designs which adds IDispose to the interface
	/// </summary>
	public interface IDesigns : Microsoft.Office.Interop.PowerPoint.Designs, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Designs Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Design which adds IDispose to the interface
	/// </summary>
	public interface IDesign : Microsoft.Office.Interop.PowerPoint.Design, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Design Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNode which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNode : Microsoft.Office.Interop.PowerPoint.DiagramNode, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DiagramNode Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodeChildren which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodeChildren : Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DiagramNodeChildren Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DiagramNodes which adds IDispose to the interface
	/// </summary>
	public interface IDiagramNodes : Microsoft.Office.Interop.PowerPoint.DiagramNodes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DiagramNodes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Diagram which adds IDispose to the interface
	/// </summary>
	public interface IDiagram : Microsoft.Office.Interop.PowerPoint.Diagram, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Diagram Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TimeLine which adds IDispose to the interface
	/// </summary>
	public interface ITimeLine : Microsoft.Office.Interop.PowerPoint.TimeLine, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TimeLine Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sequences which adds IDispose to the interface
	/// </summary>
	public interface ISequences : Microsoft.Office.Interop.PowerPoint.Sequences, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Sequences Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Sequence which adds IDispose to the interface
	/// </summary>
	public interface ISequence : Microsoft.Office.Interop.PowerPoint.Sequence, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Sequence Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Effect which adds IDispose to the interface
	/// </summary>
	public interface IEffect : Microsoft.Office.Interop.PowerPoint.Effect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Effect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Timing which adds IDispose to the interface
	/// </summary>
	public interface ITiming : Microsoft.Office.Interop.PowerPoint.Timing, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Timing Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EffectParameters which adds IDispose to the interface
	/// </summary>
	public interface IEffectParameters : Microsoft.Office.Interop.PowerPoint.EffectParameters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.EffectParameters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for EffectInformation which adds IDispose to the interface
	/// </summary>
	public interface IEffectInformation : Microsoft.Office.Interop.PowerPoint.EffectInformation, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.EffectInformation Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnimationBehaviors which adds IDispose to the interface
	/// </summary>
	public interface IAnimationBehaviors : Microsoft.Office.Interop.PowerPoint.AnimationBehaviors, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AnimationBehaviors Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnimationBehavior which adds IDispose to the interface
	/// </summary>
	public interface IAnimationBehavior : Microsoft.Office.Interop.PowerPoint.AnimationBehavior, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AnimationBehavior Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MotionEffect which adds IDispose to the interface
	/// </summary>
	public interface IMotionEffect : Microsoft.Office.Interop.PowerPoint.MotionEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MotionEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ColorEffect which adds IDispose to the interface
	/// </summary>
	public interface IColorEffect : Microsoft.Office.Interop.PowerPoint.ColorEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ColorEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ScaleEffect which adds IDispose to the interface
	/// </summary>
	public interface IScaleEffect : Microsoft.Office.Interop.PowerPoint.ScaleEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ScaleEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for RotationEffect which adds IDispose to the interface
	/// </summary>
	public interface IRotationEffect : Microsoft.Office.Interop.PowerPoint.RotationEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.RotationEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PropertyEffect which adds IDispose to the interface
	/// </summary>
	public interface IPropertyEffect : Microsoft.Office.Interop.PowerPoint.PropertyEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PropertyEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnimationPoints which adds IDispose to the interface
	/// </summary>
	public interface IAnimationPoints : Microsoft.Office.Interop.PowerPoint.AnimationPoints, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AnimationPoints Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AnimationPoint which adds IDispose to the interface
	/// </summary>
	public interface IAnimationPoint : Microsoft.Office.Interop.PowerPoint.AnimationPoint, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AnimationPoint Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CanvasShapes which adds IDispose to the interface
	/// </summary>
	public interface ICanvasShapes : Microsoft.Office.Interop.PowerPoint.CanvasShapes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CanvasShapes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AutoCorrect which adds IDispose to the interface
	/// </summary>
	public interface IAutoCorrect : Microsoft.Office.Interop.PowerPoint.AutoCorrect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AutoCorrect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Options which adds IDispose to the interface
	/// </summary>
	public interface IOptions : Microsoft.Office.Interop.PowerPoint.Options, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Options Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CommandEffect which adds IDispose to the interface
	/// </summary>
	public interface ICommandEffect : Microsoft.Office.Interop.PowerPoint.CommandEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CommandEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FilterEffect which adds IDispose to the interface
	/// </summary>
	public interface IFilterEffect : Microsoft.Office.Interop.PowerPoint.FilterEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.FilterEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SetEffect which adds IDispose to the interface
	/// </summary>
	public interface ISetEffect : Microsoft.Office.Interop.PowerPoint.SetEffect, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SetEffect Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomLayouts which adds IDispose to the interface
	/// </summary>
	public interface ICustomLayouts : Microsoft.Office.Interop.PowerPoint.CustomLayouts, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CustomLayouts Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomLayout which adds IDispose to the interface
	/// </summary>
	public interface ICustomLayout : Microsoft.Office.Interop.PowerPoint.CustomLayout, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CustomLayout Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableStyle which adds IDispose to the interface
	/// </summary>
	public interface ITableStyle : Microsoft.Office.Interop.PowerPoint.TableStyle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TableStyle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for CustomerData which adds IDispose to the interface
	/// </summary>
	public interface ICustomerData : Microsoft.Office.Interop.PowerPoint.CustomerData, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.CustomerData Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Research which adds IDispose to the interface
	/// </summary>
	public interface IResearch : Microsoft.Office.Interop.PowerPoint.Research, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Research Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TableBackground which adds IDispose to the interface
	/// </summary>
	public interface ITableBackground : Microsoft.Office.Interop.PowerPoint.TableBackground, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TableBackground Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TextFrame2 which adds IDispose to the interface
	/// </summary>
	public interface ITextFrame2 : Microsoft.Office.Interop.PowerPoint.TextFrame2, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TextFrame2 Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileConverters which adds IDispose to the interface
	/// </summary>
	public interface IFileConverters : Microsoft.Office.Interop.PowerPoint.FileConverters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.FileConverters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for FileConverter which adds IDispose to the interface
	/// </summary>
	public interface IFileConverter : Microsoft.Office.Interop.PowerPoint.FileConverter, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.FileConverter Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axes which adds IDispose to the interface
	/// </summary>
	public interface IAxes : Microsoft.Office.Interop.PowerPoint.Axes, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Axes Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Axis which adds IDispose to the interface
	/// </summary>
	public interface IAxis : Microsoft.Office.Interop.PowerPoint.Axis, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Axis Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for AxisTitle which adds IDispose to the interface
	/// </summary>
	public interface IAxisTitle : Microsoft.Office.Interop.PowerPoint.AxisTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.AxisTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Chart which adds IDispose to the interface
	/// </summary>
	public interface IChart : Microsoft.Office.Interop.PowerPoint.Chart, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Chart Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartBorder which adds IDispose to the interface
	/// </summary>
	public interface IChartBorder : Microsoft.Office.Interop.PowerPoint.ChartBorder, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartBorder Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartCharacters which adds IDispose to the interface
	/// </summary>
	public interface IChartCharacters : Microsoft.Office.Interop.PowerPoint.ChartCharacters, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartCharacters Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartArea which adds IDispose to the interface
	/// </summary>
	public interface IChartArea : Microsoft.Office.Interop.PowerPoint.ChartArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartColorFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartColorFormat : Microsoft.Office.Interop.PowerPoint.ChartColorFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartColorFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartData which adds IDispose to the interface
	/// </summary>
	public interface IChartData : Microsoft.Office.Interop.PowerPoint.ChartData, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartData Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFillFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFillFormat : Microsoft.Office.Interop.PowerPoint.ChartFillFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartFillFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFormat which adds IDispose to the interface
	/// </summary>
	public interface IChartFormat : Microsoft.Office.Interop.PowerPoint.ChartFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroup which adds IDispose to the interface
	/// </summary>
	public interface IChartGroup : Microsoft.Office.Interop.PowerPoint.ChartGroup, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartGroup Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartGroups which adds IDispose to the interface
	/// </summary>
	public interface IChartGroups : Microsoft.Office.Interop.PowerPoint.ChartGroups, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartGroups Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartTitle which adds IDispose to the interface
	/// </summary>
	public interface IChartTitle : Microsoft.Office.Interop.PowerPoint.ChartTitle, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartTitle Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Corners which adds IDispose to the interface
	/// </summary>
	public interface ICorners : Microsoft.Office.Interop.PowerPoint.Corners, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Corners Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataLabel which adds IDispose to the interface
	/// </summary>
	public interface IDataLabel : Microsoft.Office.Interop.PowerPoint.DataLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DataLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataLabels which adds IDispose to the interface
	/// </summary>
	public interface IDataLabels : Microsoft.Office.Interop.PowerPoint.DataLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DataLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DataTable which adds IDispose to the interface
	/// </summary>
	public interface IDataTable : Microsoft.Office.Interop.PowerPoint.DataTable, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DataTable Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DisplayUnitLabel which adds IDispose to the interface
	/// </summary>
	public interface IDisplayUnitLabel : Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DisplayUnitLabel Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DownBars which adds IDispose to the interface
	/// </summary>
	public interface IDownBars : Microsoft.Office.Interop.PowerPoint.DownBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DownBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for DropLines which adds IDispose to the interface
	/// </summary>
	public interface IDropLines : Microsoft.Office.Interop.PowerPoint.DropLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.DropLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ErrorBars which adds IDispose to the interface
	/// </summary>
	public interface IErrorBars : Microsoft.Office.Interop.PowerPoint.ErrorBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ErrorBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Floor which adds IDispose to the interface
	/// </summary>
	public interface IFloor : Microsoft.Office.Interop.PowerPoint.Floor, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Floor Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ChartFont which adds IDispose to the interface
	/// </summary>
	public interface IChartFont : Microsoft.Office.Interop.PowerPoint.ChartFont, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ChartFont Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Gridlines which adds IDispose to the interface
	/// </summary>
	public interface IGridlines : Microsoft.Office.Interop.PowerPoint.Gridlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Gridlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for HiLoLines which adds IDispose to the interface
	/// </summary>
	public interface IHiLoLines : Microsoft.Office.Interop.PowerPoint.HiLoLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.HiLoLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Interior which adds IDispose to the interface
	/// </summary>
	public interface IInterior : Microsoft.Office.Interop.PowerPoint.Interior, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Interior Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LeaderLines which adds IDispose to the interface
	/// </summary>
	public interface ILeaderLines : Microsoft.Office.Interop.PowerPoint.LeaderLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.LeaderLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Legend which adds IDispose to the interface
	/// </summary>
	public interface ILegend : Microsoft.Office.Interop.PowerPoint.Legend, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Legend Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntries which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntries : Microsoft.Office.Interop.PowerPoint.LegendEntries, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.LegendEntries Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendEntry which adds IDispose to the interface
	/// </summary>
	public interface ILegendEntry : Microsoft.Office.Interop.PowerPoint.LegendEntry, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.LegendEntry Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for LegendKey which adds IDispose to the interface
	/// </summary>
	public interface ILegendKey : Microsoft.Office.Interop.PowerPoint.LegendKey, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.LegendKey Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for PlotArea which adds IDispose to the interface
	/// </summary>
	public interface IPlotArea : Microsoft.Office.Interop.PowerPoint.PlotArea, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.PlotArea Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Point which adds IDispose to the interface
	/// </summary>
	public interface IPoint : Microsoft.Office.Interop.PowerPoint.Point, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Point Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Points which adds IDispose to the interface
	/// </summary>
	public interface IPoints : Microsoft.Office.Interop.PowerPoint.Points, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Points Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Series which adds IDispose to the interface
	/// </summary>
	public interface ISeries : Microsoft.Office.Interop.PowerPoint.Series, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Series Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesCollection which adds IDispose to the interface
	/// </summary>
	public interface ISeriesCollection : Microsoft.Office.Interop.PowerPoint.SeriesCollection, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SeriesCollection Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SeriesLines which adds IDispose to the interface
	/// </summary>
	public interface ISeriesLines : Microsoft.Office.Interop.PowerPoint.SeriesLines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SeriesLines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for TickLabels which adds IDispose to the interface
	/// </summary>
	public interface ITickLabels : Microsoft.Office.Interop.PowerPoint.TickLabels, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.TickLabels Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendline which adds IDispose to the interface
	/// </summary>
	public interface ITrendline : Microsoft.Office.Interop.PowerPoint.Trendline, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Trendline Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Trendlines which adds IDispose to the interface
	/// </summary>
	public interface ITrendlines : Microsoft.Office.Interop.PowerPoint.Trendlines, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Trendlines Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for UpBars which adds IDispose to the interface
	/// </summary>
	public interface IUpBars : Microsoft.Office.Interop.PowerPoint.UpBars, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.UpBars Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Walls which adds IDispose to the interface
	/// </summary>
	public interface IWalls : Microsoft.Office.Interop.PowerPoint.Walls, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Walls Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MediaFormat which adds IDispose to the interface
	/// </summary>
	public interface IMediaFormat : Microsoft.Office.Interop.PowerPoint.MediaFormat, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MediaFormat Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for SectionProperties which adds IDispose to the interface
	/// </summary>
	public interface ISectionProperties : Microsoft.Office.Interop.PowerPoint.SectionProperties, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.SectionProperties Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Player which adds IDispose to the interface
	/// </summary>
	public interface IPlayer : Microsoft.Office.Interop.PowerPoint.Player, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Player Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ResampleMediaTask which adds IDispose to the interface
	/// </summary>
	public interface IResampleMediaTask : Microsoft.Office.Interop.PowerPoint.ResampleMediaTask, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ResampleMediaTask Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ResampleMediaTasks which adds IDispose to the interface
	/// </summary>
	public interface IResampleMediaTasks : Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ResampleMediaTasks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MediaBookmark which adds IDispose to the interface
	/// </summary>
	public interface IMediaBookmark : Microsoft.Office.Interop.PowerPoint.MediaBookmark, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MediaBookmark Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for MediaBookmarks which adds IDispose to the interface
	/// </summary>
	public interface IMediaBookmarks : Microsoft.Office.Interop.PowerPoint.MediaBookmarks, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.MediaBookmarks Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Coauthoring which adds IDispose to the interface
	/// </summary>
	public interface ICoauthoring : Microsoft.Office.Interop.PowerPoint.Coauthoring, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Coauthoring Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for Broadcast which adds IDispose to the interface
	/// </summary>
	public interface IBroadcast : Microsoft.Office.Interop.PowerPoint.Broadcast, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.Broadcast Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProtectedViewWindows which adds IDispose to the interface
	/// </summary>
	public interface IProtectedViewWindows : Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ProtectedViewWindows Resource { get; }
	}

	/// <summary>
	/// Wrapper interface for ProtectedViewWindow which adds IDispose to the interface
	/// </summary>
	public interface IProtectedViewWindow : Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow, System.IDisposable  
	{ 
		/// <summary>
        /// Gets the proxied resource.
        /// </summary>
        /// <value>The resource.</value>
		Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow Resource { get; }
	}

	}