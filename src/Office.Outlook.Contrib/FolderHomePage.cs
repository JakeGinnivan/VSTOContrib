using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace Office.Outlook.Contrib
{
    /// <summary>
    /// Helper class to register a control to display in a folder.
    /// You MUST mark the assembly for COM Registration and 
    /// register the control with RegisterSafeForScripting. 
    /// This operation requires elevation, recommend using MSI installer and call registration
    /// during installation. 
    /// </summary>
    public sealed class FolderHomePage
	{
		/// <summary>
		/// List of web view files that have been written out during this Outlook instance
		/// </summary>
		private static readonly List<string> ListWebViewFiles = new List<string>();

        /// <summary>
        /// Registers a specific managed type as a folder home page. Returns a file path for the folder home page
        /// </summary>
        /// <returns>file path to the folder home page</returns>
        public static string RegisterType<T>() where T : Control
		{
			//Create the Local App Data directory for the Web view files to reside in
			var webViewDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Properties.Resources.WebViewsDirectory);
			if (Directory.Exists(webViewDirectory) == false)
				Directory.CreateDirectory(webViewDirectory);

			//Create the web view file name based on the viewType guid in the web view directory
		    var viewType = typeof(T);
		    var webViewFile = Path.Combine(webViewDirectory, viewType.GUID.ToString("N") + ".htm");

			//if the file has been written out already in this session, return
			if (ListWebViewFiles.Contains(webViewFile))
				return webViewFile;

			//If the file exists, delete it (for versioning reasons)
			if (File.Exists(webViewFile))
				File.Delete(webViewFile);

			//Open a file stream and text writer for the Web view stream
			using (var stm = new FileStream(webViewFile, FileMode.Create, FileAccess.Write))
			using (var writer = new StreamWriter(stm, System.Text.Encoding.ASCII))
			{
    			//Look to see if the viewType has an init method that takes a single Outlook App parameter
    			var initInfo = viewType.GetMethod("Initialize", new[] { typeof(Application) });

    			//If the viewType doesn't have an Init method, just write out the html page header
    			//TODO move HTML code to resource strings
    			if (initInfo == null)
    			{
    				writer.WriteLine("<html><body rightmargin = '0' leftmargin ='0' topmargin ='0' bottommargin = '0'>");
    			}
    			//If the viewType does have an Init method, write script to trap the Body.OnLoad event and call the Init method
    			//passing in the window.external.OutlookApplication object as the parameter
    			else
    			{
    				writer.WriteLine("<html><body rightmargin = '0' leftmargin ='0' topmargin ='0' bottommargin = '0' onload='OnBodyLoad()'>");
    				writer.WriteLine("<script>\n\tfunction OnBodyLoad()\n\t{\n\t\tvar oApp = window.external.OutlookApplication;");
    				writer.WriteLine("\t\t{0}.Initialize(oApp);", viewType.Name);
    				writer.WriteLine("\t}\n</script>");
    			}

    			//Write out an object tag that loads up the viewType as a com object via its class id
    			writer.WriteLine("<object classid='clsid:{0}' ID='{1}' VIEWASTEXT width='100%' height='100%'/>", viewType.GUID, viewType.Name);
    			writer.WriteLine("</body></html>");

    			//Close the file
    			writer.Close();
    			stm.Close();
            }
			//save this file name so we don't write it out multiple times per outlook session
			ListWebViewFiles.Add(webViewFile);

			return webViewFile;
		}

		private const string CatidSafeForScripting = "7DD95801-9882-11CF-9FA9-00AA006C42C4";
		private const string CatidSafeForInitializing = "7DD95802-9882-11CF-9FA9-00AA006C42C4";

		/// <summary>
		/// Registers a managed type that's exposed for COM interop as safe for initializing and scripting
		/// </summary>
		/// <param name="comType"></param>
		public static void RegisterSafeForScripting(Type comType)
		{
			var clsid = comType.GUID;
			var interfaceSafeScripting = new Guid(CatidSafeForScripting);
			var interfaceSafeForInitializing = new Guid(CatidSafeForInitializing);

			var reg = (ICatRegister)new ComComponentCategoriesManager();
			reg.RegisterClassImplCategories(ref clsid, 1, new[] { interfaceSafeScripting });
			reg.RegisterClassImplCategories(ref clsid, 1, new[] { interfaceSafeForInitializing });
		}

		/// <summary>
		/// Unregisters a managed type that's exposed for COM interop as safe for initializing and scripting
		/// </summary>
		/// <param name="comType"></param>
		public static void UnregisterSafeForScripting(Type comType)
		{
			var clsid = comType.GUID;
			var interfaceSafeScripting = new Guid(CatidSafeForScripting);
			var interfaceSafeForInitializing = new Guid(CatidSafeForInitializing);

			var reg = (ICatRegister)new ComComponentCategoriesManager();
			reg.UnRegisterClassImplCategories(ref clsid, 1, new[] { interfaceSafeScripting });
			reg.UnRegisterClassImplCategories(ref clsid, 1, new[] { interfaceSafeForInitializing });
		}

	}
}
