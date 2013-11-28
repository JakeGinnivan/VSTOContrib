using System.Drawing;
using System.Windows.Forms;
using stdole;

namespace VSTOContrib.Core.Helpers
{
    /// <summary>
    /// Converts Images to stdole pictures
    /// </summary>
    public class PictureConverter : AxHost
    {
        private PictureConverter() : base(string.Empty) { }

        /// <summary>
        /// Converts image to IPictureDisp for use in the ribbon
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public static IPictureDisp ImageToPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        /// <summary>
        /// Converts icon to IPictureDisp for use in the ribbon
        /// </summary>
        /// <param name="icon"></param>
        /// <returns></returns>
        public static IPictureDisp IconToPictureDisp(Icon icon)
        {
            return ImageToPictureDisp(icon.ToBitmap());
        }

        /// <summary>
        /// Reverse conversion
        /// </summary>
        /// <param name="picture"></param>
        /// <returns></returns>
        public static Image PictureDispToImage(IPictureDisp picture)
        {
            return GetPictureFromIPicture(picture);
        }
    }
}