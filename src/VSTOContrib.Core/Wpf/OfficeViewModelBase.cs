using System;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;
using VSTOContrib.Core.Helpers;
using stdole;

namespace VSTOContrib.Core.Wpf
{
    /// <summary>
    /// View Model base for office ribbon view models
    /// </summary>
    public class OfficeViewModelBase : NotifyPropertyChanged
    {
        /// <summary>
        /// OOTB support for /Resources/Image.png (as embedded resource),
        ///  or storing the image in the Resources and use the Image overload
        /// 
        /// pack://application:,,,/MyAddin.Logic;component/Resources/someImage.jpg
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public virtual Bitmap GetPicture(string image)
        {
            var memoryStream = new MemoryStream();
            var bitmap = new Bitmap(memoryStream);
            if (!image.StartsWith("/"))
                image = string.Concat("/", image);

            var encoder = new BmpBitmapEncoder();
            var packApplicationComponent = string.Format(
                "pack://application:,,,/{0};component{1}",
                GetType().Assembly.GetName().Name,
                image);
            encoder.Frames.Add(BitmapFrame.Create(new Uri(packApplicationComponent)));
            encoder.Save(memoryStream);
            return bitmap;
        }

        /// <summary>
        /// Converts a Image into a IPictureDisp image
        /// </summary>
        /// <param name="fromImage"></param>
        /// <returns></returns>
        protected virtual IPictureDisp GetPicture(Image fromImage)
        {
            return PictureConverter.ImageToPictureDisp(fromImage);
        }

        /// <summary>
        /// Converts a Icon into a IPictureDisp image
        /// </summary>
        /// <param name="fromIcon"></param>
        /// <returns></returns>
        protected virtual IPictureDisp GetPicture(Icon fromIcon)
        {
            return PictureConverter.IconToPictureDisp(fromIcon);
        }
    }
}
