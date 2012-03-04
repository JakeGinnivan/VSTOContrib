using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq.Expressions;
using System.Windows.Media.Imaging;
using VSTOContrib.Core.Helpers;
using stdole;

namespace VSTOContrib.Core.Wpf
{
    /// <summary>
    /// View Model base for office ribbon view models
    /// </summary>
    public class OfficeViewModelBase : INotifyPropertyChanged
    {
        /// <summary>
        /// OOTB support for /Resources/Image.png (as embedded resource),
        ///  or storing the image in the Resources and use the Image overload
        /// 
        /// pack://application:,,,/MyAddin.Logic;component/Resources/someImage.jpg
        /// </summary>
        /// <param name="image"></param>
        /// <returns></returns>
        public virtual IPictureDisp GetPicture(string image)
        {
            using (var memoryStream = new MemoryStream())
            using (var bitmap = new Bitmap(memoryStream))
            {
                if (!image.StartsWith("/"))
                    image = string.Concat("/", image);

                var encoder = new BmpBitmapEncoder();
                var packApplicationComponent = string.Format(
                    "pack://application:,,,/{0};component{1}",
                    GetType().Assembly.GetName().Name,
                    image);
                encoder.Frames.Add(BitmapFrame.Create(new Uri(packApplicationComponent)));
                encoder.Save(memoryStream);
                return PictureConverter.ImageToPictureDisp(bitmap);
            }
        }

        protected virtual IPictureDisp GetPicture(Image fromImage)
        {
            return PictureConverter.ImageToPictureDisp(fromImage);
        }

        protected virtual IPictureDisp GetPicture(Icon fromIcon)
        {
            return PictureConverter.IconToPictureDisp(fromIcon);
        }

        /// <summary>
        /// Notifies subscribers of the property change.
        /// </summary>
        /// <typeparam name="TProperty">The type of the property.</typeparam>
        /// <param name="property">The property expression.</param>
        protected virtual void OnPropertyChanged<TProperty>(Expression<Func<TProperty>> property)
        {
            RaisePropertyChanged(property);
        }

        /// <summary>
        /// Notifies subscribers of the property change.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        protected virtual void OnPropertyChanged(string propertyName)
        {
            RaisePropertyChanged(propertyName);
        }

        /// <summary>
        /// Notifies subscribers of the property change.
        /// </summary>
        /// <typeparam name="TProperty">The type of the property.</typeparam>
        /// <param name="property">The property expression.</param>
        protected virtual void RaisePropertyChanged<TProperty>(Expression<Func<TProperty>> property)
        {
            var memberExpression = property.Body as MemberExpression;
            if (memberExpression != null)
            {
                RaisePropertyChanged(memberExpression.Member.Name);
            }
        }

        /// <summary>
        /// Notifies subscribers of the property change.
        /// </summary>
        /// <param name="propertyName">Name of the property.</param>
        protected virtual void RaisePropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
    }
}
