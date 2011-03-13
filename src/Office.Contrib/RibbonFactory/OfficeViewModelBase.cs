using System;
using System.ComponentModel;
using System.Linq.Expressions;

namespace Office.Contrib.RibbonFactory
{
    /// <summary>
    /// View Model base for office ribbon view models
    /// </summary>
    public class OfficeViewModelBase : INotifyPropertyChanged
    {
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
