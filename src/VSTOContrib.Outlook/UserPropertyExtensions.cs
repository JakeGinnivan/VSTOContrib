using System;
using Microsoft.Office.Interop.Outlook;
using VSTOContrib.Core.Extensions;

namespace VSTOContrib.Outlook
{
    /// <summary>
    /// Helper extension methods to simplify dealing with OutlookItem.UserProperties.
    /// </summary>
    public static class UserPropertyExtensions
    {
        /// <summary>
        /// Gets the property value for a <see cref="_ContactItem">_ContactItem</see> user property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="contactItem">The contact item.</param>
        /// <param name="name">The name of the user property.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="create">if set to <c>false</c> the property will not be created if it doesn't exist.</param>
        /// <param name="converter">The converter to use to convert the object to.</param>
        /// <param name="defaultValue">The default value to use if user property not found.</param>
        /// <returns>User property vlaue or default</returns>
        public static T GetPropertyValue<T>(this _ContactItem contactItem, string name, OlUserPropertyType type, bool create, Func<object, T> converter, T defaultValue)
        {
            using (var userProperties = contactItem.UserProperties.WithComCleanup())
                return GetPropertyValue(userProperties.Resource, name, type, create, converter, defaultValue);
        }

        /// <summary>
        /// Gets the property value for a <see cref="_AppointmentItem">_AppointmentItem</see> user property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="appointment">The contact item.</param>
        /// <param name="name">The name of the user property.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="create">if set to <c>false</c> the property will not be created if it doesn't exist.</param>
        /// <param name="converter">The converter to use to convert the object to.</param>
        /// <param name="defaultValue">The default value to use if user property not found.</param>
        /// <returns>User property vlaue or default</returns>
        public static T GetPropertyValue<T>(this _AppointmentItem appointment, string name, OlUserPropertyType type, bool create, Func<object, T> converter, T defaultValue)
        {
            using (var userProperties = appointment.UserProperties.WithComCleanup())
                return GetPropertyValue(userProperties.Resource, name, type, create, converter, defaultValue);
        }

        /// <summary>
        /// Gets the property value for a <see cref="_TaskItem">_TaskItem</see> user property.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="task">The contact item.</param>
        /// <param name="name">The name of the user property.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="create">if set to <c>false</c> the property will not be created if it doesn't exist.</param>
        /// <param name="converter">The converter to use to convert the object to.</param>
        /// <param name="defaultValue">The default value to use if user property not found.</param>
        /// <returns>User property vlaue or default</returns>
        public static T GetPropertyValue<T>(this _TaskItem task, string name, OlUserPropertyType type, bool create, Func<object, T> converter, T defaultValue)
        {
            using (var userProperties = task.UserProperties.WithComCleanup())
                return GetPropertyValue(userProperties.Resource, name, type, create, converter, defaultValue);
        }

        /// <summary>
        /// Gets the user property value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="userProperties">The user properties.</param>
        /// <param name="name">The name of the user property.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="create">if set to <c>false</c> the property will not be created if it doesn't exist.</param>
        /// <param name="converter">The converter to use to convert the object to.</param>
        /// <param name="defaultValue">The default value to use if user property not found.</param>
        /// <returns>User property vlaue or default</returns>
        private static T GetPropertyValue<T>(UserProperties userProperties, string name, OlUserPropertyType type, bool create, Func<object, T> converter, T defaultValue)
        {
            using (var property = userProperties.Find(name, true).WithComCleanup())
            {
                var format = type == OlUserPropertyType.olInteger ? OlFormatNumber.olFormatNumberAllDigits : Type.Missing;

                if (property.Resource == null && create)
                    userProperties.Add(name, type, false, format).ReleaseComObject();

                if (property.Resource == null)
                    return defaultValue;

                var value = property.Resource.Value;
                return converter(value);
            }
        }

        /// <summary>
        /// Sets the user property value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="contactItem">The contact item.</param>
        /// <param name="name">The name of the user property to set.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="value">The value to set.</param>
        /// <param name="addToFolder">if set to <c>true</c> add to containing folder. Enables search/display column for user property.</param>
        public static void SetPropertyValue<T>(this _ContactItem contactItem, string name, OlUserPropertyType type, T value, bool addToFolder)
        {
            using (var userProperties = contactItem.UserProperties.WithComCleanup())
                SetPropertyValue(userProperties.Resource, name, type, value, addToFolder);
        }

        /// <summary>
        /// Sets the user property value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="taskitem">The contact item.</param>
        /// <param name="name">The name of the user property to set.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="value">The value to set.</param>
        /// <param name="addToFolder">if set to <c>true</c> add to containing folder. Enables search/display column for user property.</param>
        public static void SetPropertyValue<T>(this _TaskItem taskitem, string name, OlUserPropertyType type, T value, bool addToFolder)
        {
            using (var userProperties = taskitem.UserProperties.WithComCleanup())
                SetPropertyValue(userProperties.Resource, name, type, value, addToFolder);
        }

        /// <summary>
        /// Sets the user property value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="contactItem">The appointment item.</param>
        /// <param name="name">The name of the user property to set.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="value">The value to set.</param>
        /// <param name="addToFolder">if set to <c>true</c> add to containing folder. Enables search/display column for user property.</param>
        public static void SetPropertyValue<T>(this _AppointmentItem contactItem, string name, OlUserPropertyType type, T value, bool addToFolder)
        {
            using (var userProperties = contactItem.UserProperties.WithComCleanup())
                SetPropertyValue(userProperties.Resource, name, type, value, addToFolder);
        }

        /// <summary>
        /// Sets the user property value.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="userProperties">The user properties collection to set user property for.</param>
        /// <param name="name">The name of the user property to set.</param>
        /// <param name="type">The type of the user property.</param>
        /// <param name="value">The value to set.</param>
        /// <param name="addToFolder">if set to <c>true</c> add to containing folder. Enables search/display column for user property.</param>
        private static void SetPropertyValue<T>(UserProperties userProperties, string name, OlUserPropertyType type, T value, bool addToFolder)
        {
            using (var property = userProperties.Find(name, true).WithComCleanup())
            {
                var format = type == OlUserPropertyType.olInteger ? OlFormatNumber.olFormatNumberAllDigits : Type.Missing;

                if (property.Resource == null) 
                using (var newProperty = userProperties.Add(name, type, addToFolder, format).WithComCleanup())
                {
                    newProperty.Resource.Value = value;
                }
                else
                    property.Resource.Value = value;
            }
        }
    }
}
