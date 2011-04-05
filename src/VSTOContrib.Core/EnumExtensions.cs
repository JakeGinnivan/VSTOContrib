using System;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;

namespace VSTOContrib.Core
{
    ///<summary>
    /// Extension methods to extend enum functionality
    ///</summary>
    public static class EnumExtensions
    {
        /// <summary>
        /// Gets the enum description.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string GetEnumDescription(this Enum value)
        {
            var fi = value.GetType().GetField(value.ToString());
            var attributes =
              (DescriptionAttribute[])fi.GetCustomAttributes
              (typeof(DescriptionAttribute), false);
            return (attributes.Length > 0) ? attributes[0].Description : value.ToString();
        }

        /// <summary>
        /// Gets the name of the enum.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="description">The description.</param>
        /// <returns></returns>
        public static string GetEnumName(Type value, string description)
        {
            var fis = value.GetFields();
            foreach (var fi in from fi in fis
                               let attributes = (DescriptionAttribute[]) fi.GetCustomAttributes(typeof (DescriptionAttribute), false)
                               where attributes.Length > 0
                               where attributes[0].Description == description
                               select fi)
            {
                return fi.Name;
            }
            return description;
        }

        ///<summary>
        /// Gets a enum from the enum description
        ///</summary>
        ///<param name="description">Enum description</param>
        ///<typeparam name="T">Enum type</typeparam>
        ///<returns>The enum value</returns>
        ///<exception cref="ArgumentException">Thrown if no matching enum exists</exception>
        public static T EnumFromDescription<T>(string description) where T : struct 
        {
            var type = typeof(T);
            var enumName = GetEnumName(type, description);
            return (T)Enum.Parse(type, enumName);
        }

        ///<summary>
        ///  Gets a enum from the enum description
        ///</summary>
        ///<param name="description">Enum description</param>
        ///<param name="enumValue">Output value</param>
        ///<typeparam name="T">Enum type</typeparam>
        ///<returns></returns>
        public static bool TryGetEnumFromDescription<T>(string description, ref T enumValue) where T : struct 
        {
            var type = typeof(T);
            var enumName = GetEnumName(type, description);
            try
            {
                enumValue = (T)Enum.Parse(type, enumName);
                return true;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        ///<summary>
        /// Gets the method name used in an expression. Provides strongly typed method name resolution.
        /// ()=>MyMethod(null, null) - will return MyMethod.
        ///</summary>
        ///<param name="expression">The action calling the method</param>
        ///<returns>The name of the called method</returns>
        public static string GetMethodName(this Expression<Action> expression)
        {
            return ((MethodCallExpression)expression.Body).Method.Name;
        }

        ///<summary>
        /// Gets the method name used in an expression. Provides strongly typed method name resolution.
        /// ()=>MyMethod(null, null) - will return MyMethod.
        ///</summary>
        ///<param name="expression">The action calling the method</param>
        ///<returns>The name of the called method</returns>
        public static string GetMethodName<T>(this Expression<Action<T>> expression)
        {
            return ((MethodCallExpression)expression.Body).Method.Name;
        }
    }
}
