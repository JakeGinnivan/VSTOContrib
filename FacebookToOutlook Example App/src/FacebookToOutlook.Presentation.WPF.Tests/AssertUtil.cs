using System;
using System.Reflection;
using System.Linq.Expressions;
using System.Diagnostics.CodeAnalysis;
using System.ComponentModel;
using Xunit;

namespace FacebookToOutlook.Presentation.WPF.Tests
{
    public static class AssertUtil
    {
        public static void PropertyChangedEvent<T>(T observable, Expression<Func<T, object>> expression, Action raisePropertyChanged)
            where T : INotifyPropertyChanged
        {
            string propertyName = GetProperty(expression).Name;
            int propertyChangedCount = 0;

            observable.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                Assert.Equal(observable, sender);

                if (e.PropertyName == propertyName)
                {
                    propertyChangedCount++;
                }
            };

            raisePropertyChanged();

            Assert.Equal(1, propertyChangedCount);
        }


        [SuppressMessage("Microsoft.Design", "CA1006:DoNotNestGenericTypesInMemberSignatures")]
        private static PropertyInfo GetProperty<TType>(Expression<Func<TType, object>> propertySelector)
        {
            var expression = propertySelector.Body;

            // If the Property returns a ValueType then a Convert is required => Remove it
            if (expression.NodeType == ExpressionType.Convert || expression.NodeType == ExpressionType.ConvertChecked)
            {
                expression = ((UnaryExpression)expression).Operand;
            }

            // If this isn't a member access expression then the expression isn't valid
            var memberExpression = expression as MemberExpression;
            if (memberExpression == null)
            {
                ThrowExpressionArgumentException("propertySelector");
                return null;
            }

            expression = memberExpression.Expression;

            // If the Property returns a ValueType then a Convert is required => Remove it
            if (expression.NodeType == ExpressionType.Convert || expression.NodeType == ExpressionType.ConvertChecked)
            {
                expression = ((UnaryExpression)expression).Operand;
            }

            // Check if the expression is the parameter itself
            if (expression.NodeType != ExpressionType.Parameter)
            {
                ThrowExpressionArgumentException("propertySelector");
            }

            // Finally retrieve the MemberInfo
            var propertyInfo = memberExpression.Member as PropertyInfo;
            if (propertyInfo == null)
            {
                ThrowExpressionArgumentException("propertySelector");
            }

            return propertyInfo;
        }

        private static void ThrowExpressionArgumentException(string argumentName)
        {
            throw new ArgumentException("It's just the simple expression 'x => x.Property' allowed.", argumentName);
        }
    }
}
