using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace Office.Contrib.RibbonFactory.Interfaces
{
    internal interface IRibbonFactoryImpl
    {
        IDisposable Initialise(
            string loadMethodName,
            Dictionary<string, Dictionary<string, Expression<Action>>> ribbonElements,
            Func<Type, IRibbonViewModel> ribbonFactory,
            CustomTaskPaneCollection customTaskPaneCollection,
            params Assembly[] assemblies);

        string GetCustomUI(string ribbonId);
        object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters);
        void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters);
        void RibbonLoaded(IRibbonUI ribbonUi);
    }
}