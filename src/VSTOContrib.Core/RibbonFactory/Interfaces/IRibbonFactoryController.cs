using System;
using System.Linq.Expressions;
using Microsoft.Office.Core;

namespace VSTOContrib.Core.RibbonFactory.Interfaces
{
    public interface IRibbonFactoryController : IDisposable
    {
        void Initialise(IViewProvider viewProvider);
        string GetCustomUI(string ribbonId);
        object InvokeGet(IRibbonControl control, Expression<Action> caller, params object[] parameters);
        void Invoke(IRibbonControl control, Expression<Action> caller, params object[] parameters);
        void RibbonLoaded(IRibbonUI ribbonUi);
    }
}