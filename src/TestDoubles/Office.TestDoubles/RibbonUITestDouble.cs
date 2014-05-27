using Microsoft.Office.Core;

namespace Office.TestDoubles
{
    public class RibbonUITestDouble : IRibbonUI
    {
        public void Invalidate()
        {
            throw new System.NotImplementedException();
        }

        public void InvalidateControl(string ControlID)
        {
        }

        public void InvalidateControlMso(string ControlID)
        {
            throw new System.NotImplementedException();
        }

        public void ActivateTab(string ControlID)
        {
            throw new System.NotImplementedException();
        }

        public void ActivateTabMso(string ControlID)
        {
            throw new System.NotImplementedException();
        }

        public void ActivateTabQ(string ControlID, string Namespace)
        {
            throw new System.NotImplementedException();
        }
    }
}