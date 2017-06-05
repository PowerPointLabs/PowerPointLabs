using System;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetEnabled class
    /// </summary>
    public abstract class EnabledHandler : BaseHandler
    {
        public bool Get(string ribbonId)
        {
            try
            {
                return GetEnabled(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return false;
            }
        }

        protected abstract bool GetEnabled(string ribbonId);
        
        protected bool HasPlaceholderInSelection()
        {
            var selection = this.GetCurrentSelection();
            foreach (Shape shape in selection.ShapeRange)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
