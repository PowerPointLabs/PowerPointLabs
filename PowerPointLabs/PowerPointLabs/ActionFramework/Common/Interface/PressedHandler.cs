using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetPressed call
    /// </summary>
    public abstract class PressedHandler : BaseHandler
    {
        public bool Get(string ribbonId)
        {
            try
            {
                return GetPressed(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogBox.ShowDialog("PowerPointLabs", e.Message, e);
                return false;
            }
        }

        protected abstract bool GetPressed(string ribbonId);
    }
}
