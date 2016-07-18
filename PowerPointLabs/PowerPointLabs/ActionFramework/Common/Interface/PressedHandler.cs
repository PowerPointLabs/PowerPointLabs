using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetPressed call
    /// </summary>
    public abstract class PressedHandler
    {
        public bool Get(string ribbonId, string ribbonTag)
        {
            try
            {
                return GetPressed(ribbonId, ribbonTag);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return false;
            }
        }

        protected abstract bool GetPressed(string ribbonId, string ribbonTag);
    }
}
