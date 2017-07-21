using System;

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
                Views.ErrorDialogBox.ShowDialog("PowerPointLabs", e.Message, e);
                return false;
            }
        }

        protected abstract bool GetEnabled(string ribbonId);
    }
}
