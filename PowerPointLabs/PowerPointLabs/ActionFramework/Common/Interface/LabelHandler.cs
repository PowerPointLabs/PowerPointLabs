using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetLabel call
    /// </summary>
    public abstract class LabelHandler
    {
        public string Get(string ribbonId, string ribbonTag)
        {
            try
            {
                return GetLabel(ribbonId, ribbonTag);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetLabel(string ribbonId, string ribbonTag);
    }
}
