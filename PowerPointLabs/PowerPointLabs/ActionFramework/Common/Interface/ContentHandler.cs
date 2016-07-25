using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetContent call
    /// </summary>
    public abstract class ContentHandler
    {
        public string Get(string ribbonId)
        {
            try
            {
                return GetContent(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetContent(string ribbonId);
    }
}
