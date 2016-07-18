using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetSupertip call
    /// </summary>
    public abstract class SupertipHandler
    {
        public string Get(string ribbonId, string ribbonTag)
        {
            try
            {
                return GetSupertip(ribbonId, ribbonTag);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetSupertip(string ribbonId, string ribbonTag);
    }
}
