using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetSupertip call
    /// </summary>
    public abstract class SupertipHandler : BaseHandler
    {
        public string Get(string ribbonId)
        {
            try
            {
                return GetSupertip(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogBox.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetSupertip(string ribbonId);
    }
}
