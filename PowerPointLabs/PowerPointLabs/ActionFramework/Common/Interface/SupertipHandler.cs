using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    public abstract class SupertipHandler
    {
        public string Get(string ribbonId)
        {
            try
            {
                return GetSupertip(ribbonId);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetSupertip(string ribbonId);
    }
}
