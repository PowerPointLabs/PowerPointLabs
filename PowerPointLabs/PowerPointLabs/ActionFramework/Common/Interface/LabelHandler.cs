using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    public abstract class LabelHandler
    {
        public string Get(string ribbonId)
        {
            try
            {
                return GetLabel(ribbonId);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetLabel(string ribbonId);
    }
}
