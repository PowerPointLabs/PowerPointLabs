using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetLabel call
    /// </summary>
    public abstract class LabelHandler : BaseHandler
    {
        public string Get(string ribbonId)
        {
            try
            {
                return GetLabel(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogBox.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetLabel(string ribbonId);
    }
}
