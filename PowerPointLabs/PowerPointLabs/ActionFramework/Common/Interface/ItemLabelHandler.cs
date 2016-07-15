using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetItemLabel call
    /// </summary>
    public abstract class ItemLabelHandler
    {
        public string Get(string ribbonId, int index)
        {
            try
            {
                return GetItemLabel(ribbonId, index);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetItemLabel(string ribbonId, int index);
    }
}
