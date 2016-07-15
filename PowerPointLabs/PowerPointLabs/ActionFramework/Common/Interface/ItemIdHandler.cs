using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetItemId call
    /// </summary>
    public abstract class ItemIdHandler
    {
        public string Get(string ribbonId, int index)
        {
            try
            {
                return GetItemId(ribbonId, index);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return "";
            }
        }

        protected abstract string GetItemId(string ribbonId, int index);
    }
}
