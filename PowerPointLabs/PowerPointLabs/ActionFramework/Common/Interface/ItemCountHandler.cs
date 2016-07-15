using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles GetItemCount call
    /// </summary>
    public abstract class ItemCountHandler
    {
        public int Get(string ribbonId)
        {
            try
            {
                return GetItemCount(ribbonId);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
                return 0;
            }
        }

        protected abstract int GetItemCount(string ribbonId);
    }
}
