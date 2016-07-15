using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles OnGalleryAction call
    /// </summary>
    public abstract class GalleryActionHandler
    {
        public void Execute(string ribbonId, string selectedId, int selectedIndex)
        {
            try
            {
                ExecuteGalleryAction(ribbonId, selectedId, selectedIndex);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
            }
        }

        protected abstract void ExecuteGalleryAction(string ribbonId, string selectedId, int selectedIndex);
    }
}
