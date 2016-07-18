using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles OnAction call
    /// </summary>
    public abstract class ActionHandler
    {
        public void Execute(string ribbonId, string ribbonTag)
        {
            try
            {
                ExecuteAction(ribbonId, ribbonTag);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
            }
        }

        protected abstract void ExecuteAction(string ribbonId, string ribbonTag);
    }
}
