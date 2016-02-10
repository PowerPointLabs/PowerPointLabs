using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    public abstract class ActionHandler
    {
        public void Execute(string ribbonId)
        {
            try
            {
                ExecuteAction(ribbonId);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
            }
        }

        protected abstract void ExecuteAction(string ribbonId);
    }
}
