using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles OnCheckBoxAction call
    /// </summary>
    public abstract class CheckBoxActionHandler : BaseActionHandler
    {
        public void Execute(string ribbonId, bool pressed)
        {
            try
            {
                ExecuteCheckBoxAction(ribbonId, pressed);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
            }
        }

        protected abstract void ExecuteCheckBoxAction(string ribbonId, bool pressed);
    }
}
