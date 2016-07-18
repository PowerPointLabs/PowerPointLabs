using System;

namespace PowerPointLabs.ActionFramework.Common.Interface
{
    /// <summary>
    /// Handler that handles OnCheckBoxAction call
    /// </summary>
    public abstract class CheckBoxActionHandler
    {
        public void Execute(string ribbonId, string ribbonTag, bool pressed)
        {
            try
            {
                ExecuteCheckBoxAction(ribbonId, ribbonTag, pressed);
            }
            catch (Exception e)
            {
                Log.Logger.LogException(e, ribbonId);
                Views.ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
            }
        }

        protected abstract void ExecuteCheckBoxAction(string ribbonId, string ribbonTag, bool pressed);
    }
}
