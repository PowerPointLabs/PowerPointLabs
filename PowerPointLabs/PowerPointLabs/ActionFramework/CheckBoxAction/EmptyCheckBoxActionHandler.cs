using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CheckBoxAction
{
    /// <summary>
    /// the check box action handler that does nothing
    /// </summary>
    class EmptyCheckBoxActionHandler : CheckBoxActionHandler
    {
        protected override void ExecuteCheckBoxAction(string ribbonId, bool pressed)
        {
        }
    }
}
