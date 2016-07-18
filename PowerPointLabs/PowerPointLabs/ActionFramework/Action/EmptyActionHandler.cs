using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    /// <summary>
    /// the action handler that does nothing
    /// </summary>
    class EmptyActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId, string ribbonTag)
        {
        }
    }
}
