using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for ActionHandler
    /// </summary>
    public class ActionHandlerFactory : BaseHandlerFactory<ActionHandler>
    {
        protected override ActionHandler GetEmptyHandler()
        {
            return new EmptyActionHandler();
        }
    }
}
