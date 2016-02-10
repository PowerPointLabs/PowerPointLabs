using PowerPointLabs.ActionFramework.Action;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    public class ActionHandlerFactory : BaseHandlerFactory<ActionHandler>
    {
        protected override ActionHandler GetEmptyHandler()
        {
            return new EmptyActionHandler();
        }
    }
}
