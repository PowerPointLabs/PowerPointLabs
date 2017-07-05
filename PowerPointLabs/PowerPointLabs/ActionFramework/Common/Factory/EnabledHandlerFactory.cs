using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for EnabledHandler
    /// </summary>
    class EnabledHandlerFactory : BaseHandlerFactory<EnabledHandler>
    {
        protected override EnabledHandler GetEmptyHandler()
        {
            return new EmptyEnabledHandler();
        }
    }
}
