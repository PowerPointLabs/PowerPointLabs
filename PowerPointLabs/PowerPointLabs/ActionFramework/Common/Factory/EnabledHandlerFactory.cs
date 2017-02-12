using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Enabled;

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
