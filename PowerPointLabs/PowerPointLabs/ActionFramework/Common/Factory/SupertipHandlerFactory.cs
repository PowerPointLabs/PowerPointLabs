using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for SupertipHandler
    /// </summary>
    public class SupertipHandlerFactory : BaseHandlerFactory<SupertipHandler>
    {
        protected override SupertipHandler GetEmptyHandler()
        {
            return new EmptySupertipHandler();
        }
    }
}
