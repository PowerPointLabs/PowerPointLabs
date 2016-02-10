using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Supertip;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    public class SupertipHandlerFactory : BaseHandlerFactory<SupertipHandler>
    {
        protected override SupertipHandler GetEmptyHandler()
        {
            return new EmptySupertipHandler();
        }
    }
}
