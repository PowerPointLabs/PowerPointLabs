using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for PressedHandler
    /// </summary>
    public class PressedHandlerFactory : BaseHandlerFactory<PressedHandler>
    {
        protected override PressedHandler GetEmptyHandler()
        {
            return new EmptyPressedHandler();
        }
    }
}
