using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Pressed;

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
