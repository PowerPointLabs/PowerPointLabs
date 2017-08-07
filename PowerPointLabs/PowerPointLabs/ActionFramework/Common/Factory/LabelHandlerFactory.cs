using PowerPointLabs.ActionFramework.Common.Handlers;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for LabelHandler
    /// </summary>
    public class LabelHandlerFactory : BaseHandlerFactory<LabelHandler>
    {
        protected override LabelHandler GetEmptyHandler()
        {
            return new EmptyLabelHandler();
        }
    }
}
