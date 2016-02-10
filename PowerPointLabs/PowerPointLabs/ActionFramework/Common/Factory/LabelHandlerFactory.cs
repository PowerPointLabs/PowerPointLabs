using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Label;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    public class LabelHandlerFactory : BaseHandlerFactory<LabelHandler>
    {
        protected override LabelHandler GetEmptyHandler()
        {
            return new EmptyLabelHandler();
        }
    }
}
