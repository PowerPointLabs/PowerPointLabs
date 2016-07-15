using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.ItemLabel;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for ItemLabelHandler
    /// </summary>
    public class ItemLabelHandlerFactory : BaseHandlerFactory<ItemLabelHandler>
    {
        protected override ItemLabelHandler GetEmptyHandler()
        {
            return new EmptyItemLabelHandler();
        }
    }
}
