using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.ItemCount;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for ItemCountHandler
    /// </summary>
    public class ItemCountHandlerFactory : BaseHandlerFactory<ItemCountHandler>
    {
        protected override ItemCountHandler GetEmptyHandler()
        {
            return new EmptyItemCountHandler();
        }
    }
}
