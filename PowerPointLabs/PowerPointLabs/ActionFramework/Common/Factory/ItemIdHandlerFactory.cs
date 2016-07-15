using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.ItemId;

namespace PowerPointLabs.ActionFramework.Common.Factory
{
    /// <summary>
    /// Factory for ItemIdHandler
    /// </summary>
    public class ItemIdHandlerFactory : BaseHandlerFactory<ItemIdHandler>
    {
        protected override ItemIdHandler GetEmptyHandler()
        {
            return new EmptyItemIdHandler();
        }
    }
}
