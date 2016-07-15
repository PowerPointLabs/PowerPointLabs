using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ItemId
{
    class EmptyItemIdHandler : ItemIdHandler
    {
        protected override string GetItemId(string ribbonId, int index)
        {
            return "";
        }
    }
}
