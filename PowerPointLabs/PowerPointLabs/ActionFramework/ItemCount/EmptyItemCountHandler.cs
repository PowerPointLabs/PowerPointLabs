using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ItemCount
{
    class EmptyItemCountHandler : ItemCountHandler
    {
        protected override int GetItemCount(string ribbonId)
        {
            return 0;
        }
    }
}
