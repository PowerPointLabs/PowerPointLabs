using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ItemLabel
{
    class EmptyItemLabelHandler : ItemLabelHandler
    {
        protected override string GetItemLabel(string ribbonId, int index)
        {
            return "";
        }
    }
}
