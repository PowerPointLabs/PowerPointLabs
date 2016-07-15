using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ItemCount
{
    [ExportItemCountRibbonId("EffectsLabBlurSelectedGallery")]
    class EffectsLabBlurrinessItemCountHandler : ItemCountHandler
    {
        protected override int GetItemCount(string ribbonId)
        {
            return 7;
        }
    }
}
