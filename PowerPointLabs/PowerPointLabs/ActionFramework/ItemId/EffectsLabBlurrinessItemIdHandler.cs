using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ItemId
{
    [ExportItemIdRibbonId("EffectsLabBlurSelectedGallery")]
    class EffectsLabBlurrinessItemIdHandler : ItemIdHandler
    {
        protected override string GetItemId(string ribbonId, int selectedIndex)
        {
            return ribbonId.Replace("Gallery", (selectedIndex + 4) + "0");
        }
    }
}
