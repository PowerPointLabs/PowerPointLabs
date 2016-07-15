using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ItemLabel
{
    [ExportItemLabelRibbonId("EffectsLabBlurSelectedGallery")]
    class EffectsLabBlurrinessItemLabelHandler : ItemLabelHandler
    {
        protected override string GetItemLabel(string ribbonId, int selectedIndex)
        {
            return (selectedIndex + 4) + "0% " + TextCollection.EffectsLabBlurItemLabel;
        }
    }
}
