using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.PasteLab
{
    [ExportSupertipRibbonId("PasteToFillSlideButton")]
    class PasteToFillSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PasteToFillSlideSupertip;
        }
    }
}
