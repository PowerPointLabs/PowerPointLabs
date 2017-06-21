using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportSupertipRibbonId(TextCollection.PasteToFillSlideTag)]
    class PasteToFillSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PasteToFillSlideSupertip;
        }
    }
}
