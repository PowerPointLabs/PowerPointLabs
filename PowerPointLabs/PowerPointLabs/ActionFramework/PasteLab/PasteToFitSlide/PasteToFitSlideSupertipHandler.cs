using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportSupertipRibbonId(PasteLabText.PasteToFitSlideTag)]
    class PasteToFitSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return PasteLabText.PasteToFitSlideSupertip;
        }
    }
}
