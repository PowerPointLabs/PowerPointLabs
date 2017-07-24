using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportSupertipRibbonId(HighlightLabText.RemoveHighlightTag)]
    class RemoveHighlightSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return HighlightLabText.RemoveHighlightButtonSupertip;
        }
    }
}
