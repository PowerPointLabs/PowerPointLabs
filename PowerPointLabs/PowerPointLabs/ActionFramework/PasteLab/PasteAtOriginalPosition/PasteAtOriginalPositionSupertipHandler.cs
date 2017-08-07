using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportSupertipRibbonId(PasteLabText.PasteAtOriginalPositionTag)]
    class PasteAtOriginalPositionSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return PasteLabText.PasteAtOriginalPositionSupertip;
        }
    }
}
