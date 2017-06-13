using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.PasteLab
{
    [ExportSupertipRibbonId("PasteAtOriginalPositionButton")]
    class PasteAtOriginalPositionSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PasteAtOriginalPositionSupertip;
        }
    }
}
