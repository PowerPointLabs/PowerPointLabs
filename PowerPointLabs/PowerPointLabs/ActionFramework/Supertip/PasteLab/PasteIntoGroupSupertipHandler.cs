using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.PasteLab
{
    [ExportSupertipRibbonId("PasteIntoGroupButton")]
    class PasteIntoGroupSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PasteIntoGroupSupertip;
        }
    }
}
