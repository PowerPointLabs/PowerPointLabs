using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportSupertipRibbonId(TextCollection.PasteIntoGroupTag)]
    class PasteIntoGroupSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PasteIntoGroupSupertip;
        }
    }
}
