using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportSupertipRibbonId(TextCollection.PasteLabMenuId)]
    class PasteLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PasteLabMenuSupertip;
        }
    }
}
