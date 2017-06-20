using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.NarrationsLab
{
    [ExportSupertipRibbonId(TextCollection.NarrationsLabMenuId)]
    class NarrationsLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.NarrationsLabMenuSupertip;
        }
    }
}
