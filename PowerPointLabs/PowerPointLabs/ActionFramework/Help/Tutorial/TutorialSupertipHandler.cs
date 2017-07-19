using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportSupertipRibbonId(TextCollection.TutorialTag)]
    class TutorialSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.TutorialButtonSupertip;
        }
    }
}
