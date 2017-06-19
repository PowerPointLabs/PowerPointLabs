using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportSupertipRibbonId("AnimationLabMenu")]
    class AnimationLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.AnimationLabMenuSupertip;
        }
    }
}
