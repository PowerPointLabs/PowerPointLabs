
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningLabMenu
{
    [ExportSupertipRibbonId(ELearningLabText.RibbonMenuId)]
    class ELearningLabMenuSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ELearningLabText.RibbonMenuSupertip;
        }
    }
}
