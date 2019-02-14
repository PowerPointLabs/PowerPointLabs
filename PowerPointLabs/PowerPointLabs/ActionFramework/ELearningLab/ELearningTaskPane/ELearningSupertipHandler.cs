using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningTaskPane
{
    [ExportSupertipRibbonId(ELearningLabText.ELearningTaskPaneTag)]
    class ELearningSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return ELearningLabText.ELearningTaskPaneSuperTip;
        }
    }
}
