using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningTaskPane
{
    [ExportLabelRibbonId(ELearningLabText.ELearningTaskPaneTag)]
    class ELearningLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ELearningLabText.ELearningTaskPaneLabel;
        }
    }
}
