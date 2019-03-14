using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningLabMenu
{
    [ExportLabelRibbonId(ELearningLabText.RibbonMenuId)]
    class ELearningLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ELearningLabText.RibbonMenuLabel;
        }
    }
}
