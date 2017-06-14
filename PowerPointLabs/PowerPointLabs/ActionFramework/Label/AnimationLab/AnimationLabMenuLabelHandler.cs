using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.AnimationLab
{
    [ExportLabelRibbonId("AnimationLabMenu")]
    class AnimationLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AnimationLabMenuLabel;
        }
    }
}
