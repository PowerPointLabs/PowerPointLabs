using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportLabelRibbonId(TextCollection1.AnimationLabMenuId)]
    class AnimationLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return AnimationLabText.AnimationLabMenuLabel;
        }
    }
}
