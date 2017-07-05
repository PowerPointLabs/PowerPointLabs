using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.RibbonGroups
{
    [ExportLabelRibbonId(TextCollection.AnimationsGroupId)]
    class AnimationsGroupLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AnimationsGroupLabel;
        }
    }
}
