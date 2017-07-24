using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.RibbonGroups
{
    [ExportLabelRibbonId(CommonText.EffectsGroupId)]
    class EffectsGroupLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CommonText.EffectsGroupLabel;
        }
    }
}
