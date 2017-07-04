using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.RibbonGroups
{
    [ExportLabelRibbonId(TextCollection.FormattingGroupId)]
    class FormattingGroupLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.FormattingGroupLabel;
        }
    }
}
