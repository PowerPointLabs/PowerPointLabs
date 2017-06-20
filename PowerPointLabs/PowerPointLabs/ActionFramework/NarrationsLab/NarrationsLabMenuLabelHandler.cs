using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.NarrationsLab
{
    [ExportLabelRibbonId(TextCollection.NarrationsLabMenuId)]
    class NarrationsLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.NarrationsLabMenuLabel;
        }
    }
}
