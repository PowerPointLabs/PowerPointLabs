using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportLabelRibbonId(TextCollection.HighlightTextTag)]
    class HighlightTextLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HighlightTextFragmentsButtonLabel;
        }
    }
}
