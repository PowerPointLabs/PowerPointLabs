using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportLabelRibbonId(TextCollection.HighlightBackgroundTag)]
    class HighlightBackgroundLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HighlightBulletsBackgroundButtonLabel;
        }
    }
}
