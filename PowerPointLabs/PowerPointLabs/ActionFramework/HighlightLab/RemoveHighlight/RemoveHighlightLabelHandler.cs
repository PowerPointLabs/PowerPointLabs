using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportLabelRibbonId(HighlightLabText.RemoveHighlightTag)]
    class RemoveHighlightLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HighlightLabText.RemoveHighlightButtonLabel;
        }
    }
}
