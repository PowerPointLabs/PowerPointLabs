using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportLabelRibbonId(ShortcutsLabText.HideShapeTag)]
    class HideShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ShortcutsLabText.HideSelectedShapeLabel;
        }
    }
}
