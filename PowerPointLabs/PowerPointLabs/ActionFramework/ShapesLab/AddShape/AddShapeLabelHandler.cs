using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportLabelRibbonId(ShortcutsLabText.AddCustomShapeTag)]
    class AddShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ShortcutsLabText.AddCustomShapeLabel;
        }
    }
}
