using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportLabelRibbonId(CropLabText.CropToShapeTag)]
    class CropToShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CropLabText.CropToShapeButtonLabel;
        }
    }
}
