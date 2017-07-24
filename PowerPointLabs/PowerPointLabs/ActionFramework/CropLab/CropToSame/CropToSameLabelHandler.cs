using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportLabelRibbonId(CropLabText.CropToSameDimensionsTag)]
    class CropToSameLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CropLabText.CropToSameButtonLabel;
        }
    }
}
