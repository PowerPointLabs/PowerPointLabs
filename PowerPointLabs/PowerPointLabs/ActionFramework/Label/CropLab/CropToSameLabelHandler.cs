using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.CropLab
{
    [ExportLabelRibbonId(TextCollection.CropToSameDimensionsTag + TextCollection.RibbonButton)]
    class CropToSameLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.CropToSameButtonLabel;
        }
    }
}
