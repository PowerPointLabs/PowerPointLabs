using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportLabelRibbonId(CropLabText.CropOutPaddingTag)]
    class CropOutPaddingLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CropLabText.CropOutPaddingButtonLabel;
        }
    }
}
