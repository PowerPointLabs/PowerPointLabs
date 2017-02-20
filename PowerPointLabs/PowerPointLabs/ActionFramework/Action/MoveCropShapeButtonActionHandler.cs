using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("MoveCropShapeButton")]
    class MoveCropShapeButtonActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLab.CropToShape.Crop(this.GetCurrentSelection());
        }
    }
}
