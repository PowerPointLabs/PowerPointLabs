using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropToSameButton")]
    class CropToSameActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            CropLab.CropToSame.StartCropToSame(this.GetCurrentSelection());
        }
    }
}
