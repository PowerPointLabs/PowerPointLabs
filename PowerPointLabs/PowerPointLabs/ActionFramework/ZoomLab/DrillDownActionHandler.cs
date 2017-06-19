using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportActionRibbonId("DrillDown")]
    class DrillDownActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            AutoZoom.AddDrillDownAnimation();
        }
    }
}
