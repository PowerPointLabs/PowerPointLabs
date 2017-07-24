using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.ZoomLab;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportActionRibbonId(ZoomLabText.ZoomToAreaTag)]
    class ZoomToAreaActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            ZoomToArea.AddZoomToArea();
        }
    }
}
