using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportLabelRibbonId("StepBack")]
    class StepBackLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddZoomOutButtonLabel;
        }
    }
}
