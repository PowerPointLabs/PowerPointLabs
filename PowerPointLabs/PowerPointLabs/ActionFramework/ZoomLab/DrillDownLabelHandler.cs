using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportLabelRibbonId("DrillDown")]
    class DrillDownLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AddZoomInButtonLabel;
        }
    }
}
