using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportLabelRibbonId(ZoomLabText.DrillDownTag)]
    class DrillDownLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return ZoomLabText.AddZoomInButtonLabel;
        }
    }
}
