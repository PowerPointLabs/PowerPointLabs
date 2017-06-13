using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.ZoomLab
{
    [ExportLabelRibbonId("ZoomLabMenu")]
    class ZoomLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ZoomLabMenuLabel;
        }
    }
}
