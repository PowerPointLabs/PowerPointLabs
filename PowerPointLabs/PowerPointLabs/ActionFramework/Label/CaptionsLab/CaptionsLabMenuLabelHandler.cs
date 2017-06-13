using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.CaptionsLab
{
    [ExportLabelRibbonId("CaptionsLabMenu")]
    class CaptionsLabMenuLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.CaptionsLabMenuLabel;
        }
    }
}
