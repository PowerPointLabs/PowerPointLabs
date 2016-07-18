using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId("PositionsLabButton")]
    class PositionsLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId, string ribbonTag)
        {
            return TextCollection.PositionsLabButtonLabel;
        }
    }
}
