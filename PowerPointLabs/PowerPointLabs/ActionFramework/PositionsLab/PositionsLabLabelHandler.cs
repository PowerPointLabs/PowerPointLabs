using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PositionsLab
{
    [ExportLabelRibbonId(TextCollection.PositionsLabTag)]
    class PositionsLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PositionsLabButtonLabel;
        }
    }
}
