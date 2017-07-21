using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PositionsLab
{
    [ExportLabelRibbonId(TextCollection1.PositionsLabTag)]
    class PositionsLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return PositionsLabText.RibbonMenuLabel;
        }
    }
}
