using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportLabelRibbonId(TextCollection.ShapesLabTag)]
    class ShapesLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.CustomeShapeButtonLabel;
        }
    }
}
