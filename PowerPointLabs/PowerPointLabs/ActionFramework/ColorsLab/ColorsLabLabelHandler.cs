using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportLabelRibbonId(TextCollection.ColorsLabTag)]
    class ColorsLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ColorPickerButtonLabel;
        }
    }
}
