using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportLabelRibbonId(TextCollection1.PictureSlidesLabTag)]
    class PictureSlidesLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection1.PictureSlidesLabMenuLabel;
        }
    }
}
