using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportLabelRibbonId(PictureSlidesLabText.PaneTag)]
    class PictureSlidesLabLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return PictureSlidesLabText.RibbonMenuLabel;
        }
    }
}
