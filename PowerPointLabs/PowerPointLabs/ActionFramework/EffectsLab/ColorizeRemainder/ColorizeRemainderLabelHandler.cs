using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportLabelRibbonId(TextCollection.ColorizeRemainderMenuId)]
    class ColorizeRemainderLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EffectsLabRecolorRemainderButtonLabel;
        }
    }
}
