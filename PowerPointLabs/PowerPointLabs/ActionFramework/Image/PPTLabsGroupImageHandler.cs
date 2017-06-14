using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image
{
    [ExportImageRibbonId(
        "AnimationsGroup", "AudioGroup", "EffectsGroup",
        "FormattingGroup", "MoreLabsGroup", "HelpGroup")]
    class PptLabsGroupImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.PptlabsContextMenu);
        }
    }
}
