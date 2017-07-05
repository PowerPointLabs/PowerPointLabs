using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportImageRibbonId(TextCollection.EffectsLabMenuId)]
    class EffectsLabMenuImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.EffectsLab);
        }
    }
}
