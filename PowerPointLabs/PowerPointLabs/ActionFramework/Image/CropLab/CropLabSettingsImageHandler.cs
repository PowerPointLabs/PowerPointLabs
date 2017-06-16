using System.Drawing;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Image.CropLab
{
    [ExportImageRibbonId(TextCollection.CropLabSettingsId + TextCollection.RibbonButton)]
    class CropLabSettingsImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.AgendaSettings);
        }
    }
}
