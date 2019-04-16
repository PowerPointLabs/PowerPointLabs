using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ELearningLab.ELearningLabMenu
{
    [ExportImageRibbonId(ELearningLabText.RibbonMenuId)]
    class ELearningLabMenuImageHandler : ImageHandler
    {
        protected override Bitmap GetImage(string ribbonId)
        {
            return new Bitmap(Properties.Resources.ELearningLabIcon2);
        }
    }
}
