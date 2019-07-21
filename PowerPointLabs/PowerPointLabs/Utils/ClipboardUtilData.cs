using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs.Utils
{
    public class ClipboardUtilData
    {
        public PowerPointSlide tempClipboardSlide = null;
        public ShapeRange tempClipboardShapes = null;
        public SlideRange tempPastedSlide = null;
    }
}
