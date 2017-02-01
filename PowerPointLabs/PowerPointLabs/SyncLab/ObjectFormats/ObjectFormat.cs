using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace PowerPointLabs.SyncLab
{
    public abstract class ObjectFormat
    {

        public static readonly Size DISPLAY_IMAGE_SIZE = new Size(32, 32);

        protected string displayText;
        protected Image displayImage;

        public abstract void ApplyTo(Shape shape);

        public string DisplayText
        {
            get
            {
                return displayText;
            }
        }

        public Image DisplayImage
        {
            get
            {
                return displayImage;
            }
        }
    }
}
