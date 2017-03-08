using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab
{
    public abstract class ObjectFormat
    {

        public static readonly Size DISPLAY_IMAGE_SIZE = new Size(32, 32);

        protected string displayText;
        protected Image displayImage;
        protected Shape formatShape; // A shape whose format we want

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

        public Shape FormatShape
        {
            get
            {
                return formatShape;
            }
        }
    }
}
