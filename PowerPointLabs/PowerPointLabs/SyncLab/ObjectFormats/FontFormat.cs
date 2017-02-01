using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontFormat : ObjectFormat
    {
        public FontFormat(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            this.displayText = "Font";
            this.displayImage = Utils.Graphics.ShapeToImage(shape);
            this.formatShape = shape;
        }

        public override void ApplyTo(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            SyncFormatUtil.SyncFontFormat(shape.TextFrame.TextRange.Font, formatShape.TextFrame.TextRange.Font);
        }
    }
}
