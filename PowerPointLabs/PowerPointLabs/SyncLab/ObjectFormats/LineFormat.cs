using System;
using System.Diagnostics;

using Microsoft.Office.Core;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineFormat : ObjectFormat
    {
        public LineFormat(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            this.displayText = "Line";
            this.displayImage = Utils.Graphics.ShapeToImage(shape);
            this.formatShape = shape;
    }

        public override void ApplyTo(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            SyncFormatUtil.SyncLineFormat(shape.Line, formatShape.Line);
        }
    }
}