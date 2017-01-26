using Microsoft.Office.Core;
using System;
using System.Diagnostics;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineFormat : ObjectFormat
    {
#pragma warning disable 0618
        #region Properties
        private readonly MsoArrowheadLength beginArrowheadLength;
        private readonly MsoArrowheadStyle beginArrowheadStyle;
        private readonly MsoArrowheadWidth beginArrowheadWidth;
        private readonly MsoLineDashStyle dashStyle;
        private readonly MsoArrowheadLength endArrowheadLength;
        private readonly MsoArrowheadStyle endArrowheadStyle;
        private readonly MsoArrowheadWidth endArrowheadWidth;
        private readonly Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor;
        private readonly MsoTriState insetPen;
        private readonly MsoPatternType pattern;
        private readonly MsoLineStyle style;
        private readonly float transparency;
        private readonly MsoTriState visible;
        private readonly float weight;
        #endregion

        public LineFormat(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            this.displayText = "Line";
            this.displayImage = Utils.Graphics.ShapeToImage(shape);
            Microsoft.Office.Interop.PowerPoint.LineFormat format = shape.Line;

            this.beginArrowheadLength = format.BeginArrowheadLength;
            this.beginArrowheadStyle = format.BeginArrowheadStyle;
            this.beginArrowheadWidth = format.BeginArrowheadWidth;
            this.dashStyle = format.DashStyle;
            this.endArrowheadLength = format.EndArrowheadLength;
            this.endArrowheadStyle = format.EndArrowheadStyle;
            this.endArrowheadWidth = format.EndArrowheadWidth;
            this.foreColor = format.ForeColor;
            this.insetPen = format.InsetPen;
            this.pattern = format.Pattern;
            this.style = format.Style;
            this.transparency = format.Transparency;
            this.visible = format.Visible;
            this.weight = format.Weight;
    }

        public override void ApplyTo(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Microsoft.Office.Interop.PowerPoint.LineFormat format = shape.Line;
            format.BeginArrowheadLength = this.beginArrowheadLength;
            format.BeginArrowheadStyle = this.beginArrowheadStyle;
            format.BeginArrowheadWidth = this.beginArrowheadWidth;
            format.DashStyle = this.dashStyle;
            format.EndArrowheadLength = this.endArrowheadLength;
            format.EndArrowheadStyle = this.endArrowheadStyle;
            format.EndArrowheadWidth = this.endArrowheadWidth;
            format.ForeColor = this.foreColor;
            format.InsetPen = this.insetPen;
            format.Pattern = this.pattern;
            format.Style = this.style;
            format.Transparency = this.transparency;
            format.Visible = this.visible;
            format.Weight = this.weight;
        }
    }
}