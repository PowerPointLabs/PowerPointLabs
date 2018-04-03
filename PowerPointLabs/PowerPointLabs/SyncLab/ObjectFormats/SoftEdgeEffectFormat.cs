using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class SoftEdgeEffectFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            SoftEdgeFormat softEdge = formatShape.SoftEdge;
            return softEdge.Radius > 0 && softEdge.Type != MsoSoftEdgeType.msoSoftEdgeTypeNone;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync SoftEdgeEffect Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            // bump up soft edge to make effect in display more visible
            
            float oldRadius = formatShape.SoftEdge.Radius;
            MsoSoftEdgeType oldType = formatShape.SoftEdge.Type;

            // must change the type to Type6, not all types allow changing the radius
            // Note: Setting MixedType throws an exception
            formatShape.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType6;
            float minEdge = Math.Min(formatShape.Height, formatShape.Width);
            float threshold = (float) (minEdge * 0.2);
            if (oldRadius < threshold)
            {
                formatShape.SoftEdge.Radius = threshold;
            }
            
            Bitmap image = GraphicsUtil.ShapeToBitmap(formatShape);
            formatShape.SoftEdge.Type = oldType;
            formatShape.SoftEdge.Radius = oldRadius;
            
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                SoftEdgeFormat dest = newShape.SoftEdge;
                SoftEdgeFormat source = formatShape.SoftEdge;

                dest.Type = source.Type;
                dest.Radius = source.Radius;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
