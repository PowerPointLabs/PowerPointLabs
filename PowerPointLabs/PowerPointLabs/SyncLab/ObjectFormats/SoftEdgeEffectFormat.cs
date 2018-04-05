using System;
using System.Drawing;
using System.Speech.Recognition.SrgsGrammar;
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
            // do not check softEdge.Type, it can sometimes == msoSoftEdgeTypeNone when there is a soft edge
            return softEdge.Radius > 0;
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
            // bump up soft edge to make effect in preview more visible
            // work on a duplicate to avoid complex control flow to revert SoftEdge.Type
            // see comments in Sync(..) for more details
            Shape duplicate = formatShape.Duplicate()[1];
            duplicate.SoftEdge.Type = MsoSoftEdgeType.msoSoftEdgeType6;
            float minEdge = Math.Min(formatShape.Height, formatShape.Width);
            float threshold = (float) (minEdge * 0.2);
            if (duplicate.SoftEdge.Radius < threshold)
            {
                formatShape.SoftEdge.Radius = threshold;
            }
            
            Bitmap image = GraphicsUtil.ShapeToBitmap(duplicate);
            duplicate.Delete();
            
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                SoftEdgeFormat dest = newShape.SoftEdge;
                SoftEdgeFormat source = formatShape.SoftEdge;

                // skip setting type, SoftEdgeFormat.Type is not reliable
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
