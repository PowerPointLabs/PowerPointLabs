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

                if (source.Type == MsoSoftEdgeType.msoSoftEdgeTypeMixed)
                {
                    // skip setting type, setting msoSoftEdgeTypeMixed will throw an error
                    // configuring the settings manually will automatically set the Type to TypeMixed
                    dest.Radius = source.Radius;
                }
                else
                {
                    // skip changing the settings, setting Type will automatically change settings
                    dest.Type = source.Type;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
