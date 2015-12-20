using System;
using PowerPointLabs.ImagesLab.Handler.Effect;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ImagesLab.Util
{
    public class ShapeUtil
    {
        public static void ChangeName(PowerPoint.Shape shape, EffectName effectName, string shapeNamePrefix)
        {
            shape.Name = shapeNamePrefix + "_" + effectName + "_" + Guid.NewGuid().ToString().Substring(0, 7);
        }

        public static void AddTag(PowerPoint.Shape shape, string tagName, String value)
        {
            if (StringUtil.IsEmpty(shape.Tags[tagName]) && value != null)
            {
                shape.Tags.Add(tagName, value);
            }
        }
    }
}
