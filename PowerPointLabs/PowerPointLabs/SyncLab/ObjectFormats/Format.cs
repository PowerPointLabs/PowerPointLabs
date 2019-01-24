using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    public abstract class Format
    {
        
        public abstract bool CanCopy(Shape formatShape);
        
        public abstract void SyncFormat(Shape formatShape, Shape newShape);
        
        public abstract Bitmap DisplayImage(Shape formatShape);

        public override bool Equals(object obj)
        {
            return obj != null && GetType() == obj.GetType();
        }

        public override int GetHashCode()
        {
            return GetType().ToString().GetHashCode();
        }
    }
}