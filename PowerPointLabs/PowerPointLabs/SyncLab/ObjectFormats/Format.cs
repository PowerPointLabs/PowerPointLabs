using System;
using System.Drawing;
using System.Reflection;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    // Class to facilitate accessing format's functions
    public class Format
    {
        readonly Type format;

        public Format(Type format)
        {
            this.format = format;
        }

        public Type FormatType
        {
            get
            {
                return format;
            }
        }

        public bool CanCopy(Shape formatShape)
        {
            MethodInfo method = format.GetMethod("CanCopy", BindingFlags.Public | BindingFlags.Static);
            return (bool)method.Invoke(null, new Object[] { formatShape });
        }

        public void SyncFormat(Shape formatShape, Shape newShape)
        {
            MethodInfo method = format.GetMethod("SyncFormat", BindingFlags.Public | BindingFlags.Static);
            method.Invoke(null, new Object[] { formatShape, newShape });
        }

        public Bitmap DisplayImage(Shape formatShape)
        {
            MethodInfo method = format.GetMethod("DisplayImage", BindingFlags.Public | BindingFlags.Static);
            return (Bitmap)method.Invoke(null, new Object[] { formatShape, });
        }
    }
}
