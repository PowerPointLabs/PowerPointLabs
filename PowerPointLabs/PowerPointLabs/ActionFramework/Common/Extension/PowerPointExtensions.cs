using System;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public static class PowerPointExtensions
    {

        public static void SafeDelete(this ShapeRange shapeRange)
        {
            shapeRange.Delete();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapeRange);
            GC.Collect();
        }

        /// <summary>
        /// Releases all references to <seealso cref="Shape"/> before calling GC to collect.
        /// Required for protection against shape corruption from undo.
        /// </summary>
        /// <param name="shape"></param>
        public static void SafeDelete(this Shape shape)
        {
            shape.Delete();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape);
            GC.Collect();
        }
    }
}
