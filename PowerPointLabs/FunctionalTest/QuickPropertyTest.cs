using System;
using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class QuickPropertyTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "QuickProperties.pptx";
        }

        [TestMethod]
        public void FT_QuickPropertyTest()
        {
            PpOperations.SelectSlide(4);

            var shape = PpOperations.SelectShapes("ffs")[1];

            var x = PpOperations.PointsToScreenPixelsX(shape.Left + shape.Width / 2);
            var y = PpOperations.PointsToScreenPixelsY(shape.Top + shape.Height / 2);
            MouseUtil.SendMouseDoubleClick(x, y);

            ThreadUtil.WaitFor(2000);

            // Spy++ helps to look into the handles
            var pptHandle = NativeUtil.FindWindow("PPTFrameClass", null);
            Assert.AreNotEqual(IntPtr.Zero, pptHandle, "Failed to find PowerPoint handle.");

            var dockRightHandle = 
                NativeUtil.FindWindowEx(pptHandle, IntPtr.Zero, "MsoCommandBarDock", "MsoDockRight");
            Assert.AreNotEqual(IntPtr.Zero, dockRightHandle, "Failed to find Dock Right handle.");

            // AKA property handle
            var formatObjHandle =
                NativeUtil.FindWindowEx(dockRightHandle, IntPtr.Zero, "MsoCommandBar", "Format Object");
            Assert.AreNotEqual(IntPtr.Zero, formatObjHandle, "Failed to find Property handle.");
        }
    }
}
