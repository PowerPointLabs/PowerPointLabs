using System;
using System.Windows.Forms;
using Microsoft.QualityTools.Testing.Fakes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            using (ShimsContext.Create())
            {
                // how I successfully did it:
                // 1 change embedded interop type -> false, so this ref can be copied to
                // fakes gen project
                // in this case, this ref is Microsoft.OfficeInterop.PowerPoint
                // 2 same, change this ref in Unit Testing project to false for embedded interop type,
                // to avoid error when build (otherwise cannot call it from assembly)
                var mock = new Mock<PowerPoint.Selection>();
                mock.Setup(selection => selection.Type).Returns(PowerPoint.PpSelectionType.ppSelectionShapes);
                var desiredSel = mock.Object;
                PowerPointLabs.Models.Fakes.ShimPowerPointCurrentPresentationInfo.CurrentSelectionGet = () => desiredSel;
                System.Windows.Forms.Fakes.ShimMessageBox.ShowString = s => DialogResult.No;
                var r = MessageBox.Show("");
                Assert.AreEqual(DialogResult.No, r);
                var sel = PowerPointLabs.Models.PowerPointCurrentPresentationInfo.CurrentSelection;
                Assert.AreEqual(PowerPoint.PpSelectionType.ppSelectionShapes, sel.Type);

                // Crop to shape
                var mockSel = new Mock<PowerPoint.Selection>();
                mockSel.Setup(selection => selection.Type).Returns(PowerPoint.PpSelectionType.ppSelectionNone);
                System.Windows.Forms.Fakes.ShimMessageBox.ShowStringString = (s, s1) => DialogResult.OK;

                var result = PowerPointLabs.CropToShape.Crop(mockSel.Object);
                Assert.AreEqual(null, result);
            }
        }
    }
}
