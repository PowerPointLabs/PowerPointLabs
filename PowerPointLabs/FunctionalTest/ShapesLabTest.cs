using System.Drawing;
using System.Windows.Forms;
using FunctionalTest.util;
using FunctionalTestInterface;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class ShapesLabTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ShapesLab.pptx";
        }

        //use new powerpoint instance to refresh
        //ShapesLabConfig setting for FT
        //
        //every time, shapes lab in FT will use diff
        //shapesRootFolder & default category
        protected override bool IsUseNewPpInstance()
        {
            return true;
        }

        [TestMethod]
        public void FT_ShapesLabTest()
        {
            PpOperations.MaximizeWindow();
            var shapesLab = PplFeatures.ShapesLab;
            shapesLab.OpenPane();

            TestSaveShapesToShapesLab(shapesLab);
            TestImportLibrary(shapesLab);
        }

        private void TestImportLibrary(IShapesLabController shapesLab)
        {
            shapesLab.ImportLibrary(
                PathUtil.GetDocTestPresentationPath("LibraryToImport.pptlabsshapes"));
            var actualShapeDataAfterImport = shapesLab.FetchShapeGalleryPresentationData();
            var expShapeDataAfterImport = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("ExpShapeGalleryAftImport.pptx"));
            PresentationUtil.AssertEqual(expShapeDataAfterImport, actualShapeDataAfterImport);
        }

        private void TestSaveShapesToShapesLab(IShapesLabController shapesLab)
        {
            PpOperations.SelectSlide(3);
            PpOperations.SelectShapesByPrefix("selectMe");
            // save shapes
            shapesLab.SaveSelectedShapes();

            var actualSlide = PpOperations.SelectSlide(4);
            var addedThumbnail = shapesLab.GetLabeledThumbnail("My Shape Untitled");
            addedThumbnail.FinishNameEdit();
            // add shapes back
            DoubleClick(addedThumbnail as Control);
            var shapes = PpOperations.SelectShapesByPrefix("Group My Shape");
            Assert.IsTrue(shapes.Count > 0, "Failed to add shapes from Shapes Lab." +
                                            "UI test is flaky, pls re-run.");

            var expSlide = PpOperations.SelectSlide(5);

            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }

        private void DoubleClick(Control target)
        {
            var pt = target.PointToScreen(new Point(target.Width/2, target.Height/2));
            MouseUtil.SendMouseDoubleClick(pt.X, pt.Y);
        }
    }
}
