using System.Drawing;
using System.Windows.Forms;
using TestInterface;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ShapesLabTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ShapesLab\\ShapesLab.pptx";
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
        [TestCategory("FT")]
        public void FT_ShapesLabTest()
        {
            PpOperations.MaximizeWindow();
            IShapesLabController shapesLab = PplFeatures.ShapesLab;
            shapesLab.OpenPane();

            TestSaveShapesToShapesLab(shapesLab);
            TestImportLibraryAndShape(shapesLab);
            TestSaveShapesToShapesLabWithAddShapesButton(shapesLab);
        }

        private void TestImportLibraryAndShape(IShapesLabController shapesLab)
        {
            shapesLab.ImportLibrary(
                PathUtil.GetDocTestPresentationPath("ShapesLab\\LibraryToImport.pptlabsshapes"));
            shapesLab.ImportLibrary(
                PathUtil.GetDocTestPresentationPath("ShapesLab\\ShapeToImport.pptlabsshape"));
            System.Collections.Generic.List<ISlideData> actualShapeDataAfterImport = shapesLab.FetchShapeGalleryPresentationData();
            System.Collections.Generic.List<ISlideData> expShapeDataAfterImport = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath(ExpectedShapeGalleryFileName()));
            PresentationUtil.AssertEqual(expShapeDataAfterImport, actualShapeDataAfterImport);
        }

        private string ExpectedShapeGalleryFileName()
        {
            if (PpOperations.IsOffice2010())
            {
                return "ShapesLab\\ExpShapeGalleryAftImportNonWide.pptx";
            }
            else
            {
                return "ShapesLab\\ExpShapeGalleryAftImport.pptx";
            }
        }

        private void TestSaveShapesToShapesLab(IShapesLabController shapesLab)
        {
            PpOperations.SelectSlide(3);
            PpOperations.SelectShapesByPrefix("selectMe");
            // save shapes
            shapesLab.SaveSelectedShapes();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(4);
            IShapesLabLabeledThumbnail addedThumbnail = shapesLab.GetLabeledThumbnail("selectMe1");
            addedThumbnail.FinishNameEdit();
            // add shapes back
            DoubleClick(addedThumbnail as Control);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix("Group selectMe1");
            Assert.IsTrue(shapes.Count > 0, "Failed to add shapes from Shapes Lab." +
                                            "UI test is flaky, pls re-run.");

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(5);

            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }

        private void DoubleClick(Control target)
        {
            Point pt = target.PointToScreen(new Point(target.Width/2, target.Height/2));
            MouseUtil.SendMouseDoubleClick(pt.X, pt.Y);
        }

        private void TestSaveShapesToShapesLabWithAddShapesButton(IShapesLabController shapesLab)
        {
            PpOperations.SelectSlide(6);
            PpOperations.SelectShapesByPrefix("selectMeNow");
            // Need to perform clicking of button in its own UI thread at ShapesLabController
            // thus clicking cannot be performed directly in test script
            shapesLab.ClickAddShapeButton();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(7);
            IShapesLabLabeledThumbnail addedThumbnail = shapesLab.GetLabeledThumbnail("selectMeNow1");
            addedThumbnail.FinishNameEdit();
            // add shapes back
            DoubleClick(addedThumbnail as Control);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix("Group selectMeNow1");
            Assert.IsTrue(shapes.Count > 0, "Failed to add shapes from Shapes Lab." +
                                            "UI test is flaky, pls re-run.");

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);

            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }
    }
}
