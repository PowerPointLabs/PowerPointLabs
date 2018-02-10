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
            AddShapesToSlideFromShapesLab(shapesLab, "selectMe1", "Group selectMe1");

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(5);

            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }

        private void TestSaveShapesToShapesLabWithAddShapesButton(IShapesLabController shapesLab)
        {
            PpOperations.SelectSlide(6);
            PpOperations.SelectShapesByPrefix("selectMeNow");

            shapesLab.ClickAddShapeButton();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(7);
            AddShapesToSlideFromShapesLab(shapesLab, "selectMeNow1", "Group selectMeNow1");

            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(8);

            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }

        private void AddShapesToSlideFromShapesLab(IShapesLabController shapesLab, string shapeThumbnail, string expectedShapePrefix) 
        {
            IShapesLabLabeledThumbnail thumbnail = shapesLab.GetLabeledThumbnail(shapeThumbnail);
            thumbnail.FinishNameEdit();
            // Add shapes from Shapes Lab to slide by double clicking
            DoubleClick(thumbnail as Control);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix(expectedShapePrefix);
            Assert.IsTrue(shapes.Count > 0, "Failed to add shapes from Shapes Lab.");
        }

        private void DoubleClick(Control target)
        {
            Point pt = target.PointToScreen(new Point(target.Width / 2, target.Height / 2));
            MouseUtil.SendMouseDoubleClick(pt.X, pt.Y);
        }
    }
}
