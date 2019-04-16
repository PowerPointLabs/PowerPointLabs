using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;
using PowerPointLabs.TextCollection;

using TestInterface;
using System.Windows;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ShapesLabTest : BaseFunctionalTest
    {
        private const int SaveShapesShapesSlide = 3;
        private const int SaveShapesTestSlide = 4;
        private const int SaveShapesExpSlide = 5;
        private const int AddShapesShapesSlide = 6;
        private const int AddShapesTestSlide = 7;
        private const int AddShapesExpSlide = 8;
        private const int AddShapesPlaceholderSlide = 9;

        //Check clipboard restored
        private const int SaveShapesClipboardRestoredActualSlide = 11;
        private const int SaveShapesClipboardRestoredTestSlide = 12;
        private const int SaveShapesClipboardRestoredExpSlide = 13;
        private const int AddShapesClipboardRestoredActualSlide = 14;
        private const int AddShapesClipboardRestoredTestSlide = 15;
        private const int AddShapesClipboardRestoredExpSlide = 16;


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

            TestSaveShapesToShapesLab(shapesLab, SaveShapesShapesSlide, SaveShapesTestSlide, SaveShapesExpSlide);
            TestImportLibraryAndShape(shapesLab);
            TestSaveShapesToShapesLabWithAddShapesButton(shapesLab, AddShapesShapesSlide, AddShapesTestSlide, AddShapesExpSlide);
            TestSavePlaceholderToShapesLabWithAddShapesButton(shapesLab, AddShapesPlaceholderSlide);
            IsClipboardRestoredAfterSaveShape(shapesLab, SaveShapesClipboardRestoredActualSlide, SaveShapesClipboardRestoredTestSlide, SaveShapesClipboardRestoredExpSlide);
            IsClipboardRestoredAfterAddShape(shapesLab, AddShapesClipboardRestoredActualSlide, AddShapesClipboardRestoredTestSlide,
                AddShapesClipboardRestoredExpSlide);
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

        private void SaveShapesToShapesLab(IShapesLabController shapesLab, int shapesSlideNum, int testSlideNum)
        {
            PpOperations.SelectSlide(shapesSlideNum);
            PpOperations.SelectShapesByPrefix("selectMe");
            ExpectAddShapeButtonEnabled(shapesLab);
            // save shapes
            shapesLab.SaveSelectedShapes();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(testSlideNum);
            AddShapesToSlideFromShapesLab(shapesLab, "selectMe1", "Group selectMe1");
        }
        private void TestSaveShapesToShapesLab(IShapesLabController shapesLab, int shapesSlideNum, int testSlideNum, int expSlideNum)
        {
            SaveShapesToShapesLab(shapesLab, shapesSlideNum, testSlideNum);

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(testSlideNum);
            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(expSlideNum);

            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }

        private void SaveShapesToShapesLabWithAddShapesButton(IShapesLabController shapesLab, int shapesSlideNum, int testSlideNum)
        {
            PpOperations.SelectSlide(shapesSlideNum);
            PpOperations.SelectShapesByPrefix("selectMeNow");
            ExpectAddShapeButtonEnabled(shapesLab);

            MessageBoxUtil.ExpectMessageBoxWillNotPopUp(
                            ShapesLabText.ErrorDialogTitle, ShapesLabText.ErrorAddSelectionInvalid,
                            shapesLab.ClickAddShapeButton);

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(testSlideNum);
            AddShapesToSlideFromShapesLab(shapesLab, "selectMeNow1", "Group selectMeNow1");
        }

        private void TestSaveShapesToShapesLabWithAddShapesButton(IShapesLabController shapesLab, int shapesSlideNum, int testSlideNum, int expSlideNum)
        {
            SaveShapesToShapesLabWithAddShapesButton(shapesLab, shapesSlideNum, testSlideNum);
            Microsoft.Office.Interop.PowerPoint.Slide expSlide = PpOperations.SelectSlide(expSlideNum);

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(testSlideNum);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
        }

        private void TestSavePlaceholderToShapesLabWithAddShapesButton(IShapesLabController shapesLab, int shapesSlideNum)
        {
            PpOperations.SelectSlide(shapesSlideNum);
            ExpectAddShapeButtonDisabled(shapesLab);
            PpOperations.SelectShapesByPrefix("Placeholder");
            ExpectAddShapeButtonDisabled(shapesLab);

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                            ShapesLabText.ErrorDialogTitle, ShapesLabText.ErrorAddSelectionInvalid,
                            shapesLab.ClickAddShapeButton);
        }

        private void AddShapesToSlideFromShapesLab(IShapesLabController shapesLab, string shapeName, string expectedShapePrefix)
        {
            Point point = shapesLab.GetShapeForClicking(shapeName);
            // Add shapes from Shapes Lab to slide by double clicking
            DoubleClick(point);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix(expectedShapePrefix);
            Assert.IsTrue(shapes.Count > 0, "Failed to add shapes from Shapes Lab.");
        }

        private void DoubleClick(Point point)
        {
            MouseUtil.SendMouseDoubleClick((int) point.X, (int) point.Y);
        }

        private void IsClipboardRestoredAfterSaveShape(IShapesLabController shapesLab, int actualSlideNum, int testSlideNum, int expSlideNum)
        {
            CheckIfClipboardIsRestored(() => SaveShapesToShapesLab(shapesLab, actualSlideNum, testSlideNum),
                actualSlideNum, "copyMe", expSlideNum, "Expected", "compareMe");
        }

        private void IsClipboardRestoredAfterAddShape(IShapesLabController shapesLab, int actualSlideNum, int testSlideNum, int expSlideNum)
        {
            CheckIfClipboardIsRestored(() => SaveShapesToShapesLabWithAddShapesButton(shapesLab, actualSlideNum, testSlideNum),
                actualSlideNum, "copyMe", expSlideNum, "Expected", "compareMe");
        }

        private void ExpectAddShapeButtonEnabled(IShapesLabController shapesLab)
        {
            ThreadUtil.WaitFor(1000);
            Assert.IsTrue(shapesLab.GetAddShapeButtonStatus());
        }

        private void ExpectAddShapeButtonDisabled(IShapesLabController shapesLab)
        {
            ThreadUtil.WaitFor(1000);
            Assert.IsFalse(shapesLab.GetAddShapeButtonStatus());
        }
    }
}