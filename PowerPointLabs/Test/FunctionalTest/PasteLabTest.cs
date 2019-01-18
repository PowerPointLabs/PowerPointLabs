using System.Collections.Generic;
using System.Drawing;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Test.FunctionalTest
{
    [TestClass]
    public class PasteLabTest : BaseFunctionalTest
    {
        private const string ShapeToCopyPrefix = "selectMe";
        private const string ShapeToClick = "Rectangle 1";
        private const string ShapeToReplace = "Rectangle 5";
        private const string GroupToPaste = "Group 1";
        private const string ShapeToCopyToClipboard = "pictocopy";
        private const string ShapeToCompareCopied = "copied";

        //Slide Numbers
        private const int OriginalPasteToFillSlideSlideNo = 4;
        private const int ExpectedPasteToFillSlideSlideNo = 5;
        private const int OriginalDiagonalPasteToFillSlideSlideNo = 6;
        private const int ExpectedDiagonalPasteToFillSlideSlideNo = 7;
        private const int OriginalMultiplePasteToFillSlideSlideNo = 8;
        private const int ExpectedMultiplePasteToFillSlideSlideNo = 9;
        private const int OriginalGroupPasteToFillSlideSlideNo = 10;
        private const int ExpectedGroupPasteToFillSlideSlideNo = 11;

        private const int OriginalPasteAtCursorSlideNo = 13;
        private const int ExpectedPasteAtCursorSlideNo = 14;
        private const int OriginalPasteAtOriginalSlideNo = 15;
        private const int ExpectedPasteAtOriginalSlideNo = 16;

        private const int OriginalReplaceWithClipboardSlideNo = 18;
        private const int ExpectedReplaceWithClipboardSlideNo = 19;
        private const int OriginalGroupReplaceWithClipboardSlideNo = 20;
        private const int ExpectedGroupReplaceWithClipboardSlideNo = 21;

        private const int OriginalPasteIntoGroupSlideNo = 23;
        private const int ExpectedPasteIntoGroupSlideNo = 24;

        private const int OriginalPasteToFitSlideSlideNo = 26;
        private const int ExpectedPasteToFitSlideSlideNo = 27;
        private const int OriginalDiagonalPasteToFitSlideSlideNo = 28;
        private const int ExpectedDiagonalPasteToFitSlideSlideNo = 29;
        private const int OriginalMultiplePasteToFitSlideSlideNo = 30;
        private const int ExpectedMultiplePasteToFitSlideSlideNo = 31;
        private const int OriginalGroupPasteToFitSlideSlideNo = 32;
        private const int ExpectedGroupPasteToFitSlideSlideNo = 33;

        private const int OrigIsClipboardRestoredReplaceWithClipboardSlideNo = 35;
        private const int ExpIsClipboardRestoredReplaceWithClipboardSlideNo = 36;

        private const int OrigIsClipboardRestoredPasteIntoGroupSlideNo = 37;
        private const int ExpIsClipboardRestoredPasteIntoGroupSlideNo = 38;

        protected override string GetTestingSlideName()
        {
            return "PasteLab\\PasteLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_PasteLabTest()
        {
            PasteToFillSlide(OriginalPasteToFillSlideSlideNo, ExpectedPasteToFillSlideSlideNo);
            PasteToFillSlide(OriginalDiagonalPasteToFillSlideSlideNo, ExpectedDiagonalPasteToFillSlideSlideNo);
            PasteToFillSlide(OriginalMultiplePasteToFillSlideSlideNo, ExpectedMultiplePasteToFillSlideSlideNo);
            PasteToFillSlide(OriginalGroupPasteToFillSlideSlideNo, ExpectedGroupPasteToFillSlideSlideNo);

            PasteAtCursorPosition(OriginalPasteAtCursorSlideNo, ExpectedPasteAtCursorSlideNo);
            PasteAtOriginalPosition(OriginalPasteAtOriginalSlideNo, ExpectedPasteAtOriginalSlideNo);

            ReplaceWithClipboard(OriginalReplaceWithClipboardSlideNo, ExpectedReplaceWithClipboardSlideNo);
            ReplaceWithClipboard(OriginalGroupReplaceWithClipboardSlideNo, ExpectedGroupReplaceWithClipboardSlideNo);

            PasteIntoGroup(OriginalPasteIntoGroupSlideNo, ExpectedPasteIntoGroupSlideNo);

            PasteToFitSlide(OriginalPasteToFitSlideSlideNo, ExpectedPasteToFitSlideSlideNo);
            PasteToFitSlide(OriginalDiagonalPasteToFitSlideSlideNo, ExpectedDiagonalPasteToFitSlideSlideNo);
            PasteToFitSlide(OriginalMultiplePasteToFitSlideSlideNo, ExpectedMultiplePasteToFitSlideSlideNo);
            PasteToFitSlide(OriginalGroupPasteToFitSlideSlideNo, ExpectedGroupPasteToFitSlideSlideNo);

            IsClipboardRestoredReplaceWithClipboard(OrigIsClipboardRestoredReplaceWithClipboardSlideNo, ExpIsClipboardRestoredReplaceWithClipboardSlideNo);
            IsClipboardRestoredPasteIntoGroup(OrigIsClipboardRestoredPasteIntoGroupSlideNo, ExpIsClipboardRestoredPasteIntoGroupSlideNo);
        }

        private void PasteToFillSlide(int originalSlideNo, int expSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PplFeatures.PasteToFillSlide();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteAtCursorPosition(int originalSlideNo, int expSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            RightClick(GetShapesByPrefix(originalSlideNo, ShapeToClick)[1]);
            // wait for awhile for click to register properly
            ThreadUtil.WaitFor(500);
            PplFeatures.PasteAtCursorPosition();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteAtOriginalPosition(int originalSlideNo, int expSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PplFeatures.PasteAtOriginalPosition();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void ReplaceWithClipboard(int originalSlideNo, int expSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PpOperations.SelectShapes(new List<string> { ShapeToReplace });
            PplFeatures.ReplaceWithClipboard();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteIntoGroup(int originalSlideNo, int expSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PpOperations.SelectShape(GroupToPaste);
            PplFeatures.PasteIntoGroup();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteToFitSlide(int originalSlideNo, int expSlideNo)
        {
            var shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PplFeatures.PasteToFitSlide();

            AssertIsSame(originalSlideNo, expSlideNo);
        }
        private void IsClipboardRestoredReplaceWithClipboard(int originalSlideNo, int expSlideNo)
        {
            CheckIfClipboardIsRestored(() =>
            {
                Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(OrigIsClipboardRestoredReplaceWithClipboardSlideNo, ShapeToCopyPrefix);
                // This should be restored to clipboard later
                shapes.Cut();

                PpOperations.SelectShapes(new List<string> { ShapeToReplace });
                PplFeatures.ReplaceWithClipboard();

            }, originalSlideNo, ShapeToCopyPrefix, expSlideNo, "", ShapeToCompareCopied);
        }

        private void IsClipboardRestoredPasteIntoGroup(int originalSlideNo, int expSlideNo)
        {
            CheckIfClipboardIsRestored(() =>
            {
                Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = GetShapesByPrefix(OrigIsClipboardRestoredPasteIntoGroupSlideNo, ShapeToCopyPrefix);
                // This should be restored to clipboard later
                shapes.Cut();

                PpOperations.SelectShape(GroupToPaste);
                PplFeatures.PasteIntoGroup();

            }, originalSlideNo, ShapeToCopyPrefix, expSlideNo, "", ShapeToCompareCopied);
        }

        private void AssertIsSame(int actualSlideNo, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide = PpOperations.SelectSlide(actualSlideNo);
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = PpOperations.SelectSlide(expectedSlideNo);

            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }

        private Microsoft.Office.Interop.PowerPoint.ShapeRange GetShapesByPrefix(int slideNo, string shapePrefix)
        {
            PpOperations.SelectSlide(slideNo);
            return PpOperations.SelectShapesByPrefix(shapePrefix);
        }

        private void RightClick(Shape target)
        {
            Point pt = new Point(
                PpOperations.PointsToScreenPixelsX(target.Left + target.Width / 2),
                PpOperations.PointsToScreenPixelsY(target.Top + target.Height / 2));
            MouseUtil.SendMouseRightClick(pt.X, pt.Y);
        }
    }
}
