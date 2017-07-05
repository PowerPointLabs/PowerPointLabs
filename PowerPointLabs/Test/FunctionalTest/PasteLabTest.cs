using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Drawing;
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
        }

        private void PasteToFillSlide(int originalSlideNo, int expSlideNo)
        {
            var shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PplFeatures.PasteToFillSlide();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteAtCursorPosition(int originalSlideNo, int expSlideNo)
        {
            var shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            RightClick(GetShapesByPrefix(originalSlideNo, ShapeToClick)[1]);
            // wait for awhile for click to register properly
            ThreadUtil.WaitFor(500);
            PplFeatures.PasteAtCursorPosition();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteAtOriginalPosition(int originalSlideNo, int expSlideNo)
        {
            var shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PplFeatures.PasteAtOriginalPosition();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void ReplaceWithClipboard(int originalSlideNo, int expSlideNo)
        {
            var shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PpOperations.SelectShapes(new List<string> { ShapeToReplace });
            PplFeatures.ReplaceWithClipboard();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void PasteIntoGroup(int originalSlideNo, int expSlideNo)
        {
            var shapes = GetShapesByPrefix(originalSlideNo, ShapeToCopyPrefix);
            shapes.Cut();

            PpOperations.SelectShape(GroupToPaste);
            PplFeatures.PasteIntoGroup();

            AssertIsSame(originalSlideNo, expSlideNo);
        }

        private void AssertIsSame(int actualSlideNo, int expectedSlideNo)
        {
            var actualSlide = PpOperations.SelectSlide(actualSlideNo);
            var expectedSlide = PpOperations.SelectSlide(expectedSlideNo);

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
            var pt = new Point(
                PpOperations.PointsToScreenPixelsX(target.Left + target.Width / 2),
                PpOperations.PointsToScreenPixelsY(target.Top + target.Height / 2));
            MouseUtil.SendMouseRightClick(pt.X, pt.Y);
        }
    }
}
