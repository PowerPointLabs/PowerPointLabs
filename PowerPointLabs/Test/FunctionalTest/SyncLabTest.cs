using System;
using System.Collections.Generic;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs;
using PowerPointLabs.TextCollection;

using TestInterface;

using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class SyncLabTest : BaseFunctionalTest
    {
        private const int MaxRetry = 5;
        private const int CategoryIndexPosition = 0;
        private const int FormatItemIndexPosition = 1;

        private const int OriginalSyncGroupToShapeSlideNo = 36;
        private const int ExpectedSyncGroupToShapeSlideNo = 37;
        private const int OriginalSyncShapeToGroupSlideNo = 38;
        private const int ExpectedSyncShapeToGroupSlideNo = 39;

        private const string Line = "Straight Connector 2";
        private const string RotatedArrow = "Right Arrow 5";
        private const string Group = "Group 1";
        private const string Oval = "Oval 4";
        private const string CopyFromShape = "CopyFrom";
        private const string UnrotatedRectangle = "Rectangle 3";

        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_SyncLabTest()
        {
            var syncLab = PplFeatures.SyncLab;
            syncLab.OpenPane();

            TestSync(syncLab);
            TestErrorDialogs(syncLab);
        }

        private void TestErrorDialogs(ISyncLabController syncLab)
        {
            PpOperations.SelectSlide(OriginalSyncGroupToShapeSlideNo);

            // no selection copy
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                "Please select one shape to copy.", syncLab.Copy, "Ok");

            // 2 item selected copy
            List<String> shapes = new List<string> { Line, RotatedArrow };
            PpOperations.SelectShapes(shapes);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                "Please select one shape to copy.", syncLab.Copy, "Ok");

            // group selected copy
            PpOperations.SelectShape(Group);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                "Please select one shape to copy.", syncLab.Copy, "Ok");

            // copy blank item for the paste error dialog test
            PpOperations.SelectShape(Line);    
            CopyStyle(syncLab);

            // no selection sync
            PpOperations.SelectSlide(ExpectedSyncShapeToGroupSlideNo);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                "Please select at least one item to apply this format to.", () => syncLab.Sync(0), "Ok");
        }

        private void TestSync(ISyncLabController syncLab)
        {
            Sync(syncLab, OriginalSyncGroupToShapeSlideNo, ExpectedSyncGroupToShapeSlideNo, CopyFromShape, RotatedArrow, 1, 0);
            Sync(syncLab, OriginalSyncShapeToGroupSlideNo, ExpectedSyncShapeToGroupSlideNo, Line, Oval, 1, 4);
        }

        private void Sync(ISyncLabController syncLab, int originalSlide, int expectedSlide,
                string fromShape, string toShape, int categoryPosition, int itemPosition)
        {
            PpOperations.SelectSlide(originalSlide);
            PpOperations.SelectShape(fromShape);

            CopyStyle(syncLab, categoryPosition, itemPosition);

            PpOperations.SelectShape(toShape);
            syncLab.Sync(0);

            IsSame(originalSlide, expectedSlide, toShape);
        }

        private void IsSame(int originalSlideNo, int expectedSlideNo, string shapeToCheck)
        {
            var actualSlide = PpOperations.SelectSlide(originalSlideNo);
            var actualShape = PpOperations.SelectShape(shapeToCheck)[1];
            var expectedSlide = PpOperations.SelectSlide(expectedSlideNo);
            var expectedShape = PpOperations.SelectShape(shapeToCheck)[1];
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameShape(expectedShape, actualShape);
        }

        private void CopyStyle(ISyncLabController syncLab)
        {
            new Task(() =>
            {
                ThreadUtil.WaitFor(1000);
                syncLab.DialogClickOk();
            }).Start();
            syncLab.Copy();
        }

        private void CopyStyle(ISyncLabController syncLab, int categoryPosition, int itemPosition)
        {
            int[,] dialogItems = new int[,] { { categoryPosition, itemPosition } };
            CopyStyle(syncLab, dialogItems);
        }

        private void CopyStyle(ISyncLabController syncLab, int[,] dialogItems)
        {
            new Task(() =>
            {
                ThreadUtil.WaitFor(1000);
                for (int i = 0; i < dialogItems.GetLength(0); i++)
                {
                    syncLab.DialogSelectItem(dialogItems[i, CategoryIndexPosition], dialogItems[i, FormatItemIndexPosition]);
                }
                syncLab.DialogClickOk();
            }).Start();
            syncLab.Copy();
        }
    }
}
