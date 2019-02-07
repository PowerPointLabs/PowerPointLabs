using System;
using System.Collections.Generic;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TextCollection;

using Test.Util;

using TestInterface;

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
        
        private const int HorizontalPlaceHolderSlideNo = 41;
        private const int CenterPlaceHolderSlideNo = 42;
        private const int VerticalPlaceHolderSlideNo = 43;
        
        private const int TablePlaceHolderSlideNo = 44;
        private const int ChartPlaceHolderSlideNo = 45;
        private const int PicturePlaceHolderSlideNo = 46;
        private const int SmartArtSlideNo = 47;
        
        private const int CustomPerspectiveShadowSlideNo = 48;

        private const string Line = "Straight Connector 2";
        private const string RotatedArrow = "Right Arrow 5";
        private const string Group = "Group 1";
        private const string Oval = "Oval 4";
        private const string CopyFromShape = "CopyFrom";
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string CustomPerspectiveShadow = "Custom Perspective Shadow";
        
        private const string HorizontalTitle = "Title 1";
        private const string HorizontalBody = "Content Placeholder 2";
        private const string CenterTitle= "Title 1";
        private const string Subtitle= "Subtitle 2";
        private const string VerticalTitle = "Title 1";
        private const string VerticalBody = "Content Placeholder 2";
        private const string Table = "Content Placeholder 3";
        private const string Chart = "Content Placeholder 5";
        private const string Picture = "Content Placeholder 4";
        private const string SmartArt = "SmartArt 1";

        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_SyncLabTest()
        {
            ISyncLabController syncLab = PplFeatures.SyncLab;
            syncLab.OpenPane();

            TestSync(syncLab);
            TestErrorDialogs(syncLab);
            TestCopySupportedPlaceHolders(syncLab);
            TestCopyUnsupportedPlaceHolders(syncLab);
            TestCopyCustomPerspectiveShadow(syncLab);
        }

        private void TestCopySupportedPlaceHolders(ISyncLabController syncLab)
        {
            // Ensure that one style can be copied from placeholder objects.
            // Testing copying of all valid placeholder styles is done in unit tests.
            PpOperations.SelectSlide(HorizontalPlaceHolderSlideNo);
            PpOperations.SelectShape(HorizontalBody);
            CopyStyle(syncLab, 1, 0);
            PpOperations.SelectShape(HorizontalTitle);
            CopyStyle(syncLab, 1, 0);
            
            PpOperations.SelectSlide(VerticalPlaceHolderSlideNo);
            PpOperations.SelectShape(VerticalBody);
            CopyStyle(syncLab, 1, 0);
            PpOperations.SelectShape(VerticalTitle);
            CopyStyle(syncLab, 1, 0);
            
            PpOperations.SelectSlide(CenterPlaceHolderSlideNo);
            PpOperations.SelectShape(Subtitle);
            CopyStyle(syncLab, 1, 0);
            PpOperations.SelectShape(CenterTitle);
            CopyStyle(syncLab, 1, 0);
            
            PpOperations.SelectSlide(PicturePlaceHolderSlideNo);
            PpOperations.SelectShape(Picture);
            CopyStyle(syncLab, 1, 0);
        }
        
        private void TestCopyUnsupportedPlaceHolders(ISyncLabController syncLab)
        {
            
            PpOperations.SelectSlide(ChartPlaceHolderSlideNo);
            PpOperations.SelectShape(Chart);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorCopySelectionInvalid, syncLab.Copy, "Ok");
            
            PpOperations.SelectSlide(TablePlaceHolderSlideNo);
            PpOperations.SelectShape(Table);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorCopySelectionInvalid, syncLab.Copy, "Ok");
        }
        
        private void TestCopyCustomPerspectiveShadow(ISyncLabController syncLab)
        {
            
            PpOperations.SelectSlide(CustomPerspectiveShadowSlideNo);
            PpOperations.SelectShape(CustomPerspectiveShadow);

            // shadow item
            int categoryIndex = 3;
            int itemIndex = 0;
            
            new Task(() =>
            {
                ThreadUtil.WaitFor(1000);
                syncLab.DialogSelectItem(categoryIndex, itemIndex);
                MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.WarningDialogTitle,
                    SyncLabText.WarningSyncPerspectiveShadow, syncLab.DialogClickOk, "Ok");
            }).Start();
            syncLab.Copy();
        }

        private void TestErrorDialogs(ISyncLabController syncLab)
        {
            PpOperations.SelectSlide(OriginalSyncGroupToShapeSlideNo);

            // no selection copy
            ExpectCopyButtonDisabled(syncLab);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorCopySelectionInvalid, syncLab.Copy, "Ok");

            // 2 item selected copy
            List<String> shapes = new List<string> { Line, RotatedArrow };
            PpOperations.SelectShapes(shapes);
            ExpectCopyButtonEnabled(syncLab);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorCopySelectionInvalid, syncLab.Copy, "Ok");

            // group selected copy
            PpOperations.SelectShape(Group);
            ExpectCopyButtonEnabled(syncLab);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorCopySelectionInvalid, syncLab.Copy, "Ok");

            // copy blank item for the paste error dialog test
            PpOperations.SelectShape(Line);
            ExpectCopyButtonEnabled(syncLab);
            CopyStyle(syncLab);

            // no selection sync
            PpOperations.SelectSlide(ExpectedSyncShapeToGroupSlideNo);
            ExpectCopyButtonDisabled(syncLab);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorPasteSelectionInvalid, () => syncLab.Sync(0), "Ok");
            
            // smart art
            PpOperations.SelectSlide(SmartArtSlideNo);
            PpOperations.SelectShape(SmartArt);
            ExpectCopyButtonEnabled(syncLab);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(SyncLabText.ErrorDialogTitle,
                SyncLabText.ErrorSmartArtUnsupported, syncLab.Copy, "Ok");
        }

        private void TestSync(ISyncLabController syncLab)
        {
            Sync(syncLab, OriginalSyncGroupToShapeSlideNo, ExpectedSyncGroupToShapeSlideNo, CopyFromShape, RotatedArrow, 1, 0);
            Sync(syncLab, OriginalSyncShapeToGroupSlideNo, ExpectedSyncShapeToGroupSlideNo, Line, Oval, 0, 4);
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
            Slide actualSlide = PpOperations.SelectSlide(originalSlideNo);
            Shape actualShape = PpOperations.SelectShape(shapeToCheck)[1];
            Slide expectedSlide = PpOperations.SelectSlide(expectedSlideNo);
            Shape expectedShape = PpOperations.SelectShape(shapeToCheck)[1];
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

        private void ExpectCopyButtonEnabled(ISyncLabController syncLab)
        {
            ThreadUtil.WaitFor(500);
            Assert.IsTrue(syncLab.GetCopyButtonEnabledStatus());
        }

        private void ExpectCopyButtonDisabled(ISyncLabController syncLab)
        {
            ThreadUtil.WaitFor(500);
            Assert.IsFalse(syncLab.GetCopyButtonEnabledStatus());
        }
    }
}
