using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.Utils;

namespace Test.UnitTest
{
    /// <summary>
    /// Checks if copied Placeholders replacements support the same formats
    /// Also checks that unsupported placeholders cannot be copied
    /// </summary>
    [TestClass]
    public class PlaceHolderCopyTest : BaseUnitTest
    {
        private const int HorizontalPlaceholdersSlide = 2;
        private const int CenterPlaceholdersSlide = 3;
        private const int VerticalPlaceholdersSlide = 4;
        
        private const string HorizontalTitle = "Title 1";
        private const string HorizontalBody = "Content Placeholder 2";
        private const string CenterTitle= "Title 1";
        private const string Subtitle= "Subtitle 2";
        private const string VerticalTitle = "Title 1";
        private const string VerticalBody = "Content Placeholder 2";
        
        // not yet supported 
        private const int TablePlaceholderSlide = 5;
        private const int ChartPlaceholderSlide = 6;
        private const int PicturePlaceholderSlide = 7;
        
        private const string Table = "Content Placeholder 3";
        private const string Chart = "Content Placeholder 5";
        private const string Picture = "Content Placeholder 4";

        private Shapes _templateShapes;
        
        protected override string GetTestingSlideName()
        {
            return "CopyPlaceHolder.pptx";
        }
        
        private Shape GetShape(int slideNumber, string shapeName)
        {
            PpOperations.SelectSlide(slideNumber);
            return PpOperations.SelectShape(shapeName)[1];
        }
        
        private Shapes GetShapesObject(int slideNumber)
        {
            return PpOperations.SelectSlide(slideNumber).Shapes;
        }

        [TestInitialize]
        public void TestInitialize()
        {
            _templateShapes = GetShapesObject(1);
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyTitles()
        {
            Shape horizontalTitle = GetShape(HorizontalPlaceholdersSlide, HorizontalTitle);
            Shape verticalTitle = GetShape(VerticalPlaceholdersSlide, VerticalTitle);
            Shape centerTitle = GetShape(CenterPlaceholdersSlide, CenterTitle);
            Shape[] titles = {horizontalTitle, verticalTitle, centerTitle};

            foreach (Shape title in titles)
            {
                EnsureFormatsAreRetainedAfterCopy(title);
            }

        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyBodies()
        {
            Shape horizontalBody = GetShape(HorizontalPlaceholdersSlide, HorizontalBody);
            Shape verticalBody = GetShape(VerticalPlaceholdersSlide, VerticalBody);
            Shape subtitle = GetShape(CenterPlaceholdersSlide, Subtitle);
            Shape[] bodies = {horizontalBody, verticalBody, subtitle};

            foreach (Shape body in bodies)
            {
                EnsureFormatsAreRetainedAfterCopy(body);
            }
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyTable()
        {
            Shape table = GetShape(TablePlaceholderSlide, Table);
            Shape copy = ShapeUtil.CopyMsoPlaceHolder(new Format[0], table, _templateShapes);
            Assert.AreEqual(copy, null);
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyPicture()
        {
            Shape picture = GetShape(PicturePlaceholderSlide, Picture);
            EnsureFormatsAreRetainedAfterCopy(picture);
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyChart()
        {
            Shape chart = GetShape(ChartPlaceholderSlide, Chart);
            Shape copy = ShapeUtil.CopyMsoPlaceHolder(new Format[0], chart, _templateShapes);
            Assert.AreEqual(copy, null);
        }
        
        private void EnsureFormatsAreRetainedAfterCopy(Shape placeHolder)
        {
            // ensure that each placeholder's copy supports the same or more formats
            Format[] formatsFromOriginal = ShapeUtil.GetCopyableFormats(placeHolder);
            
            Shape copy = ShapeUtil.CopyMsoPlaceHolder(formatsFromOriginal, placeHolder, _templateShapes);
            Format[] formatsFromCopy = ShapeUtil.GetCopyableFormats(copy);

            IEnumerable<Format> formatsInBoth = formatsFromCopy.Intersect(formatsFromOriginal);
            Assert.AreEqual(formatsInBoth.Count(), formatsFromOriginal.Count());
        }

    }
}
