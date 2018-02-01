using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.SyncLab;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    /// <summary>
    /// Checks if copied Placeholders replacements support the same formats
    /// Also checks that unsupported placeholders cannot be copied
    /// </summary>
    [TestClass]
    public class PlaceHolderCopyTest : BaseSyncLabTest
    {
        private const int HorizontalPlaceholdersSlide = 41;
        private const int CenterPlaceholdersSlide = 42;
        private const int VerticalPlaceholdersSlide = 43;
        
        private const string HorizontalTitle = "Title 1";
        private const string HorizontalBody = "Content Placeholder 2";
        private const string CenterTitle= "Title 1";
        private const string Subtitle= "Subtitle 2";
        private const string VerticalTitle = "Title 1";
        private const string VerticalBody = "Content Placeholder 2";
        
        // not yet supported by synclab
        private const int TablePlaceholderSlide = 44;
        private const int ChartPlaceholderSlide = 45;
        private const int PicturePlaceholderSlide = 46;
        
        private const string Table = "Content Placeholder 3";
        private const string Chart = "Content Placeholder 5";
        private const string Picture = "Content Placeholder 4";
        
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
            Shape copy = SyncFormatUtil.CopyMsoPlaceHolder(new Format[0], table);
            Assert.Equals(copy, null);
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyPicture()
        {
            Shape picture = GetShape(PicturePlaceholderSlide, Picture);
            Shape copy = SyncFormatUtil.CopyMsoPlaceHolder(new Format[0], picture);
            Assert.Equals(copy, null);
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestCopyChart()
        {
            Shape chart = GetShape(ChartPlaceholderSlide, Chart);
            Shape copy = SyncFormatUtil.CopyMsoPlaceHolder(new Format[0], chart);
            Assert.Equals(copy, null);
        }
        
        private void EnsureFormatsAreRetainedAfterCopy(Shape placeHolder)
        {
            // ensure that each placeholder's copy supports the same or more formats
            Format[] formatsFromOriginal = GetCopyableFormats(placeHolder);
            List<Type> typesFromOriginal = formatsFromOriginal.Select(format => format.FormatType).ToList();
            
            Shape copy = SyncFormatUtil.CopyMsoPlaceHolder(formatsFromOriginal, placeHolder);
            Format[] formatsFromCopy = GetCopyableFormats(copy);
            IEnumerable<Type> typesFromCopy = formatsFromCopy.Select(format => format.FormatType);

            IEnumerable<Type> typesInBoth = typesFromCopy.Intersect(typesFromOriginal);
            Assert.Equals(typesInBoth.Count(), typesFromOriginal.Count);
        }

        private Format[] GetCopyableFormats(Shape shape)
        {
            return (from format in SyncFormatConstants.Formats 
                    where format.CanCopy(shape) select format)
                .ToArray();
        }
    }
}
