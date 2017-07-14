using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.Views.Interface;
using PowerPointLabs.PictureSlidesLab.ViewModel;

namespace Test.UnitTest.PictureSlidesLab.ViewModel
{
    [TestClass]
    public class PictureSlidesLabWindowViewModelTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void TestPersistence()
        {
            var expectedString = "Test Images Persistence";
            var pslViewModel = CreateViewModel();
            pslViewModel.ImageSelectionList.Clear();
            // the first image item should be placeholder for `Choose Pictures`
            // so add a dummy item here
            pslViewModel.ImageSelectionList.Add(CreateDummyImageItem());
            pslViewModel.ImageSelectionList.Add(new ImageItem
            {
                ImageFile = expectedString,
                FullSizeImageFile = "something"
            });
            pslViewModel.CleanUp();

            var pslViewModel2 = CreateViewModel();
            Assert.AreEqual(expectedString, 
                pslViewModel2.ImageSelectionList[1].ImageFile);
            pslViewModel2.ImageSelectionList.Clear();
            // create a dummy item in order to clean up
            pslViewModel2.ImageSelectionList.Add(CreateDummyImageItem());
            pslViewModel2.CleanUp();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestPersistenceWhenNoFullsizeImage()
        {
            var expectedString = "Test Images Persistence";
            var pslViewModel = CreateViewModel();
            pslViewModel.ImageSelectionList.Clear();
            // the first image item should be placeholder for `Choose Pictures`
            // so add a dummy item here
            pslViewModel.ImageSelectionList.Add(CreateDummyImageItem());
            pslViewModel.ImageSelectionList.Add(new ImageItem
            {
                ImageFile = expectedString
                // without full size image (null)
            });
            pslViewModel.CleanUp();

            var pslViewModel2 = CreateViewModel();
            Assert.AreEqual(1, pslViewModel2.ImageSelectionList.Count);
            pslViewModel2.ImageSelectionList.Clear();
            // create a dummy item in order to clean up
            pslViewModel2.ImageSelectionList.Add(CreateDummyImageItem());
            pslViewModel2.CleanUp();
        }

        private PictureSlidesLabWindowViewModel CreateViewModel()
        {
            var viewMock = new Mock<IPictureSlidesLabWindowView>();
            var stylesDesignerMock = new Mock<IStylesDesigner>();
            return new PictureSlidesLabWindowViewModel(
                viewMock.Object,
                stylesDesignerMock.Object);
        }

        private ImageItem CreateDummyImageItem()
        {
            return new ImageItem
            {
                ImageFile = "something",
                FullSizeImageFile = "something"
            };
        }
    }
}
