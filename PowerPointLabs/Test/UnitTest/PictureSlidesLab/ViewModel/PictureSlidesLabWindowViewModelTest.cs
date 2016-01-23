using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Interface;
using PowerPointLabs.PictureSlidesLab.View.Interface;
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
            pslViewModel.ImageSelectionList.Add(new ImageItem
            {
                ImageFile = expectedString
            });
            pslViewModel.CleanUp();

            var pslViewModel2 = CreateViewModel();
            Assert.AreEqual(expectedString, 
                pslViewModel2.ImageSelectionList[0].ImageFile);
            pslViewModel2.ImageSelectionList.Clear();
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
    }
}
