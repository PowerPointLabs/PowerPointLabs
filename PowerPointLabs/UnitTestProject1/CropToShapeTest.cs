using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace UnitTestProject1
{
    [TestClass]
    public class CropToShapeTest : BaseUnitTest
    {
        [TestMethod]
        public void CropSelectionNoneTest()
        {
            var mockSel = new Mock<PowerPoint.Selection>();
            mockSel.Setup(selection => selection.Type).Returns(PowerPoint.PpSelectionType.ppSelectionNone);
            System.Windows.Forms.Fakes.ShimMessageBox.ShowStringString = (s, s1) => DialogResult.OK;

            var result = PowerPointLabs.CropToShape.Crop(mockSel.Object);
            Assert.AreEqual(null, result);
        }

        [TestMethod]
        [ExpectedException(typeof(Exception))]
        public void CropSelectionNoneWithNoHandleError()
        {
            var mockSel = new Mock<PowerPoint.Selection>();
            mockSel.Setup(selection => selection.Type).Returns(PowerPoint.PpSelectionType.ppSelectionNone);
            System.Windows.Forms.Fakes.ShimMessageBox.ShowStringString = (s, s1) => DialogResult.OK;

            PowerPointLabs.CropToShape.Crop(mockSel.Object, handleError: false);
        }

        [TestMethod]
        public void CropOneShapeSuccessful()
        {
            ///
            /// Problems in this test
            /// 1. Too many mocks fell like a whitebox testing: selection, shapeRange, shape, fill, line, pptpresentation, slide, shapes..
            /// there should be something, with one line of code then can create them all by default.
            /// -- PowerPoint Object Model Factory with implemented objects..
            /// e.g. 
            /// var objModel = PpObjectModelFactory.withShape(1).selectShape(0).create();
            /// var croppedResult = CropToShape.crop(objModel.selection);
            /// Assert....
            /// 
            /// 2. Problematic global static methods:
            /// MessageBox, PowerPointPresentation, PowerPointCurrentPresentationInfo, etc
            /// should depend on interface/abstract instead of them
            /// 
            /// 3. Index starts from one.. need to handle this case in obj model.

            System.Windows.Forms.Fakes.ShimMessageBox.ShowStringString = (s, s1) => DialogResult.OK;
            var mockSel = new Mock<PowerPoint.Selection>();
            mockSel.Setup(sel => sel.Type).Returns(PowerPoint.PpSelectionType.ppSelectionShapes);
            var mockShapeRange = new Mock<PowerPoint.ShapeRange>();
            var mockShapeRangeAsGenericIEnumerable = mockShapeRange.As<IEnumerable<PowerPoint.Shape>>();
            var mockShapeRangeAsIEnumerable = mockShapeRange.As<IEnumerable>();

            mockSel.Setup(sel => sel.ShapeRange).Returns(mockShapeRange.Object);

            var shapeList = new List<PowerPoint.Shape>();
            // add a dummy shape
            var dummyShape = new Mock<PowerPoint.Shape>();
            dummyShape.Setup(dummySh => dummySh.Type).Returns(Office.MsoShapeType.msoAutoShape);
            dummyShape.SetupGet(sh => sh.Name).Returns("name-of-shape-dummy");
            dummyShape.SetupGet(sh => sh.Left).Returns(123);
            dummyShape.SetupGet(sh => sh.Top).Returns(123);
            dummyShape.SetupSet(sh => sh.Name += It.IsAny<String>()).Verifiable();
            dummyShape.SetupSet(sh => sh.Left += It.IsAny<float>()).Verifiable();
            dummyShape.SetupSet(sh => sh.Top += It.IsAny<float>()).Verifiable();
            dummyShape.SetupSet(sh => sh.Visible = It.IsAny<Office.MsoTriState>()).Verifiable();
            shapeList.Add(dummyShape.Object);
            var mockShape = new Mock<PowerPoint.Shape>();
            var mockFill = new Mock<PowerPoint.FillFormat>();
            mockFill.Setup(fill => fill.UserPicture(It.IsAny<String>())).Verifiable();
            var mockLine = new Mock<PowerPoint.LineFormat>();
            mockLine.SetupSet(line => line.Visible = It.IsAny<Office.MsoTriState>());
            mockShape.Setup(sh => sh.Type).Returns(Office.MsoShapeType.msoAutoShape);
            mockShape.Setup(sh => sh.Rotation).Returns(0);
            mockShape.Setup(sh => sh.Line).Returns(mockLine.Object);
            mockShape.Setup(sh => sh.Fill).Returns(mockFill.Object);
            mockShape.SetupGet(sh => sh.Name).Returns("name-of-shape");
            mockShape.SetupGet(sh => sh.Left).Returns(123);
            mockShape.SetupGet(sh => sh.Top).Returns(123);
            mockShape.SetupSet(sh => sh.Name += It.IsAny<String>()).Verifiable();
            mockShape.SetupSet(sh => sh.Left += It.IsAny<float>()).Verifiable();
            mockShape.SetupSet(sh => sh.Top += It.IsAny<float>()).Verifiable();
            mockShape.SetupSet(sh => sh.Visible = It.IsAny<Office.MsoTriState>()).Verifiable();
            shapeList.Add(mockShape.Object);
            mockShapeRange.Setup(sr => sr.Count).Returns(1);
            mockShapeRangeAsGenericIEnumerable.Setup(sr => sr.GetEnumerator()).Returns(shapeList.GetEnumerator());
            mockShapeRangeAsIEnumerable.Setup(sr => sr.GetEnumerator()).Returns(shapeList.GetEnumerator());
            mockShapeRange.Setup(sr => sr.Cut()).Verifiable();
            mockShapeRange.Setup(sr => sr.Copy()).Verifiable();
            mockShapeRange.Setup(sr => sr.Delete()).Verifiable();
            mockShapeRange.Setup(sr => sr[It.IsAny<Int32>()]).Returns(mockShape.Object);

            var present = new Mock<PowerPointPresentation>();
            present.SetupGet(ps => ps.SlideWidth).Returns(1);
            present.SetupGet(ps => ps.SlideHeight).Returns(1);
            PowerPointLabs.Models.Fakes.ShimPowerPointPresentation.CurrentGet = () => present.Object;

            var mockCurSlide = new Mock<PowerPoint.Slide>();
            var mockShapes = new Mock<PowerPoint.Shapes>();
            mockCurSlide.Setup(cs => cs.Shapes).Returns(mockShapes.Object);
            mockCurSlide.Setup(cs => cs.Export(It.IsAny<String>(), It.IsAny<String>(), It.IsAny<int>(), It.IsAny<int>())).Verifiable();
            mockShapes.Setup(ss => ss.Paste()).Returns(mockShapeRange.Object);
            mockShapes.Setup(ss => ss.Count).Returns(1);
            mockShapes.Setup(ss => ss.Range(It.IsAny<string[]>())).Returns(mockShapeRange.Object);

            PowerPointLabs.Models.Fakes.ShimPowerPointCurrentPresentationInfo.CurrentSlideGet 
                = () => new PowerPointSlide(mockCurSlide.Object);
            PowerPointLabs.Fakes.ShimCropToShape.CreateFillInBackgroundForShapeShapeDouble = (shape, d) => { };

            var result = PowerPointLabs.CropToShape.Crop(mockSel.Object);
            Assert.AreSame(mockShape.Object, result);
            mockFill.VerifyAll();
            mockLine.VerifyAll();
            mockShapeRange.VerifyAll();
        }
    }
}
