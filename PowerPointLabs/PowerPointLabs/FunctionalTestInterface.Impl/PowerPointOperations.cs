using System;
using System.Collections.Generic;
using FunctionalTestInterface;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointOperations : MarshalByRefObject, IPowerPointOperations
    {
        public void EnterFunctionalTest()
        {
            PowerPointCurrentPresentationInfo.IsInFunctionalTest = true;
        }

        public void ExitFunctionalTest()
        {
            PowerPointCurrentPresentationInfo.IsInFunctionalTest = false;
        }

        public bool IsInFunctionalTest()
        {
            return PowerPointCurrentPresentationInfo.IsInFunctionalTest;
        }

        public void ClosePresentation()
        {
            EnterFunctionalTest();
            Globals.ThisAddIn.Application.ActivePresentation.Close();
        }

        public void ActivatePresentation()
        {
            MessageBox.Show(new Form() { TopMost = true },
                "###__DO_NOT_OPEN_OTHER_WINDOW__###\n" +
                "###___DURING_FUNCTIONAL_TEST___###", "PowerPointLabs FT");
        }

        public int PointsToScreenPixelsX(float x)
        {
            return Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(x);
        }

        public int PointsToScreenPixelsY(float y)
        {
            return Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(y);
        }

        public Boolean IsOffice2010()
        {
            return Globals.ThisAddIn.Application.Version == Globals.ThisAddIn.OfficeVersion2010;
        }

        public Boolean IsOffice2013()
        {
            return Globals.ThisAddIn.Application.Version == Globals.ThisAddIn.OfficeVersion2013;
        }

        public Slide GetCurrentSlide()
        {
            return PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
        }

        public Slide[] GetAllSlides()
        {
            return PowerPointPresentation.Current.Presentation.Slides.Cast<Slide>().ToArray();
        }

        public Slide SelectSlide(int index)
        {
            var slides = PowerPointPresentation.Current.Slides;
            for (int i = 0; i <= slides.Count; i++)
            {
                if (i == (index - 1))
                {
                    var slide = slides[i].GetNativeSlide();
                    slide.Select();
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(index);
                    return slide;
                }
            }
            return null;
        }

        public Slide SelectSlide(string slideName)
        {
            var slides = PowerPointPresentation.Current.Slides;
            for (int i = 0; i <= slides.Count; i++)
            {
                if (slideName == slides[i].Name)
                {
                    var slide = slides[i].GetNativeSlide();
                    slide.Select();
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(i + 1);
                    return slide;
                }
            }
            return null;
        }

        public Selection GetCurrentSelection()
        {
            return PowerPointCurrentPresentationInfo.CurrentSelection;
        }

        public ShapeRange SelectShapes(string shapeName)
        {
            var nameList = new List<String>();
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (Shape sh in shapes)
            {
                if (sh.Name == shapeName)
                {
                    nameList.Add(sh.Name);
                }
            }
            return SelectShapes(nameList);
        }

        public ShapeRange SelectShapes(IEnumerable<string> shapeNames)
        {
            var range = PowerPointCurrentPresentationInfo
                .CurrentSlide.Shapes.Range(shapeNames.ToArray());

            if (range.Count > 0)
            {
                range.Select();
                return range;
            }
            return null;
        }

        public ShapeRange SelectShapesByPrefix(string prefix)
        {
            var nameList = new List<String>();
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (Shape sh in shapes)
            {
                if (sh.Name.StartsWith(prefix))
                {
                    nameList.Add(sh.Name);
                }
            }
            return SelectShapes(nameList);
        }

        public FileInfo ExportSelectedShapes()
        {
            var shapes = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange;
            var hashCode = DateTime.Now.GetHashCode();
            var pathName = Path.GetTempPath() + "shapeName" + hashCode;
            shapes.Export(pathName, PpShapeFormat.ppShapeFormatPNG);
            return new FileInfo(pathName);
        }

        public string SelectAllTextInShape(string shapeName)
        {
            var shape = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes
                                                                      .Cast<Shape>()
                                                                      .FirstOrDefault(sh => sh.Name == shapeName);
            var textRange = shape.TextFrame2.TextRange;
            textRange.Select();
            return textRange.Text;
        }

        public string SelectTextInShape(string shapeName, int startIndex, int endIndex)
        {
            var shape = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes
                                                                      .Cast<Shape>()
                                                                      .FirstOrDefault(sh => sh.Name == shapeName);
            var textRange = shape.TextFrame2.TextRange.Characters[startIndex, endIndex - startIndex];
            textRange.Select();
            return textRange.Text;
        }
    }
}
