using System;
using System.Collections.Generic;
using TestInterface;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointOperations : MarshalByRefObject, IPowerPointOperations
    {
        public void MaximizeWindow()
        {
            Globals.ThisAddIn.Application.ActiveWindow.WindowState = PpWindowState.ppWindowMaximized;
        }

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

        public List<ISlideData> FetchPresentationData(string pathToPresentation)
        {
            var presentation = Globals.ThisAddIn.Application.Presentations.Open(pathToPresentation,
                                                                                WithWindow: MsoTriState.msoFalse);
            var slideData = presentation.Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            presentation.Close();
            return slideData;
        }

        public List<ISlideData> FetchCurrentPresentationData()
        {
            var slideData = PowerPointPresentation.Current.Presentation
                .Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            return slideData;
        }

        public void SavePresentationAs(string presName)
        {
            Globals.ThisAddIn.Application.ActivePresentation.SaveCopyAs(presName);
        }

        public void ClosePresentation()
        {
            EnterFunctionalTest();
            Globals.ThisAddIn.Application.ActivePresentation.Close();
        }

        public void ClosePowerPointApplication()
        {
            EnterFunctionalTest();
            Globals.ThisAddIn.Application.Quit();
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

        public string GetNotesPageText(Slide slide)
        {
            if (slide == null || slide.HasNotesPage == MsoTriState.msoFalse)
            {
                return string.Empty;
            }

            var notesPagePlaceholders = slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
            var notesPageBody = notesPagePlaceholders
                .FirstOrDefault(shape => shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody);

            string notesText = notesPageBody != null ? notesPageBody.TextFrame.TextRange.Text : string.Empty;
            return notesText;
        }

        public void SetNotesPageText(Slide slide, string text)
        {
            if (slide == null)
            {
                return;
            }

            var notesPagePlaceholders = slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
            var notesPageBody = notesPagePlaceholders
                .FirstOrDefault(shape => shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody);

            if (notesPageBody != null)
            {
                notesPageBody.TextFrame.TextRange.Text = text;
            }
        }

        public Selection GetCurrentSelection()
        {
            return PowerPointCurrentPresentationInfo.CurrentSelection;
        }

        public ShapeRange SelectShape(string shapeName)
        {
            var nameList = new List<String>();
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (Shape sh in shapes)
            {
                if (sh.Name == shapeName)
                {
                    nameList.Add(sh.Name);
                    break;
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

        public Shape RecursiveGetShapeWithPrefix(params string[] prefixes)
        {
            var parentShape = SelectShapesByPrefix(prefixes[0])[1];
            for (int i = 1; i < prefixes.Length; ++i)
            {
                parentShape = parentShape.GroupItems.Cast<Shape>().FirstOrDefault(shape => shape.Name.StartsWith(prefixes[i]));
            }
            return parentShape;
        }

        public FileInfo ExportSelectedShapes()
        {
            var shapes = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange;
            var hashCode = DateTime.Now.GetHashCode();
            var pathName = PathUtil.GetTempTestFolder() + "shapeName" + hashCode;
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

        public void RenameSection(int index, string newName)
        {
            PowerPointPresentation.Current.SectionProperties.Rename(index, newName);
        }

        public void AddSection(int index, string sectionName)
        {
            PowerPointPresentation.Current.SectionProperties.AddSection(index, sectionName);
        }

        public void DeleteSection(int index, bool deleteSlides)
        {
            PowerPointPresentation.Current.SectionProperties.Delete(index, deleteSlides);
        }

        public void ShowAllSlideNumbers()
        {
            Slide[] slides = GetAllSlides();
            foreach (Slide s in slides)
            {
                s.HeadersFooters.SlideNumber.Visible = MsoTriState.msoTrue;
            }
        }
    }
}
