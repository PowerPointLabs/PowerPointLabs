using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using TestInterface;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointOperations : MarshalByRefObject, IPowerPointOperations
    {
        public void MaximizeWindow()
        {
            FunctionalTestExtensions.GetCurrentWindow().WindowState = PpWindowState.ppWindowMaximized;
        }

        public void EnterFunctionalTest()
        {
            PowerPointLabsFT.IsFunctionalTestOn = true;
        }

        public void ExitFunctionalTest()
        {
            PowerPointLabsFT.IsFunctionalTestOn = false;
        }

        public bool IsInFunctionalTest()
        {
            return PowerPointLabsFT.IsFunctionalTestOn;
        }

        public void MaximizeWindow(int windowNumber)
        {
            FunctionalTestExtensions.GetApplication().Windows[windowNumber].Activate();
            FunctionalTestExtensions.GetApplication().Windows[windowNumber].WindowState = PpWindowState.ppWindowMaximized;
        }

        public void NewWindow()
        {
            Presentation presentation = FunctionalTestExtensions.GetPresentations().Add();
        }

        public int GetNumWindows()
        {
            return FunctionalTestExtensions.GetApplication().Windows.Count;
        }

        public void SetTagToAssociatedWindow()
        {
            DocumentWindow docWindow = FunctionalTestExtensions.GetCurrentWindow();
            foreach (Microsoft.Office.Tools.CustomTaskPane pane in FunctionalTestExtensions.GetAddIn().CustomTaskPanes)
            {
                if (pane.Control.Tag == null) { pane.Control.Tag = docWindow.HWND; }
            }
        }

        public HashSet<Type> GetOpenPaneTypes()
        {
            DocumentWindow docWindow = FunctionalTestExtensions.GetCurrentWindow(); //.GetApplication().Windows[1];
            HashSet<Type> result = new HashSet<Type>();
            foreach (Microsoft.Office.Tools.CustomTaskPane pane in FunctionalTestExtensions.GetAddIn().CustomTaskPanes)
            {
                if (pane.Control.Tag is int && (int)pane.Control.Tag == docWindow.HWND) { result.Add(pane.Control.GetType()); }
            }
            return result;
        }

        public List<ISlideData> FetchPresentationData(string pathToPresentation)
        {
            Presentation presentation = FunctionalTestExtensions.GetPresentations().Open(pathToPresentation,
                                                                                WithWindow: MsoTriState.msoFalse);
            List<ISlideData> slideData = presentation.Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            presentation.Close();
            return slideData;
        }

        public List<ISlideData> FetchCurrentPresentationData()
        {
            List<ISlideData> slideData = FunctionalTestExtensions.GetCurrentPresentation().Presentation
                .Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            return slideData;
        }

        public void SavePresentationAs(string presName)
        {
            FunctionalTestExtensions.GetCurrentPresentation().Presentation.SaveCopyAs(presName);
        }

        public void ClosePresentation()
        {
            EnterFunctionalTest();
            FunctionalTestExtensions.GetCurrentPresentation().Presentation.Close();
        }

        public void ClosePowerPointApplication()
        {
            EnterFunctionalTest();
            FunctionalTestExtensions.GetApplication().Quit();
        }

        public void ActivatePresentation()
        {
            MessageBoxUtil.Show("###__DO_NOT_OPEN_OTHER_WINDOW__###\n" + "###___DURING_FUNCTIONAL_TEST___###", "PowerPointLabs FT");
        }

        public int PointsToScreenPixelsX(float x)
        {
            return FunctionalTestExtensions.GetCurrentWindow().PointsToScreenPixelsX(x);
        }

        public int PointsToScreenPixelsY(float y)
        {
            return FunctionalTestExtensions.GetCurrentWindow().PointsToScreenPixelsY(y);
        }

        public bool IsOffice2010()
        {
            return FunctionalTestExtensions.GetApplication().Version == "14.0";
        }

        public bool IsOffice2013()
        {
            return FunctionalTestExtensions.GetApplication().Version == "15.0";
        }

        public Slide GetCurrentSlide()
        {
            return FunctionalTestExtensions.GetCurrentSlide().GetNativeSlide();
        }

        public Slide[] GetAllSlides()
        {
            return FunctionalTestExtensions.GetCurrentPresentation().Presentation.Slides.Cast<Slide>().ToArray();
        }

        public Slide SelectSlide(int index)
        {
            List<Models.PowerPointSlide> slides = FunctionalTestExtensions.GetCurrentPresentation().Slides;
            for (int i = 0; i <= slides.Count; i++)
            {
                if (i == (index - 1))
                {
                    Slide slide = slides[i].GetNativeSlide();
                    slide.Select();
                    FunctionalTestExtensions.GetCurrentWindow().View.GotoSlide(index);
                    return slide;
                }
            }
            return null;
        }

        public Slide SelectSlide(string slideName)
        {
            List<Models.PowerPointSlide> slides = FunctionalTestExtensions.GetCurrentPresentation().Slides;
            for (int i = 0; i <= slides.Count; i++)
            {
                if (slideName == slides[i].Name)
                {
                    Slide slide = slides[i].GetNativeSlide();
                    slide.Select();
                    FunctionalTestExtensions.GetCurrentWindow().View.GotoSlide(i + 1);
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

            IEnumerable<Shape> notesPagePlaceholders = slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
            Shape notesPageBody = notesPagePlaceholders
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

            IEnumerable<Shape> notesPagePlaceholders = slide.NotesPage.Shapes.Placeholders.Cast<Shape>();
            Shape notesPageBody = notesPagePlaceholders
                .FirstOrDefault(shape => shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderBody);

            if (notesPageBody != null)
            {
                notesPageBody.TextFrame.TextRange.Text = text;
            }
        }

        public Selection GetCurrentSelection()
        {
            return FunctionalTestExtensions.GetCurrentSelection();
        }

        public ShapeRange SelectShape(string shapeName)
        {
            List<string> nameList = new List<String>();
            nameList.Add(shapeName);
            return SelectShapes(nameList);
        }

        public ShapeRange SelectShapes(IEnumerable<string> shapeNames)
        {
            ShapeRange range = FunctionalTestExtensions.GetCurrentSlide().Shapes.Range(shapeNames.ToArray());

            if (range.Count > 0)
            {
                range.Select();
                return range;
            }
            return null;
        }

        public ShapeRange SelectShapesByPrefix(string prefix)
        {
            List<string> nameList = new List<String>();
            Microsoft.Office.Interop.PowerPoint.Shapes shapes = FunctionalTestExtensions.GetCurrentSlide().Shapes;
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
            Shape parentShape = SelectShapesByPrefix(prefixes[0])[1];
            for (int i = 1; i < prefixes.Length; ++i)
            {
                parentShape = parentShape.GroupItems.Cast<Shape>().FirstOrDefault(shape => shape.Name.StartsWith(prefixes[i]));
            }
            return parentShape;
        }

        public FileInfo ExportSelectedShapes()
        {
            ShapeRange shapes = FunctionalTestExtensions.GetCurrentSelection().ShapeRange;
            int hashCode = DateTime.Now.GetHashCode();
            string pathName = TempPath.GetTempTestFolder() + "shapeName" + hashCode;
            shapes.Export(pathName, PpShapeFormat.ppShapeFormatPNG);
            return new FileInfo(pathName);
        }

        public string SelectAllTextInShape(string shapeName)
        {
            Shape shape = FunctionalTestExtensions.GetCurrentSlide().Shapes
                                                                  .Cast<Shape>()
                                                                  .FirstOrDefault(sh => sh.Name == shapeName);
            TextRange2 textRange = shape.TextFrame2.TextRange;
            textRange.Select();
            return textRange.Text;
        }

        public string SelectTextInShape(string shapeName, int startIndex, int endIndex)
        {
            Shape shape = FunctionalTestExtensions.GetCurrentSlide().Shapes
                                                                      .Cast<Shape>()
                                                                      .FirstOrDefault(sh => sh.Name == shapeName);
            TextRange2 textRange = shape.TextFrame2.TextRange.Characters[startIndex, endIndex - startIndex];
            textRange.Select();
            return textRange.Text;
        }

        public void RenameSection(int index, string newName)
        {
            FunctionalTestExtensions.GetCurrentPresentation().SectionProperties.Rename(index, newName);
        }

        public void AddSection(int index, string sectionName)
        {
            FunctionalTestExtensions.GetCurrentPresentation().SectionProperties.AddSection(index, sectionName);
        }

        public void DeleteSection(int index, bool deleteSlides)
        {
            FunctionalTestExtensions.GetCurrentPresentation().SectionProperties.Delete(index, deleteSlides);
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
