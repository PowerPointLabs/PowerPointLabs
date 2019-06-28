using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.FunctionalTestInterface.Impl;
using PowerPointLabs.Utils;

using TestInterface;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace Test.Util
{
    class UnitTestPpOperations : IPowerPointOperations
    {
        public IntPtr Window => new IntPtr(App.HWND);

        public Presentation Pres { get; set; }

        public Application App { get; set; }

        private Slide _currentSlide;

        private Slide CurrentSlide
        {
            get { return _currentSlide ?? SelectSlide(1); }
            set { _currentSlide = value; }
        }

        private ShapeRange _currentShape;

        public UnitTestPpOperations(Presentation pres, Application app)
        {
            Pres = pres;
            App = app;
            PPLClipboard.Init(Window, true);
        }

        ~UnitTestPpOperations()
        {
            PPLClipboard.Instance.Teardown();
        }

        public void MaximizeWindow()
        {
            throw new NotImplementedException();
        }

        public void EnterFunctionalTest()
        {
            throw new NotImplementedException();
        }

        public void ExitFunctionalTest()
        {
            throw new NotImplementedException();
        }

        public bool IsInFunctionalTest()
        {
            throw new NotImplementedException();
        }

        public void MaximizeWindow(int windowNumber)
        {
            throw new NotImplementedException();
        }

        public void NewWindow()
        {
            throw new NotImplementedException();
        }

        public int GetNumWindows()
        {
            throw new NotImplementedException();
        }

        public void SetTagToAssociatedWindow()
        {
            throw new NotImplementedException();
        }

        public HashSet<Type> GetOpenPaneTypes()
        {
            throw new NotImplementedException();
        }

        public List<ISlideData> FetchPresentationData(string pathToPresentation)
        {
            Presentation presentation = App.Presentations.Open(pathToPresentation, WithWindow: MsoTriState.msoFalse);
            List<ISlideData> slideData = presentation.Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            presentation.Close();
            return slideData;
        }

        public List<ISlideData> FetchCurrentPresentationData()
        {
            List<ISlideData> slideData = Pres.Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            return slideData;
        }

        public void SavePresentationAs(string presName)
        {
            Pres.SaveCopyAs(presName);
        }

        public void ClosePresentation()
        {
            Pres.Close();
        }

        public void ActivatePresentation()
        {
            throw new NotImplementedException();
        }

        public int PointsToScreenPixelsX(float x)
        {
            throw new NotImplementedException();
        }

        public int PointsToScreenPixelsY(float y)
        {
            throw new NotImplementedException();
        }

        public bool IsOffice2010()
        {
            throw new NotImplementedException();
        }

        public bool IsOffice2013()
        {
            throw new NotImplementedException();
        }

        public Slide GetCurrentSlide()
        {
            return CurrentSlide;
        }

        public Slide SelectSlide(int index)
        {
            Slide slide = Pres.Slides[index];
            CurrentSlide = slide;
            return slide;
        }

        public Slide SelectSlide(string slideName)
        {
            foreach (Slide slide in Pres.Slides)
            {
                if (slide.Name == slideName)
                {
                    CurrentSlide = slide;
                    return slide;
                }
            }
            return null;
        }

        public Slide[] GetAllSlides()
        {
            return Pres.Slides.Cast<Slide>().ToArray();
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
            throw new NotImplementedException();
        }

        public ShapeRange SelectShape(string shapeName)
        {
            if (CurrentSlide == null) return null;

            List<string> nameList = new List<string>();
            nameList.Add(shapeName);
            _currentShape = CurrentSlide.Shapes.Range(nameList.ToArray());
            return _currentShape;
        }

        public ShapeRange SelectShapes(IEnumerable<string> shapeNames)
        {
            if (CurrentSlide == null) return null;
            _currentShape = CurrentSlide.Shapes.Range(shapeNames.ToArray());
            return _currentShape;
        }

        public ShapeRange SelectShapesByPrefix(string prefix)
        {
            if (CurrentSlide == null) return null;

            List<string> nameList = new List<string>();
            foreach (Shape shape in CurrentSlide.Shapes)
            {
                if (shape.Name.StartsWith(prefix))
                {
                    nameList.Add(shape.Name);
                }
            }
            _currentShape = CurrentSlide.Shapes.Range(nameList.ToArray());
            return _currentShape;
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
            ShapeRange shapes = _currentShape;
            int hashCode = DateTime.Now.GetHashCode();
            string pathName = TempPath.GetTempTestFolder() + "shapeName" + hashCode;
            shapes.Export(pathName, PpShapeFormat.ppShapeFormatPNG);
            return new FileInfo(pathName);
        }

        public string SelectAllTextInShape(string shapeName)
        {
            Shape shape = CurrentSlide.Shapes.Cast<Shape>().FirstOrDefault(sh => sh.Name == shapeName);
            TextRange2 textRange = shape.TextFrame2.TextRange;
            textRange.Select();
            return textRange.Text;
        }

        public string SelectTextInShape(string shapeName, int startIndex, int endIndex)
        {
            Shape shape = CurrentSlide.Shapes.Cast<Shape>().FirstOrDefault(sh => sh.Name == shapeName);
            TextRange2 textRange = shape.TextFrame2.TextRange.Characters[startIndex, endIndex - startIndex];
            textRange.Select();
            return textRange.Text;
        }

        public void RenameSection(int index, string newName)
        {
            Pres.SectionProperties.Rename(index, newName);
        }

        public void AddSection(int index, string sectionName)
        {
            Pres.SectionProperties.AddSection(index, sectionName);
        }

        public void DeleteSection(int index, bool deleteSlides)
        {
            Pres.SectionProperties.Delete(index, deleteSlides);
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
