using System;
using System.Collections.Generic;
using System.IO;

using Microsoft.Office.Interop.PowerPoint;

namespace TestInterface
{
    public interface IPowerPointOperations
    {
        # region PowerPoint Application API

        void MaximizeWindow();
        void EnterFunctionalTest();
        void ExitFunctionalTest();
        bool IsInFunctionalTest();
        void MaximizeWindow(int windowNumber);
        void NewWindow();
        int GetNumWindows();
        HashSet<Type> GetOpenPaneTypes();
        List<ISlideData> FetchPresentationData(string pathToPresentation);
        List<ISlideData> FetchCurrentPresentationData();
        void SavePresentationAs(string presName);
        void ClosePresentation();
        void ActivatePresentation();
        int PointsToScreenPixelsX(float x);
        int PointsToScreenPixelsY(float y);
        Boolean IsOffice2010();
        Boolean IsOffice2013();

        # endregion

        # region Slide-related API

        Slide GetCurrentSlide();
        Slide SelectSlide(int index);
        Slide SelectSlide(string slideName);
        Slide[] GetAllSlides();
        string GetNotesPageText(Slide slide);
        void SetNotesPageText(Slide slide, string text);
        void ShowAllSlideNumbers();

        # endregion

        # region Shape-related API

        Selection GetCurrentSelection();
        ShapeRange SelectShape(string shapeName);
        ShapeRange SelectShapes(IEnumerable<string> shapeNames);
        ShapeRange SelectShapesByPrefix(string prefix);
        Shape RecursiveGetShapeWithPrefix(params string[] prefixes);
        FileInfo ExportSelectedShapes();
        string SelectAllTextInShape(string shapeName);
        string SelectTextInShape(string shapeName, int startIndex, int endIndex);

        # endregion

        # region Section-related API

        void RenameSection(int index, string newName);
        void AddSection(int index, string sectionName);
        void DeleteSection(int index, bool deleteSlides);

        # endregion
    }
}
