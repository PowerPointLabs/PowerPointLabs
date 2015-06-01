using System;
using System.Collections.Generic;
using FunctionalTestInterface;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

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

        public Slide GetCurrentSlide()
        {
            return PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide();
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
            var result = new List<Shape>();
            var nameList = new List<String>();
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (Shape sh in shapes)
            {
                if (sh.Name == shapeName)
                {
                    sh.Select();
                    nameList.Add(sh.Name);
                }
            }
            var range = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range(nameList.ToArray());
            return range;
        }

        public ShapeRange SelectShapesByPrefix(string prefix)
        {
            var result = new List<Shape>();
            var nameList = new List<String>();
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (Shape sh in shapes)
            {
                if (sh.Name.StartsWith(prefix))
                {
                    nameList.Add(sh.Name);
                }
            }
            var range = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Range(nameList.ToArray());
            return range;
        }
    }
}
