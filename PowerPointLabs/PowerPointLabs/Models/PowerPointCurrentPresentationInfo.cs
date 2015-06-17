using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    public class PowerPointCurrentPresentationInfo
    {
        public static bool IsInFunctionalTest;

        public static PowerPointSlide CurrentSlide
        {
            get
            {
                try
                {
                    var interopSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide as Slide;
                    return PowerPointSlide.FromSlideFactory(interopSlide);
                }
                catch (COMException)
                {
                    // No slide is selected, or in view.
                    return null;
                }
            }
        }

        public static Selection CurrentSelection
        {
            get
            {
                return Globals.ThisAddIn.Application.ActiveWindow.Selection;
            }
        }

        public static IEnumerable<PowerPointSlide> SelectedSlides
        {
            get
            {
                var slides = new List<PowerPointSlide>();

                try
                {
                    var interopSlides = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;

                    foreach (Slide interopSlide in interopSlides)
                    {
                        var s = PowerPointSlide.FromSlideFactory(interopSlide);
                        slides.Add(s);
                    }
                }
                catch (COMException)
                {
                    return new List<PowerPointSlide>();
                }

                return slides;
            }
        }
    }
}
