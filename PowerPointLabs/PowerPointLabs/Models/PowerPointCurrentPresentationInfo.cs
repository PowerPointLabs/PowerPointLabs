using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    [Obsolete("DO NOT use this class! Instead, use Action Framework.")]
    public class PowerPointCurrentPresentationInfo
    {
        public static PowerPointSlide CurrentSlide
        {
            get
            {
                try
                {
                    Slide interopSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide as Slide;
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
                return GetSelectionInActiveWindow();
            }
        }

        public static IEnumerable<PowerPointSlide> SelectedSlides
        {
            get
            {
                List<PowerPointSlide> slides = new List<PowerPointSlide>();

                try
                {
                    SlideRange interopSlides = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;

                    foreach (Slide interopSlide in interopSlides)
                    {
                        PowerPointSlide s = PowerPointSlide.FromSlideFactory(interopSlide);
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

        private static Selection GetSelectionInActiveWindow() 
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveWindow.Selection;
            }
            catch (COMException) 
            {
                return null;
            }
        }
    }
}
