using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    public class PowerPointCurrentPresentationInfo
    {
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

                if (CurrentSelection.Type == PpSelectionType.ppSelectionSlides)
                {
                    var interopSlides = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;

                    foreach (Slide interopSlide in interopSlides)
                    {
                        PowerPointSlide s = PowerPointSlide.FromSlideFactory(interopSlide);
                        slides.Add(s);
                    }
                }

                return slides;
            }
        }
    }
}
