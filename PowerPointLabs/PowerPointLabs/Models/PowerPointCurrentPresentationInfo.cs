using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    internal class PowerPointCurrentPresentationInfo
    {
        public static Presentation CurrentPresentation
        {
            get { return Globals.ThisAddIn.Application.ActivePresentation; }
        }

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

        public static string CurrentPresentationName
        {
            get { return Globals.ThisAddIn.Application.ActivePresentation.Name; }
        }

        public static Selection CurrentSelection
        {
            get
            {
                return Globals.ThisAddIn.Application.ActiveWindow.Selection;
            }
        }

        public static List<string> Sections
        {
            get
            {
                var sectionProperty = CurrentPresentation.SectionProperties;
                var sectionNames = new List<string>();

                for (var i = 1; i <= sectionProperty.Count; i++)
                {
                    sectionNames.Add(sectionProperty.Name(i));
                }

                return sectionNames;
            }
        }

        public static IEnumerable<PowerPointSlide> Slides
        {
            get
            {
                var interopSlides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
                List<PowerPointSlide> slides = new List<PowerPointSlide>();

                foreach (Slide interopSlide in interopSlides)
                {
                    PowerPointSlide s = PowerPointSlide.FromSlideFactory(interopSlide);
                    slides.Add(s);
                }

                return slides;
            }
        }

        public static IEnumerable<PowerPointSlide> SelectedSlides
        {
            get
            {
                var interopSlides = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange;
                List<PowerPointSlide> slides = new List<PowerPointSlide>();

                foreach (Slide interopSlide in interopSlides)
                {
                    PowerPointSlide s = PowerPointSlide.FromSlideFactory(interopSlide);
                    slides.Add(s);
                }

                return slides;
            }
        }

        public static bool SlidesHaveCaptions(IEnumerable<PowerPointSlide> slides)
        {
            foreach (PowerPointSlide slide in slides)
            {
                if (slide.HasCaptions())
                {
                    return true;
                }
            }
            return false;
        }

        public static int SlideCount
        {
            get { return Globals.ThisAddIn.Application.ActivePresentation.Slides.Count; }
        }

        public static float SlideWidth
        {
            get
            {
                var dimensions = Globals.ThisAddIn.Application.ActivePresentation.PageSetup;
                return dimensions.SlideWidth;
            }
        }

        public static float SlideHeight
        {
            get
            {
                var dimensions = Globals.ThisAddIn.Application.ActivePresentation.PageSetup;
                return dimensions.SlideHeight;
            }
        }
    }
}
