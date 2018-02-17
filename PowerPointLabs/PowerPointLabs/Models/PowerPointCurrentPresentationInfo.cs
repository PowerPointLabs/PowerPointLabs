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
                Selection select = GetSelectionInActiveWindow();
                // Could be that there is no active window, we can try to activate one
                if (select == null && ActivateWindow())
                {
                    select = GetSelectionInActiveWindow();
                }
                return select;
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

        private static bool ActivateWindow() 
        {
            try 
            {
                Globals.ThisAddIn.Application.Activate();
                if (Globals.ThisAddIn.Application.Active == Microsoft.Office.Core.MsoTriState.msoTrue && Globals.ThisAddIn.Application.Windows.Count >= 1)
                {
                    Globals.ThisAddIn.Application.Windows[1].Activate();
                    return Globals.ThisAddIn.Application.Windows[1].Active == Microsoft.Office.Core.MsoTriState.msoTrue;
                }
                return false;
            } 
            catch (COMException) 
            {
                return false;
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
