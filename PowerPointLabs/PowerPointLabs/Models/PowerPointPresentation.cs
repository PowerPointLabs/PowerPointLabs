using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    internal class PowerPointPresentation
    {
        # region Properties

        private List<PowerPointSlide> _slides;

        public string FullName
        {
            get
            {
                return Path + @"\" + Name;
            }
        }

        public string Name { get; set; }

        public bool Opened
        {
            get
            {
                foreach (Presentation presentation in Globals.ThisAddIn.Application.Presentations)
                {
                    if (presentation.FullName == FullName)
                    {
                        Presentation = presentation;
                        return true;
                    }
                }

                return false;
            }
        }

        public string Path { get; set; }

        public Presentation Presentation { get; set; }

        public bool Saved
        {
            get { return Presentation.Saved == MsoTriState.msoTrue; }
        }

        public List<PowerPointSlide> Slides
        {
            get
            {
                if (_slides == null)
                {
                    _slides = new List<PowerPointSlide>();

                    var interopSlides = Presentation.Slides;

                    foreach (Slide interopSlide in interopSlides)
                    {
                        var s = PowerPointSlide.FromSlideFactory(interopSlide);
                        _slides.Add(s);
                    }
                }

                return _slides;
            }
        }

        public List<PowerPointSlide> SelectedSlides
        {
            get
            {
                var interopSlides = Presentation.Application.ActiveWindow.Selection.SlideRange;
                var slides = new List<PowerPointSlide>();

                foreach (Slide interopSlide in interopSlides)
                {
                    PowerPointSlide s = PowerPointSlide.FromSlideFactory(interopSlide);
                    slides.Add(s);
                }

                return slides;
            }
        }

        public int SlideCount
        {
            get { return Presentation.Slides.Count; }
        }

        public float SlideWidth
        {
            get
            {
                var dimensions = Presentation.PageSetup;
                return dimensions.SlideWidth;
            }
        }

        public float SlideHeight
        {
            get
            {
                var dimensions = Presentation.PageSetup;
                return dimensions.SlideHeight;
            }
        }
        # endregion

        # region Constructors
        public PowerPointPresentation()
        {
            Presentation = null;
        }

        public PowerPointPresentation(string path, string name)
        {
            Path = path;
            Name = name;
        }

        public PowerPointPresentation(Presentation presentation)
        {
            Presentation = presentation;
        }
        # endregion

        # region API
        public PowerPointSlide AddSlide(PpSlideLayout layout = PpSlideLayout.ppLayoutBlank, string name = "")
        {
            if (!Opened)
            {
                return null;
            }
            
            var customLayout = Presentation.SlideMaster.CustomLayouts[layout];
            var newSlide = Presentation.Slides.AddSlide(SlideCount + 1, customLayout);

            if (name != "")
            {
                newSlide.Name = name;
            }

            var slideFromFactory = PowerPointSlide.FromSlideFactory(newSlide);

            Slides.Add(slideFromFactory);

            return slideFromFactory;
        }

        public void RemoveSlide(string name)
        {
            // TODO: to be implemented
        }

        public void RemoveSlide(int index)
        {
            Slides.RemoveAt(index);
        }

        public bool Create(bool withWidow, bool focus)
        {
            if (File.Exists(Name))
            {
                return false;
            }

            var workingWindow = Globals.ThisAddIn.Application.ActiveWindow;

            Presentation = Globals.ThisAddIn.Application.Presentations.Add(BoolToMsoTriState(withWidow));
            Presentation.SaveAs(FullName);

            if (!focus)
            {
                workingWindow.Activate();
            }

            return true;
        }

        public void Close()
        {
            Presentation.Close();
            Presentation = null;
        }

        public void Open(bool readOnly = false, bool untitled = false, bool withWindow = true, bool focus = true)
        {
            if (Opened)
            {
                return;
            }

            if (Create(withWindow, focus))
            {
                return;
            }

            var workingWindow = Globals.ThisAddIn.Application.ActiveWindow;

            Presentation = Globals.ThisAddIn.Application.Presentations.Open(Name, BoolToMsoTriState(readOnly),
                                                                            BoolToMsoTriState(untitled),
                                                                            BoolToMsoTriState(withWindow));

            if (!focus)
            {
                workingWindow.Activate();
            }
        }

        public void Save()
        {
            if (Presentation != null)
            {
                Presentation.Save();
            }
        }
        # endregion

        # region Helper Functions
        private MsoTriState BoolToMsoTriState(bool value)
        {
            return value ? MsoTriState.msoTrue : MsoTriState.msoFalse;
        }
        # endregion
    }
}
