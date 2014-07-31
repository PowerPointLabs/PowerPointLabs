using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    internal class PowerPointPresentation
    {
        # region Properties

        private List<PowerPointSlide> _slides;
        private string _name;

        public string FullName
        {
            get
            {
                return Path + @"\" + Name;
            }
        }

        public string FullNameNoExtension
        {
            get
            {
                return Path + @"\" + NameNoExtension;
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                NameNoExtension = value;
                _name = value + ".pptx";
            }
        }

        public string NameNoExtension { get; private set; }

        public bool Opened
        {
            get
            {
                foreach (Presentation presentation in Globals.ThisAddIn.Application.Presentations)
                {
                    if (presentation.Name == Name)
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
                var slides = new List<PowerPointSlide>();

                var interopSlides = Presentation.Slides;

                foreach (Slide interopSlide in interopSlides)
                {
                    var s = PowerPointSlide.FromSlideFactory(interopSlide);
                    slides.Add(s);
                }

                return slides;
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
        public PowerPointSlide AddSlide(PpSlideLayout layout = PpSlideLayout.ppLayoutText, string name = "")
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

        // TODO: need to be verified
        public void RemoveSlide(string name)
        {
            foreach (var slide in Slides)
            {
                if (slide.Name == name)
                {
                    Slides.Remove(slide);
                    break;
                }
            }
        }

        public void RemoveSlide(int index)
        {
            Slides.RemoveAt(index);
        }

        public bool Create(bool withWidow, bool focus)
        {
            if (File.Exists(FullName))
            {
                return false;
            }

            var workingWindow = Globals.ThisAddIn.Application.ActiveWindow;

            Presentation = Globals.ThisAddIn.Application.Presentations.Add(BoolToMsoTriState(withWidow));
            Presentation.SaveAs(FullNameNoExtension);

            if (!focus)
            {
                workingWindow.Activate();
            }

            return true;
        }

        public virtual void Close()
        {
            Presentation.Close();
            Presentation = null;
        }

        public virtual void Open(bool readOnly = false, bool untitled = false, bool withWindow = true, bool focus = true)
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

            Presentation = Globals.ThisAddIn.Application.Presentations.Open(FullName, BoolToMsoTriState(readOnly),
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
