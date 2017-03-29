﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    public class PowerPointPresentation
    {
#pragma warning disable 0618
        #region Properties
        private string _name;
        private const string extensionRegex = "\\.pptx?$";

        public static Application Application { get; set; }

        [Obsolete("DO NOT use this property! Instead, use Action Framework.")]
        public static PowerPointPresentation Current
        {
            get
            {
                return new PowerPointPresentation(Globals.ThisAddIn.Application.ActivePresentation);
            }
        }

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
                if (Regex.IsMatch(Name, extensionRegex, RegexOptions.IgnoreCase))
                {
                    return Path + @"\" + Name;
                }
                else
                {
                    return Path + @"\" + Name + ".pptx";
                }
            }
        }

        public bool HasEmptySection
        {
            get
            {
                var sectionProperties = SectionProperties;

                // Fix for rare case where the ack slide is the only slide in the section.
                // This should be counted as an empty section. so we temporarily remove the ack slide and add it back after.
                bool hasAckSlide = HasAckSlide();
                RemoveAckSlide();

                for (var i = 1; i <= sectionProperties.Count; i++)
                {
                    if (sectionProperties.SlidesCount(i) == 0)
                    {
                        if (hasAckSlide) AddAckSlide();
                        return true;
                    }
                }
                if (hasAckSlide) AddAckSlide();

                return false;
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                if (Regex.IsMatch(value, extensionRegex, RegexOptions.IgnoreCase))
                {
                    NameNoExtension = Regex.Replace(value, extensionRegex, "", RegexOptions.IgnoreCase);
                }
                else
                {
                    NameNoExtension = value;
                }

                _name = value + ".pptx";
            }
        }

        public string NameNoExtension { get; private set; }

        public bool Opened
        {
            get
            {
                if (Globals.ThisAddIn != null)
                {
                    return
                        Globals.ThisAddIn.Application.Presentations.Cast<Presentation>().Any(
                            presentation => presentation.Name == Name);
                }
                return Application.Presentations.Cast<Presentation>().Any(
                    presentation => presentation.Name == Name);
            }
        }

        public string Path { get; set; }

        public Presentation Presentation { get; set; }

        public bool Saved
        {
            get { return Presentation.Saved == MsoTriState.msoTrue; }
        }

        public SectionProperties SectionProperties
        {
            get { return Presentation.SectionProperties; }
        }

        public List<string> Sections
        {
            get
            {
                var sectionProperty = Presentation.SectionProperties;
                var sectionNames = new List<string>();

                for (var i = 1; i <= sectionProperty.Count; i++)
                {
                    sectionNames.Add(sectionProperty.Name(i));
                }

                return sectionNames;
            }
        }

        public PowerPointSlide FirstSlide
        {
            get
            {
                var slides = Presentation.Slides;
                if (slides.Count > 0)
                {
                    return PowerPointSlide.FromSlideFactory(slides[1]);
                }
                return null;
            }
        }

        /// <summary>
        /// 0-indexed.
        /// </summary>
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

        /// <summary>
        /// 1-indexed.
        /// </summary>
        public PowerPointSlide GetSlide(int index)
        {
            return PowerPointSlide.FromSlideFactory(Presentation.Slides[index]);
        }

        public List<PowerPointSlide> SelectedSlides
        {
            get
            {
                return PowerPointCurrentPresentationInfo.SelectedSlides.ToList();
            }
        }

        public int SlideCount
        {
            get { return Presentation.Slides.Count; }
        }

        public virtual float SlideWidth
        {
            get
            {
                var dimensions = Presentation.PageSetup;
                return dimensions.SlideWidth;
            }
            set
            {
                Presentation.PageSetup.SlideWidth = value;
            }
        }

        public virtual float SlideHeight
        {
            get
            {
                var dimensions = Presentation.PageSetup;
                return dimensions.SlideHeight;
            }
            set
            {
                Presentation.PageSetup.SlideHeight = value;
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
            Name = name + ".pptx";
        }

        public PowerPointPresentation(Presentation presentation)
        {
            Presentation = presentation;
            Name = presentation.Name;
        }
        # endregion

        # region API
        public void AddAckSlide()
        {
            if (!HasAckSlide())
            {
                var lastSlide = Slides.Last();
                lastSlide.CreateAckSlide();
            }
        }

        /// <summary>
        /// Go to slide
        /// </summary>
        /// <param name="index">1-based</param>
        public void GotoSlide(int index)
        {
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(index);
        }

        /// <summary>
        /// will stay in the current slide if exceed slide count
        /// </summary>
        public void GotoNextSlide()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide == null) return;

            var index = currentSlide.Index;
            if (index < Slides.Count)
            {
                GotoSlide(index + 1);
            }
        }

        public bool IsLastSlide()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide == null) return false;

            var index = currentSlide.Index;
            return index == Slides.Count;
        }

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

        public void RemoveAckSlide()
        {
            RemoveSlide(PowerPointAckSlide.IsAckSlide, true);
        }

        public bool HasAckSlide()
        {
            return Slides.Any(PowerPointAckSlide.IsAckSlide);
        }

        public void RemoveSlide(Func<Slide, bool> condition, bool deleteAll)
        {
            var slides = Presentation.Slides.Cast<Slide>().Where(condition).ToList();

            foreach (var slide in slides)
            {
                slide.Delete();

                if (!deleteAll)
                {
                    break;
                }
            }
        }

        public void RemoveSlide(Regex rule, bool deleteAll)
        {
            var slides = Presentation.Slides.Cast<Slide>().Where(slide => rule.IsMatch(slide.Name)).ToList();

            foreach (var slide in slides)
            {
                slide.Delete();

                if (!deleteAll)
                {
                    break;
                }
            }
        }

        public void RemoveSlide(string name, bool deleteAll)
        {
            RemoveSlide(new Regex("^" + name + "$"), deleteAll);
        }

        public void RemoveSlide(int index)
        {
            // here we need to change the 0-based index to 1-based index!!!
            Presentation.Slides[index + 1].Delete();
        }

        public bool Create(bool withWidow, bool focus)
        {
            if (File.Exists(FullName))
            {
                return false;
            }

            if (Globals.ThisAddIn != null)
            {
                Presentation = Globals.ThisAddIn.Application.Presentations.Add(BoolToMsoTriState(withWidow));
                Presentation.SaveAs(FullNameNoExtension);
            }
            else if (Application != null)
            {
                Presentation = Application.Presentations.Add(BoolToMsoTriState(withWidow));
            }

            if (!focus && Globals.ThisAddIn != null)
            {
                var workingWindow = Globals.ThisAddIn.Application.ActiveWindow;
                workingWindow.Activate();
            }

            return true;
        }

        public bool CreateInBackground()
        {
            if (File.Exists(FullName))
            {
                return false;
            }

            if (Globals.ThisAddIn != null)
            {
                Presentation = Globals.ThisAddIn.Application.Presentations.Add(BoolToMsoTriState(false));
                Presentation.SaveAs(FullNameNoExtension);
            }
            else if (Application != null)
            {
                Presentation = Application.Presentations.Add(BoolToMsoTriState(false));
            }

            return true;
        }

        public virtual void Close()
        {
            Presentation.Close();
            Presentation = null;

            Trace.TraceInformation("Presentation " + NameNoExtension + " is closed.");
        }

        public virtual bool Open(bool readOnly = false, bool untitled = false, bool withWindow = true, bool focus = true)
        {
            if (Opened)
            {
                return false;
            }

            // if the file doesn't exist, create and open the file then return
            if (Create(withWindow, focus))
            {
                return true;
            }

            var workingWindow = Globals.ThisAddIn.Application.ActiveWindow;

            try
            {
                Presentation = Globals.ThisAddIn.Application.Presentations.Open(FullName, BoolToMsoTriState(readOnly),
                                                                                BoolToMsoTriState(untitled),
                                                                                BoolToMsoTriState(withWindow));
            }
            catch (System.Exception)
            {
                return false;
            }

            if (!focus)
            {
                workingWindow.Activate();
            }

            return true;
        }

        public virtual bool OpenInBackground()
        {
            if (Opened)
            {
                return false;
            }

            // if the file doesn't exist, create and open the file then return
            if (CreateInBackground())
            {
                return true;
            }
            try
            {
                Presentation = Globals.ThisAddIn.Application.Presentations.Open(FullName, BoolToMsoTriState(false),
                                                                                BoolToMsoTriState(false),
                                                                                BoolToMsoTriState(false));
            }
            catch (Exception)
            {
                return false;
            }

            return true;
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
