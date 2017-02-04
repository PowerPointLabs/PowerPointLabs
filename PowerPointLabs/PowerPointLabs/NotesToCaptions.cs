﻿using System;
using System.Collections.Generic;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{
    class NotesToCaptions
    {
#pragma warning disable 0618
        public static void EmbedCaptionsOnSelectedSlides()
        {
            foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides) 
            {
                RemoveCaptionsFromSlide(slide);
                EmbedCaptionsOnSlide(slide);
            }
        }

        private static void EmbedCaptionsOnSlide(PowerPointSlide s)
        {
            String rawNotes = s.NotesPageText;

            if (String.IsNullOrWhiteSpace(rawNotes))
            {
                return;
            }

            var separatedNotes = SplitNotesByClicks(rawNotes);
            var captionCollection = ConvertSectionsToCaptions(separatedNotes);
            if (captionCollection.Count == 0)
            {
                return;
            }

            Shape previous = null;
            for (int i = 0; i < captionCollection.Count; i++)
            {
                String currentCaption = captionCollection[i];
                Shape captionBox = AddCaptionBoxToSlide(currentCaption, s);
                captionBox.Name = "PowerPointLabs Caption " + i;

                if (i == 0)
                {
                    s.SetShapeAsAutoplay(captionBox);
                }

                if (i != 0)
                {
                    s.ShowShapeAfterClick(captionBox, i);
                    s.HideShapeAfterClick(previous, i);
                }

                if (i == captionCollection.Count - 1)
                {
                    s.HideShapeAsLastClickIfNeeded(captionBox);
                }
                previous = captionBox;
            }
        }

        private static IEnumerable<string> SplitNotesByClicks(string rawNotes)
        {
            TaggedText taggedNotes = new TaggedText(rawNotes);
            List<String> splitByClicks = taggedNotes.SplitByClicks();
            return splitByClicks;
        }

        private static List<string> ConvertSectionsToCaptions(IEnumerable<string> separatedNotes)
        {
            List<String> captionCollection = new List<string>();
            foreach (string text in separatedNotes)
            {
                TaggedText section = new TaggedText(text);
                String currentCaption = section.ToPrettyString().Trim();
                if (!string.IsNullOrEmpty(currentCaption))
                {
                    captionCollection.Add(currentCaption);
                }
            }
            return captionCollection;
        }

        private static Shape AddCaptionBoxToSlide(string caption, PowerPointSlide s)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;
            
            Shape textBox = s.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, slideHeight - 100,
                slideWidth, 100);
            textBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            textBox.TextFrame.TextRange.Text = caption;
            textBox.TextFrame.WordWrap = MsoTriState.msoTrue;
            textBox.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            textBox.TextFrame.TextRange.Font.Size = 12;
            textBox.Fill.BackColor.RGB = 0;
            textBox.Fill.Transparency = 0.2f;
            textBox.TextFrame.TextRange.Font.Color.RGB = 0xffffff;

            textBox.Top = slideHeight - textBox.Height;
            return textBox;
        }

        public static void EmbedCaptionsOnCurrentSlide()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide != null)
            {
                RemoveCaptionsFromSlide(currentSlide);
                EmbedCaptionsOnSlide(currentSlide);
            }
        }

        public static void RemoveCaptionsFromCurrentSlide()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide != null)
            {
                RemoveCaptionsFromSlide(currentSlide);
            }
        }

        public static void RemoveCaptionsFromSelectedSlides()
        {
            foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                RemoveCaptionsFromSlide(slide);   
            }
        }

        public static void RemoveCaptionsFromAllSlides()
        {
            foreach (PowerPointSlide s in PowerPointPresentation.Current.Slides)
            {
                RemoveCaptionsFromSlide(s);
            }
        }

        private static void RemoveCaptionsFromSlide(PowerPointSlide slide)
        {
            if (slide != null)
            {
                slide.DeleteShapesWithPrefixTimelineInvariant("PowerPointLabs Caption ");
            }
        }
    }
}
