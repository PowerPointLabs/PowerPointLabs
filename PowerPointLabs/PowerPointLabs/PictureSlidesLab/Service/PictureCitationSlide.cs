using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    class PictureCitationSlide : PowerPointSlide
    {
        public const string PictureCitationSlideName = "PPTLabsPictureCitationSlide";
        public const string PictureCitationSlideTitle = "Picture Sources";

        private List<PowerPointSlide> AllSlides { get; set; }

        public PictureCitationSlide(Slide slide, List<PowerPointSlide> allSlides) : base(slide)
        {
            slide.Name = PictureCitationSlideName;
            AllSlides = allSlides;
        }

        public void CreatePictureCitations()
        {
            var citations = GenerateCitations();
            foreach (Shape shape in Shapes)
            {
                try
                {
                    if (shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                    {
                        continue;
                    }

                    switch (shape.PlaceholderFormat.Type)
                    {
                        case PpPlaceholderType.ppPlaceholderTitle:
                        case PpPlaceholderType.ppPlaceholderCenterTitle:
                        case PpPlaceholderType.ppPlaceholderVerticalTitle:
                            shape.TextFrame2.TextRange.Text = PictureCitationSlideTitle;
                            break;
                        case PpPlaceholderType.ppPlaceholderBody:
                            shape.TextFrame2.TextRange.Text = citations;
                            break;
                    }
                }
                catch (COMException)
                {
                    // non-placeholder shapes don't have PlaceholderFormat
                    // and will cause exception
                }
            }

            DeleteShapesWithPrefix(PptLabsIndicatorShapeName);
            _slide.SlideShowTransition.Hidden = MsoTriState.msoTrue;
        }

        public static bool IsCitationSlide(PowerPointSlide slide)
        {
            if (slide == null)
            {
                return false;
            }

            return slide.Name == PictureCitationSlideName;
        }

        private string GenerateCitations()
        {
            try
            {
                var strBuilder = new StringBuilder("");
                var isAnyCitation = false;
                var slideIndex = 1;
                foreach (var slide in AllSlides)
                {
                    var originalShapeList = slide.GetShapesWithPrefix(
                        EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Original_DO_NOT_REMOVE);
                    if (originalShapeList.Count == 0)
                    {
                        continue;
                    }

                    var originalImageShape = originalShapeList[0];
                    var source = originalImageShape.Tags[Tag.ReloadImgSource];
                    if (string.IsNullOrEmpty(source))
                    {
                        source = "somewhere";
                    }
                    var citation = "Picture taken from " + source;
                    strBuilder.Append("Slide" + slideIndex + ": " + citation + "\n");
                    isAnyCitation = true;
                    slideIndex++;
                }
                if (!isAnyCitation)
                {
                    strBuilder.Append("No citation.");
                }
                else
                {
                    // remove last '\n' char
                    strBuilder.Remove(strBuilder.Length - 1, 1);
                }
                return strBuilder.ToString();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GenerateCitations");
                return "";
            }
        }
    }
}
