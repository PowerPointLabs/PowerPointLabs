using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
#pragma warning disable 618
    public class CaptionService
    {
        private static Shape templatedShape;

        public static Effect CreateAppearEffectCaptionAnimation(PowerPointSlide slide, string captionText, 
            int clickNo, int tagNo, bool isSeperateClick)
        {
            Shape shape = InsertCaptionShapeToSlide(slide, captionText, tagNo);
            if (isSeperateClick)
            {
                return AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectAppear, clickNo - 1);
            }
            else
            {
                return AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectAppear, clickNo);
            }
        }

        public static Effect CreateExitEffectCaptionAnimation(PowerPointSlide slide, Shape shape, int clickNo)
        {
            Effect effect = AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectAppear, clickNo + 1);
            effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            return effect;
        }
        
        public static void SetShapeAsHidden(PowerPointSlide slide, string captionText, int tagNo)
        {
            Shape shape = InsertCaptionShapeToSlide(slide, captionText, tagNo);
            shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;            
        }

        public static void DeleteCaptionShape(PowerPointSlide slide, int tagNo)
        {
            string shapeName = string.Format(ELearningLabText.CaptionShapeNameFormat, tagNo);
            slide.DeleteShapeWithName(shapeName);
        }

        private static Shape InsertCaptionShapeToSlide(PowerPointSlide slide, string captionText, int tagNo)
        {
            string shapeName = string.Format(ELearningLabText.CaptionShapeNameFormat, tagNo);
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            if (slide.ContainShapeWithExactName(shapeName))
            {
                Shape shape = ShapeUtility.ReplaceTextForShape(slide.GetShapeWithName(shapeName)[0], captionText);
                shape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                templatedShape = shape;
                shape.Top = slideHeight - shape.Height;
                return shape;
            }
            try
            {
                return ShapeUtility.InsertTemplatedShapeToSlide(slide, templatedShape, shapeName, captionText);
            }
            catch
            {
                return ShapeUtility.InsertDefaultCaptionBoxToSlide(slide, shapeName, captionText);
            }
        }
    }
}
