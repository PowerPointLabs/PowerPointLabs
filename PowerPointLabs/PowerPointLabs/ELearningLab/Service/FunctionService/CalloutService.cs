using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
    public class CalloutService
    {
        private static Shape templatedShape;

        public static Effect CreateAppearEffectCalloutAnimation(PowerPointSlide slide, string calloutText,
            int clickNo, int tagNo, bool isSeperateClick)
        {
            Shape shape = InsertCalloutShapeToSlide(slide, calloutText, tagNo);
            if (isSeperateClick)
            {
                return AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectAppear, clickNo - 1);
            }
            else
            {
                return AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectAppear, clickNo);
            }
        }

        public static Effect CreateExitEffectCalloutAnimation(PowerPointSlide slide, Shape shape, int clickNo)
        {
            Effect effect = AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectAppear, clickNo + 1);
            effect.Exit = Microsoft.Office.Core.MsoTriState.msoTrue;
            return effect;
        }

        public static void DeleteCalloutShape(PowerPointSlide slide, int tagNo)
        {
            string shapeName = string.Format(ELearningLabText.CalloutShapeNameFormat, tagNo);
            slide.DeleteShapeWithName(shapeName);
        }

        private static Shape InsertCalloutShapeToSlide(PowerPointSlide slide, string calloutText, int tagNo)
        {
            string shapeName = string.Format(ELearningLabText.CalloutShapeNameFormat, tagNo);
            if (slide.ContainShapeWithExactName(shapeName))
            {
                Shape shape = ShapeUtility.ReplaceTextForShape(slide.GetShapeWithName(shapeName)[0], calloutText);
                templatedShape = shape;
                return shape;
            }
            try
            {
                return ShapeUtility.InsertTemplatedShapeToSlide(slide, templatedShape, shapeName, calloutText);
            }
            catch
            {
                return ShapeUtility.InsertDefaultCalloutBoxToSlide(slide, shapeName, calloutText);
            }
        }
    }
}
