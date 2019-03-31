using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.Extensions
{
    public static class PowerPointSlideExtension
    {
        public static bool IsFirstAnimationTriggeredByClick(this PowerPointSlide slide)
        {
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            return effects.Count() > 0 &&
                effects.ElementAt(0).Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick;
        }

        public static IEnumerable<Effect> GetCustomEffectsForClick(this PowerPointSlide slide, int clickNo)
        {
            DateTime start = DateTime.Now;
            List<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>().ToList();
            Sequence sequence = slide.TimeLine.MainSequence;
            try
            {
                Effect effectBefore = sequence.FindFirstAnimationForClick(clickNo);
                Effect effectAfter = sequence.FindFirstAnimationForClick(clickNo + 1);
                // from idxStart inclusive, idxEnd exclusive
                int idxStart = effectBefore == null ? effects.Count() : effects.IndexOf(effectBefore);
                int idxEnd = effectAfter == null ? effects.Count() : effects.IndexOf(effectAfter);

                IEnumerable<Effect> customEffects = effects.GetRange(idxStart, idxEnd - idxStart).Where(x =>
                SelfExplanationTagService.ExtractTagNo(x.Shape.Name) == -1);
                if (clickNo > 0 && customEffects.Count() > 0)
                {
                    customEffects.ElementAt(0).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                }
                return customEffects;
            }
            catch
            {
                // most likely caused by idxStart out of bound
                return new List<Effect>();
            }
        }
        public static IEnumerable<Effect> GetPPTLEffectsForClick(this PowerPointSlide slide, int clickNo)
        {
            DateTime start = DateTime.Now;
            List<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>().ToList();
            Sequence sequence = slide.TimeLine.MainSequence;
            try
            {
                Effect effectBefore = sequence.FindFirstAnimationForClick(clickNo);
                Effect effectAfter = sequence.FindFirstAnimationForClick(clickNo + 1);
                // from idxStart inclusive, idxEnd exclusive
                int idxStart = effectBefore == null ? effects.Count() : effects.IndexOf(effectBefore);
                int idxEnd = effectAfter == null ? effects.Count() : effects.IndexOf(effectAfter);
                return effects.GetRange(idxStart, idxEnd - idxStart).Where(x =>
                SelfExplanationTagService.ExtractTagNo(x.Shape.Name) != -1 && 
                x.Exit != Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch
            {
                // most likely caused by idxStart out of bound
                return new List<Effect>();
            }
        }

        public static void RemoveAnimationsForShapeWithPrefix(this PowerPointSlide slide, string prefix)
        {
            slide.RemoveAnimationsForShapes(
                slide.Shapes.Cast<Shape>().Where(x => x.Name.Contains(prefix)).ToList());           
        }

        public static bool ContainShapeWithExactName(this PowerPointSlide slide, string shapeName)
        {
            return slide.Shapes.Cast<Shape>().Where(x => x.Name.Trim().Equals(shapeName.Trim())).Count() > 0;
        }

        public static bool ContainShape(this PowerPointSlide slide, Shape shape)
        {
            return slide.Shapes.Cast<Shape>().Where(x => x.Equals(shape)).Count() > 0;
        }

        public static void DeleteShapeWithNameRegex(this PowerPointSlide slide, string regexExpr)
        {
            Regex regex = new Regex(regexExpr);
            List<Shape> shapes = slide.Shapes.Cast<Shape>().Where(x => regex.Match(x.Name).Success).ToList();
            foreach (Shape s in shapes)
            {
                s.Delete();
            }
        }

        public static List<Shape> GetShapesWithNameRegex(this PowerPointSlide slide, string regexExpr)
        {
            Regex regex = new Regex(regexExpr);
            return slide.Shapes.Cast<Shape>().Where(x => regex.Match(x.Name).Success).ToList();
        }
    }
}
