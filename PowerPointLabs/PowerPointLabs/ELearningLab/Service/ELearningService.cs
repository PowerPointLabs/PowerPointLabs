using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Views;

namespace PowerPointLabs.ELearningLab.Service
{
    public class ELearningService
    {
        public static bool IsELearningWorkspaceEnabled { get; set; } = false;
        public static void SyncLabItemToAnimationPane(PowerPointSlide slide, List<SelfExplanationClickItem> selfExplanationItems)
        {
            SyncAppearEffectAnimations(slide, selfExplanationItems);
            SyncExitEffectAnimations(slide, selfExplanationItems);
        }
        public static void DeleteShapesForUnusedItem(PowerPointSlide slide, SelfExplanationClickItem selfExplanationClickItem)
        {
            CalloutService.DeleteCalloutShape(slide, selfExplanationClickItem.tagNo);
            CaptionService.DeleteCaptionShape(slide, selfExplanationClickItem.tagNo);
        }
        private static void SyncAppearEffectAnimations(PowerPointSlide slide, List<SelfExplanationClickItem> selfExplanationItems)
        {
         //   slide.DeleteShapeWithNameRegex(ELearningLabText.VoiceShapeNameRegex);
            ProcessingStatusForm progressBarForm = new ProcessingStatusForm();
            progressBarForm.Show();

            int totalSelfExplanationItemsCount = selfExplanationItems.Count();
            for (int i = 0; i < totalSelfExplanationItemsCount; i++)
            {
                int percentage = (int)Math.Round(((double)i + 1) / totalSelfExplanationItemsCount * 100);
                progressBarForm.UpdateProgress(percentage);
                progressBarForm.UpdateSlideNumber(i, totalSelfExplanationItemsCount);

                SelfExplanationClickItem selfExplanationItem = selfExplanationItems.ElementAt(i);
                if (!selfExplanationItem.IsCaption && !selfExplanationItem.IsCallout && !selfExplanationItem.IsVoice)
                {
                    DeleteShapesForUnusedItem(slide, selfExplanationItem);
                }
                    CreateAppearEffectAnimation(slide, selfExplanationItem);
            }

            DeleteUnusedCalloutShapes(slide);
            DeleteUnusedAudioShapes(slide);
            progressBarForm.Close();
        }

        private static void SyncExitEffectAnimations(PowerPointSlide slide, List<SelfExplanationClickItem> selfExplanationItems)
        {
            foreach (SelfExplanationClickItem selfExplanationItem in selfExplanationItems)
            {
                CreateExitEffectAnimation(slide, selfExplanationItem);
            }
        }

        private static void CreateAppearEffectAnimation(PowerPointSlide slide, SelfExplanationClickItem selfExplanationItem)
        {
            bool isSeparateClick = selfExplanationItem.TriggerIndex == (int)TriggerType.OnClick || !selfExplanationItem.IsTriggerTypeComboBoxEnabled;
            List<Effect> effects = new List<Effect>();
            if (selfExplanationItem.IsVoice)
            {
                Effect effect = AudioService.CreateAppearEffectAudioAnimation(slide, selfExplanationItem.CaptionText, selfExplanationItem.VoiceLabel,
                    selfExplanationItem.ClickNo, selfExplanationItem.tagNo, isSeparateClick);
                if (effect != null)
                {
                    effects.Add(effect);
                }
            }
            if (selfExplanationItem.IsCallout)
            {
                Effect effect = CalloutService.CreateAppearEffectCalloutAnimation(slide, selfExplanationItem.CalloutText, 
                    selfExplanationItem.ClickNo, selfExplanationItem.tagNo, isSeparateClick);
                effects.Add(effect);
            }
            else
            {
                CalloutService.DeleteCalloutShape(slide, selfExplanationItem.tagNo);
            }
            if (selfExplanationItem.IsCaption)
            {
                Effect effect = CaptionService.CreateAppearEffectCaptionAnimation(slide, selfExplanationItem.CaptionText,
                    selfExplanationItem.ClickNo, selfExplanationItem.tagNo, isSeparateClick);
                effects.Add(effect);
            }
            else
            {
                CaptionService.DeleteCaptionShape(slide, selfExplanationItem.tagNo);
            }
            if (isSeparateClick && effects.Count() > 0 && selfExplanationItem.ClickNo > 0)
            {
                effects.First().Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
        }

        private static void CreateExitEffectAnimation(PowerPointSlide slide, SelfExplanationClickItem selfExplanationItem)
        {
            string calloutShapeName = string.Format(ELearningLabText.CalloutShapeNameFormat, selfExplanationItem.tagNo);
            string captionShapeName = string.Format(ELearningLabText.CaptionShapeNameFormat, selfExplanationItem.tagNo);
            if (selfExplanationItem.IsCallout && slide.ContainShapeWithExactName(calloutShapeName))
            {
                Shape calloutShape = slide.GetShapeWithName(calloutShapeName)[0];
                CalloutService.CreateExitEffectCalloutAnimation(slide, calloutShape, selfExplanationItem.ClickNo);
            }
            if (selfExplanationItem.IsCaption && slide.ContainShapeWithExactName(captionShapeName))
            {
                Shape captionShape = slide.GetShapeWithName(captionShapeName)[0];
                CaptionService.CreateExitEffectCaptionAnimation(slide, captionShape, selfExplanationItem.ClickNo);
            }
        }

        private static void DeleteUnusedAudioShapes(PowerPointSlide slide)
        {
            List<Shape> shapes = slide.GetShapesWithNameRegex(ELearningLabText.VoiceShapeNameRegex);
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            foreach (Effect effect in effects)
            {
                if (shapes.Contains(effect.Shape))
                {
                    shapes.Remove(effect.Shape);
                }
            }
            foreach (Shape shape in shapes)
            {
                shape.Delete();
            }
        }

        private static void DeleteUnusedCalloutShapes(PowerPointSlide slide)
        {
            List<Shape> shapes = slide.GetShapesWithNameRegex(ELearningLabText.CalloutShapeNameRegex);
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            foreach (Effect effect in effects)
            {
                if (shapes.Contains(effect.Shape))
                {
                    shapes.Remove(effect.Shape);
                }
            }
            foreach (Shape shape in shapes)
            {
                shape.Delete();
            }
        }
    }
}
