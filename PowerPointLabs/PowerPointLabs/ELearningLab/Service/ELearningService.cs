using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
    public class ELearningService
    {
        public static bool IsELearningWorkspaceEnabled { get; set; } = false;
        private PowerPointSlide _slide;
        private List<ExplanationItem> _selfExplanationItems;

        public ELearningService() { }
        public ELearningService(PowerPointSlide slide, List<ExplanationItem> selfExplanationItems)
        {
            _slide = slide;
            _selfExplanationItems = selfExplanationItems;
        }
        public void DeleteShapesForUnusedItem(PowerPointSlide slide, ExplanationItem selfExplanationClickItem)
        {
            CalloutService.DeleteCalloutShape(slide, selfExplanationClickItem.tagNo);
            CaptionService.DeleteCaptionShape(slide, selfExplanationClickItem.tagNo);
        }

        public void SyncExitEffectAnimations()
        {
            SyncExitEffectAnimations(_slide, _selfExplanationItems);
            DeleteUnusedCalloutShapes(_slide);
            DeleteUnusedAudioShapes(_slide);
            DeleteUnusedCaptionShapes(_slide);
            DeleteExitAnimationInLastClick(_slide);
        }

        public void SyncAppearEffectAnimationsForSelfExplanationItem(int i)
        {
            ExplanationItem selfExplanationItem = _selfExplanationItems.ElementAt(i);
            CreateAppearEffectAnimation(_slide, selfExplanationItem);
        }

        public void RemoveLabAnimationsFromAnimationPane()
        {
            _slide.RemoveAnimationsForShapeWithPrefix(ELearningLabText.Identifier);
        }

        public int GetExplanationItemsCount()
        {
            return _selfExplanationItems.Count;
        }

        private void SyncExitEffectAnimations(PowerPointSlide slide, List<ExplanationItem> selfExplanationItems)
        {
            foreach (ExplanationItem selfExplanationItem in selfExplanationItems)
            {
                CreateExitEffectAnimation(slide, selfExplanationItem);
            }
        }

        private void CreateAppearEffectAnimation(PowerPointSlide slide, ExplanationItem selfExplanationItem)
        {
            bool isSeparateClick = selfExplanationItem.TriggerIndex == (int)TriggerType.OnClick || !selfExplanationItem.IsTriggerTypeComboBoxEnabled;
            List<Effect> effects = new List<Effect>();
            if (selfExplanationItem.IsVoice)
            {
                Effect effect;
                if (selfExplanationItem.IsVoiceLabelInvalid)
                {
                    string defaultVoiceLabel = string.Format(ELearningLabText.AudioDefaultLabelFormat, AudioSettingService.selectedVoice.VoiceName);
                    effect = AudioService.CreateAppearEffectAudioAnimation(slide, selfExplanationItem.CaptionText, defaultVoiceLabel,
                        selfExplanationItem.ClickNo, selfExplanationItem.tagNo, isSeparateClick);
                    selfExplanationItem.VoiceLabel = defaultVoiceLabel;
                }
                else
                {
                    effect = AudioService.CreateAppearEffectAudioAnimation(slide, selfExplanationItem.CaptionText, selfExplanationItem.VoiceLabel, 
                        selfExplanationItem.ClickNo, selfExplanationItem.tagNo, isSeparateClick);
                }
                if (effect != null)
                {
                    effects.Add(effect);
                }
            }
            if (selfExplanationItem.IsCallout)
            {
                string calloutText = selfExplanationItem.HasShortVersion ? selfExplanationItem.CalloutText : selfExplanationItem.CaptionText;
                Effect effect = CalloutService.CreateAppearEffectCalloutAnimation(slide, calloutText,
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

        private void CreateExitEffectAnimation(PowerPointSlide slide, ExplanationItem selfExplanationItem)
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

        private void DeleteExitAnimationInLastClick(PowerPointSlide slide)
        {
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            List<Effect> effectsToDelete = new List<Effect>();

            for (int i = effects.Count() - 1; i > 0 && i >= effects.Count() - 2; i--)
            {
                Effect effect = effects.ElementAt(i);
                bool isTriggeredByClick = effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                if (effect.Exit != Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    return;
                }
                string shapeName = effect.Shape.Name;
                if (StringUtility.IsPPTLShape(shapeName) && effect.Exit == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    effectsToDelete.Add(effect);
                    //  effect.Delete();
                }
                if (isTriggeredByClick)
                {
                    if (effectsToDelete.Count + i == effects.Count())
                    {
                        foreach (Effect _effect in effectsToDelete)
                        {
                            _effect.Delete();
                        }
                    }
                    return;
                }
            }
        }

        private void DeleteUnusedAudioShapes(PowerPointSlide slide)
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

        private void DeleteUnusedCalloutShapes(PowerPointSlide slide)
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

        private void DeleteUnusedCaptionShapes(PowerPointSlide slide)
        {
            List<Shape> shapes = slide.GetShapesWithNameRegex(ELearningLabText.CaptionShapeNameRegex);
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
