﻿using System;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.AnimationLab
{
    internal static class AutoAnimate
    {
#pragma warning disable 0618
        public static float defaultDuration = 0.5f;
        public static bool frameAnimationChecked = false;

        private static PowerPoint.Shape[] currentSlideShapes;
        private static PowerPoint.Shape[] nextSlideShapes;
        private static int[] matchingShapeIDs;

        public static void AddAutoAnimation()
        {
            try
            {
                //Get References of current and next slides
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                if (currentSlide == null || currentSlide.Index == PowerPointPresentation.Current.SlideCount)
                {
                    MessageBox.Show(TextCollection.AnimationLabAutoAnimateErrorWrongSlide, 
                                    TextCollection.AnimationLabAutoAnimateErrorDialogTitle);
                    return;
                }

                PowerPointSlide nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];
                if (!GetMatchingShapeDetails(currentSlide, nextSlide))
                {
                    MessageBox.Show(TextCollection.AnimationLabAutoAnimateErrorNoMatchingShapes,
                                    TextCollection.AnimationLabAutoAnimateErrorDialogTitle);
                    return;
                }

                AddCompleteAnimations(currentSlide, nextSlide);           
            }
            catch (Exception e)
            { 
                Logger.LogException(e, "AddAnimationButtonClick");
                ErrorDialogWrapper.ShowDialog("PowerPointLabs", e.Message, e);
            }
            
        }
        
        private static void AddCompleteAnimations(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            var addedSlide = currentSlide.CreateAutoAnimateSlide() as PowerPointAutoAnimateSlide;
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);

            LoadingDialogBox loadingDialog = new LoadingDialogBox(content: TextCollection.AnimationLabAutoAnimateLoadingText);
            loadingDialog.Show();

            addedSlide.MoveMotionAnimation(); //Move shapes with motion animation already added
            addedSlide.PrepareForAutoAnimate();
            RenameCurrentSlide(currentSlide);
            PrepareNextSlide(nextSlide);
            addedSlide.AddAutoAnimation(currentSlideShapes, nextSlideShapes, matchingShapeIDs);
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
            PowerPointPresentation.Current.AddAckSlide();

            loadingDialog.Close();
        }

        private static void PrepareNextSlide(PowerPointSlide nextSlide)
        {
            if (nextSlide.Transition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFade && nextSlide.Transition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFadeSmoothly)
            {
                nextSlide.Transition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            }

            if (nextSlide.Name.StartsWith("PPSlideStart") || nextSlide.Name.StartsWith("PPSlideMulti"))
            {
                nextSlide.Name = "PPSlideMulti" + GetSlideIdentifier();
            }
            else
            {
                nextSlide.Name = "PPSlideEnd" + GetSlideIdentifier();
            }
        }

        private static void RenameCurrentSlide(PowerPointSlide currentSlide)
        {
            if (currentSlide.Name.StartsWith("PPSlideEnd") || currentSlide.Name.StartsWith("PPSlideMulti"))
            {
                currentSlide.Name = "PPSlideMulti" + GetSlideIdentifier();
            }
            else
            {
                currentSlide.Name = "PPSlideStart" + GetSlideIdentifier();
            }
        }

        private static string GetSlideIdentifier()
        {
            return DateTime.Now.ToString("_yyyyMMddHHmmssffff_") +
                   Guid.NewGuid().ToString("N").Substring(0, 7);
        }

        private static bool GetMatchingShapeDetails(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            currentSlideShapes = new PowerPoint.Shape[currentSlide.Shapes.Count];
            nextSlideShapes = new PowerPoint.Shape[currentSlide.Shapes.Count];
            matchingShapeIDs = new int[currentSlide.Shapes.Count];

            int counter = 0;
            PowerPoint.Shape tempMatchingShape = null;
            bool flag = false;
            
            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                tempMatchingShape = nextSlide.GetShapeWithSameIDAndName(sh);
                if (tempMatchingShape == null)
                {
                    tempMatchingShape = nextSlide.GetShapeWithSameName(sh);
                }
                
                if (tempMatchingShape != null)
                {
                    currentSlideShapes[counter] = sh;
                    nextSlideShapes[counter] = tempMatchingShape;
                    matchingShapeIDs[counter] = sh.Id;
                    counter++;
                    flag = true;
                }
            }

            return flag;
        }
    }
}
