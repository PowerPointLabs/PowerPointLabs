using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class AutoAnimate
    {
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
                if (currentSlide == null || currentSlide.Index == PowerPointCurrentPresentationInfo.SlideCount)
                {
                   System.Windows.Forms.MessageBox.Show("Please select the correct slide", "Unable to Add Animations");
                   return;
                }

                PowerPointSlide nextSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(currentSlide.Index);
                if (!GetMatchingShapeDetails(currentSlide, nextSlide))
                {
                    System.Windows.Forms.MessageBox.Show("No matching Shapes were found on the next slide", "Animation Not Added");
                    return;
                }

                AddCompleteAnimations(currentSlide, nextSlide);           
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddAnimationButtonClick");
                throw;
            }
            
        }

        public static void ReloadAutoAnimation()
        {
            try
            {
                var selectedSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                PowerPointSlide currentSlide = null, animatedSlide = null, nextSlide = null;

                if (selectedSlide.Name.StartsWith("PPSlideAnimated"))
                {
                    nextSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index);
                    currentSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index - 2);
                    animatedSlide = selectedSlide;
                    ManageSlidesForReload(currentSlide, nextSlide, animatedSlide);
                }
                else if (selectedSlide.Name.StartsWith("PPSlideStart"))
                {
                    animatedSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index);
                    nextSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index + 1);
                    currentSlide = selectedSlide;
                    ManageSlidesForReload(currentSlide, nextSlide, animatedSlide);
                }
                else if (selectedSlide.Name.StartsWith("PPSlideEnd"))
                {
                    animatedSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index - 2);
                    currentSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index - 3);
                    nextSlide = selectedSlide;
                    ManageSlidesForReload(currentSlide, nextSlide, animatedSlide);
                }
                else if (selectedSlide.Name.StartsWith("PPSlideMulti"))
                {
                    if (selectedSlide.Index > 2)
                    {
                        animatedSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index - 2);
                        currentSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index - 3);
                        nextSlide = selectedSlide;
                        if (animatedSlide.Name.StartsWith("PPSlideAnimated"))
                            ManageSlidesForReload(currentSlide, nextSlide, animatedSlide);
                    }

                    if (selectedSlide.Index < PowerPointCurrentPresentationInfo.SlideCount - 1)
                    {
                        animatedSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index);
                        nextSlide = PowerPointCurrentPresentationInfo.Slides.ElementAt(selectedSlide.Index + 1);
                        currentSlide = selectedSlide;
                        if (animatedSlide.Name.StartsWith("PPSlideAnimated"))
                            ManageSlidesForReload(currentSlide, nextSlide, animatedSlide);
                    }
                }
                else
                    System.Windows.Forms.MessageBox.Show("The current slide was not added by PowerPointLabs Auto Animate", "Error");
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ReloadAutoAnimation");
                throw;
            }
        }

        private static void ManageSlidesForReload(PowerPointSlide currentSlide, PowerPointSlide nextSlide, PowerPointSlide animatedSlide)
        {
            animatedSlide.Delete();
            if (!GetMatchingShapeDetails(currentSlide, nextSlide))
            {
                System.Windows.Forms.MessageBox.Show("No matching Shapes were found on the next slide", "Animation Not Added");
                return;
            }
            AddCompleteAnimations(currentSlide, nextSlide);
        }
        private static void AddCompleteAnimations(PowerPointSlide currentSlide, PowerPointSlide nextSlide)
        {
            var addedSlide = currentSlide.CreateAutoAnimateSlide() as PowerPointAutoAnimateSlide;
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.Index);

            AboutForm progressForm = new AboutForm();
            progressForm.Visible = true;

            addedSlide.MoveMotionAnimation(); //Move shapes with motion animation already added
            addedSlide.PrepareForAutoAnimate();
            RenameCurrentSlide(currentSlide);
            PrepareNextSlide(nextSlide);
            addedSlide.AddAutoAnimation(currentSlideShapes, nextSlideShapes, matchingShapeIDs);
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
            PowerPointLabsGlobals.AddAckSlide();

            progressForm.Visible = false;
        }

        private static void PrepareNextSlide(PowerPointSlide nextSlide)
        {
            if (nextSlide.Transition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFade && nextSlide.Transition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFadeSmoothly)
                nextSlide.Transition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;

            if (nextSlide.Name.StartsWith("PPSlideStart") || nextSlide.Name.StartsWith("PPSlideMulti"))
                nextSlide.Name = "PPSlideMulti" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            else
                nextSlide.Name = "PPSlideEnd" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }

        private static void RenameCurrentSlide(PowerPointSlide currentSlide)
        {
            if (currentSlide.Name.StartsWith("PPSlideEnd") || currentSlide.Name.StartsWith("PPSlideMulti"))
                currentSlide.Name = "PPSlideMulti" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
            else
                currentSlide.Name = "PPSlideStart" + DateTime.Now.ToString("yyyyMMddHHmmssffff");
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
                    tempMatchingShape = nextSlide.GetShapeWithSameName(sh);
                
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
