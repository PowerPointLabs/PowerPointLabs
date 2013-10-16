using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PowerPointLabs2007
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public PowerPoint.MsoAnimEffect defaultEffect = PowerPoint.MsoAnimEffect.msoAnimEffectCustom;
        public float defaultDuration = 0.5f;
        public bool startUp = false;
        public Dictionary<String, PowerPoint.MsoAnimEffect> effectMapping = new Dictionary<String, PowerPoint.MsoAnimEffect>
        { 
            {"Appear", PowerPoint.MsoAnimEffect.msoAnimEffectAppear},
            {"Arc Up", PowerPoint.MsoAnimEffect.msoAnimEffectArcUp},
            {"Ascend", PowerPoint.MsoAnimEffect.msoAnimEffectAscend},
            {"Blinds", PowerPoint.MsoAnimEffect.msoAnimEffectBlinds},
            {"Checkerboard", PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard},
            {"Circle", PowerPoint.MsoAnimEffect.msoAnimEffectCircle},
            {"Crawl", PowerPoint.MsoAnimEffect.msoAnimEffectCrawl},
            {"Credits", PowerPoint.MsoAnimEffect.msoAnimEffectCredits},
            {"Descend", PowerPoint.MsoAnimEffect.msoAnimEffectDescend},
            {"Diamond", PowerPoint.MsoAnimEffect.msoAnimEffectDiamond},
            {"Dissolve", PowerPoint.MsoAnimEffect.msoAnimEffectDissolve},
            {"Ease In", PowerPoint.MsoAnimEffect.msoAnimEffectEaseIn},
            {"Expand", PowerPoint.MsoAnimEffect.msoAnimEffectExpand},
            {"Fade", PowerPoint.MsoAnimEffect.msoAnimEffectFade},
            {"Faded Swivel", PowerPoint.MsoAnimEffect.msoAnimEffectFadedSwivel},
            {"Faded Zoom", PowerPoint.MsoAnimEffect.msoAnimEffectFadedZoom},
            {"Flash Bulb", PowerPoint.MsoAnimEffect.msoAnimEffectFlashBulb},
            {"Flash Once", PowerPoint.MsoAnimEffect.msoAnimEffectFlashOnce},
            {"Flicker", PowerPoint.MsoAnimEffect.msoAnimEffectFlicker},
            {"Flip", PowerPoint.MsoAnimEffect.msoAnimEffectFlip},
            {"Float", PowerPoint.MsoAnimEffect.msoAnimEffectFloat}
        };

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointLabs2007.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            //Get References of current and next slides
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            if (currentSlide != null && currentSlide.SlideIndex != presentation.Slides.Count)
            {
                PowerPoint.Slide nextSlide = GetNextSlide(currentSlide);
                AddCompleteAutoMotion(currentSlide, nextSlide);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select the correct slide", "Unable to Add Animations");
            }
        }

        public void ReloadButtonClick(Office.IRibbonControl control)
        {
            PowerPoint.Slide tempSlide = GetCurrentSlide();
            if (tempSlide.Name.Contains("PPSlide") && tempSlide.Name.Substring(0, 7).Equals("PPSlide"))
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide nextSlide = presentation.Slides[tempSlide.SlideIndex + 1];
                PowerPoint.Slide currentSlide = presentation.Slides[tempSlide.SlideIndex - 1];
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
                tempSlide.Delete();

                AddCompleteAutoMotion(currentSlide, nextSlide);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The current slide was not added by PPAutoMotion", "Error");
            }
        }

        //Dropdown Callbacks
        public int OnGetItemCount(Office.IRibbonControl control)
        {
            return effectMapping.Count;
        }

        public String OnGetItemLabel(Office.IRibbonControl control, int index)
        {
            String[] keys = effectMapping.Keys.ToArray();
            return keys[index];
        }

        public void OnSelectItem(Office.IRibbonControl control, String selectedId, int selectedIndex)
        {
            String[] keys = effectMapping.Keys.ToArray();
            defaultEffect = effectMapping[keys[selectedIndex]];
        }

        //Duration Callbacks
        public void OnChangeDuration(Office.IRibbonControl control, String text)
        {
            defaultDuration = float.Parse(text);
        }

        public String OnGetText(Office.IRibbonControl control)
        {
            return defaultDuration.ToString();
        }

        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            Form1 editForm = new Form1(this, selectedShape.Name);
            editForm.Show();
        }

        public void nameEdited(String newName)
        {
            PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            selectedShape.Name = newName;
        }

        #endregion

        #region Helpers

        private void AddCompleteAutoMotion(PowerPoint.Slide currentSlide, PowerPoint.Slide nextSlide)
        {
            //Create containers to store information on matching shapes
            PowerPoint.Shape[] shapes1;
            PowerPoint.Shape[] shapes2;
            int[] shapeIDs;

            if (GetMatchingShapeDetails(currentSlide, nextSlide, out shapes1, out shapes2, out shapeIDs))
            {
                //If an identical object exists
                PowerPoint.Slide newSlide = PrepareAnimatedSlide(currentSlide, shapeIDs);
                AddAnimationsToShapes(newSlide, shapes1, shapes2, shapeIDs);
            }
            else
            {
                //Display error message
                System.Windows.Forms.MessageBox.Show("No matching Shapes were found on the next slide", "Animation Not Added");
            }
        }

        private void AddAnimationsToShapes(PowerPoint.Slide newSlide, PowerPoint.Shape[] shapes1, PowerPoint.Shape[] shapes2, int[] shapeIDs)
        {
            int count = 0;
            PowerPoint.Sequence sequence = newSlide.TimeLine.MainSequence;
            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectRotate = null;
            PowerPoint.Effect effectFontResize = null;
            //PowerPoint.Effect effectDisappear = null;
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            PowerPoint.Shape indicatorShape = newSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightTriangle, presentation.PageSetup.SlideWidth - 40, 0, 40.0f, 40.0f);
            indicatorShape.Rotation = 180;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);

            //effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            //effectDisappear.Exit = Office.MsoTriState.msoTrue;
            //effectDisappear.Timing.Duration = 0;

            //Add animation effects to the duplicated objects
            foreach (PowerPoint.Shape sh in newSlide.Shapes)
            {
                DeleteShapeAnnimations(newSlide, sh);
                if (count < shapeIDs.Count() && sh.Id == shapeIDs[count])
                {
                    //Motion Effect
                    float finalX = (shapes2[count].Left + (shapes2[count].Width) / 2);
                    float initialX = (shapes1[count].Left + (shapes1[count].Width) / 2);
                    float finalY = (shapes2[count].Top + (shapes2[count].Height) / 2);
                    float initialY = (shapes1[count].Top + (shapes1[count].Height) / 2);

                    if ((finalX != initialX) || (finalY != initialY))
                    {
                        effectMotion = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        //PowerPoint.AnimationBehavior motion = effectMotion.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                        PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                        effectMotion.Timing.Duration = defaultDuration;
                        //effectMotion.EffectParameters.Relative = Office.MsoTriState.msoTrue;
                        //motion.MotionEffect.FromX = initialX / presentation.PageSetup.SlideWidth * 100;
                        //motion.MotionEffect.FromY = initialY / presentation.PageSetup.SlideHeight * 100;
                        //motion.MotionEffect.ToX = (finalX - initialX) / presentation.PageSetup.SlideWidth * 100;
                        //motion.MotionEffect.ToY = (finalY - initialY) / presentation.PageSetup.SlideHeight * 100;
                        motion.MotionEffect.Path = "M 0 0 C " + ((finalX - initialX) / 2) / presentation.PageSetup.SlideWidth + " " + ((finalY - initialY) / 2) / presentation.PageSetup.SlideHeight + " " + ((finalX - initialX) / 2) / presentation.PageSetup.SlideWidth + " " + ((finalY - initialY) / 2) / presentation.PageSetup.SlideHeight + " " + (finalX - initialX) / presentation.PageSetup.SlideWidth + " " + (finalY - initialY) / presentation.PageSetup.SlideHeight + " E"; 
                    }

                    //Resize Effect
                    if (sh.Type != Office.MsoShapeType.msoPlaceholder && sh.Type != Office.MsoShapeType.msoTextBox)
                    {
                        float finalWidth = shapes2[count].Width;
                        float initialWidth = shapes1[count].Width;
                        float finalHeight = shapes2[count].Height;
                        float initialHeight = shapes1[count].Height;

                        if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
                        {
                            effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            //PowerPoint.AnimationBehavior resize = effectResize.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeScale);
                            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                            effectResize.Timing.Duration = defaultDuration;
                            resize.ScaleEffect.ToX = (finalWidth / initialWidth) * 100;
                            resize.ScaleEffect.ToY = (finalHeight / initialHeight) * 100;
                        }
                    }
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame.HasText == Office.MsoTriState.msoTrue && sh.TextFrame.TextRange.Font.Size != shapes2[count].TextFrame.TextRange.Font.Size)
                    {
                        sh.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                        effectFontResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontSize, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        effectFontResize.Timing.Duration = defaultDuration;
                        PowerPoint.AnimationBehavior resizeFont = effectFontResize.Behaviors[1];
                        resizeFont.PropertyEffect.To = shapes2[count].TextFrame.TextRange.Font.Size / shapes1[count].TextFrame.TextRange.Font.Size;
                    }

                    //Rotation Effect
                    float finalRotation = shapes2[count].Rotation;
                    float initialRotation = shapes1[count].Rotation;
                    if (finalRotation != initialRotation)
                    {
                        effectRotate = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        effectRotate.Timing.Duration = defaultDuration;
                        effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                    }

                    count++;
                }
            }
        }

        private PowerPoint.Slide PrepareAnimatedSlide(PowerPoint.Slide currentSlide, int[] shapeIDs)
        {
            //Duplicate current slide
            currentSlide.Duplicate();
            //Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, currentSlide.CustomLayout);

            //Store reference of new slide
            PowerPoint.Slide newSlide = GetNextSlide(currentSlide);
            newSlide.Name = "PPSlide" + GetTimestamp(DateTime.Now);

            //Delete non-identical shapes
            foreach (PowerPoint.Shape sh in newSlide.Shapes)
            {
                if (!shapeIDs.Contains(sh.Id))
                {
                    sh.Delete();
                }
            }

            //Go to new slide
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(newSlide.SlideIndex);

            //Manage Slide Transitions
            newSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
            newSlide.SlideShowTransition.Duration = defaultDuration;
            newSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            newSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            newSlide.SlideShowTransition.AdvanceTime = 0;

            PowerPoint.Slide nextSlide = GetNextSlide(newSlide);
            nextSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;

            return newSlide;
        }

        private bool GetMatchingShapeDetails(PowerPoint.Slide currentSlide, PowerPoint.Slide nextSlide, out PowerPoint.Shape[] shapes1, out PowerPoint.Shape[] shapes2, out int[] shapeIDs)
        {
            bool flag = false;
            int counter = 0;

            shapes1 = new PowerPoint.Shape[currentSlide.Shapes.Count];
            shapes2 = new PowerPoint.Shape[currentSlide.Shapes.Count];
            shapeIDs = new int[currentSlide.Shapes.Count];

            foreach (PowerPoint.Shape sh1 in currentSlide.Shapes)
            {
                foreach (PowerPoint.Shape sh2 in nextSlide.Shapes)
                {
                    if (haveSameNames(sh1, sh2))
                    {
                        flag = true;
                        if (sh1.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            sh1.TextFrame.TextRange.Text.Trim();
                            sh1.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                        if (sh2.Type == Office.MsoShapeType.msoPlaceholder)
                        {
                            sh2.TextFrame.TextRange.Text.Trim();
                            sh2.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }

                        shapes1[counter] = sh1;
                        shapes2[counter] = sh2;
                        shapeIDs[counter] = sh1.Id;
                        counter++;
                        break;
                    }
                }
            }
            return flag;
        }

        private PowerPoint.Slide GetCurrentSlide()
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            return Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
        }

        private PowerPoint.Slide GetNextSlide(PowerPoint.Slide currentSlide)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            int slideIndex = currentSlide.SlideIndex;
            return presentation.Slides[slideIndex + 1];
        }

        private bool haveSameNames(PowerPoint.Shape sh1, PowerPoint.Shape sh2)
        {
            String name1 = sh1.Name;
            String name2 = sh2.Name;

            if (name1.ToUpper().CompareTo(name2.ToUpper()) == 0)
                return true;
            else
                return false;
        }

        public String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }

        private float GetMinimumRotation(float fromAngle, float toAngle)
        {
            fromAngle = Normalize(fromAngle);
            toAngle = Normalize(toAngle);

            float rotation1 = toAngle - fromAngle;
            float rotation2 = rotation1 == 0 ? 0 : Math.Abs(360 - Math.Abs(rotation1))*(rotation1/Math.Abs(rotation1))*-1;

            if (Math.Abs(rotation1) < Math.Abs(rotation2))
            {
                return rotation1;
            }
            else
            {
                return rotation2;
            }
        }

        private float Normalize(float i)
        {
            //find effective angle
            float d = Math.Abs(i) % 360;

            if (i < 0)
            {
                //--return positive equivalent--
                return 360 - d;
            }
            else
            {
                return d;
            }
        }

        private void DeleteShapeAnnimations(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name)
                effect.Delete();
            }

            PowerPoint.Slide nextSlide = GetNextSlide(slide);
            sequence = nextSlide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name)
                    effect.Delete();
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
