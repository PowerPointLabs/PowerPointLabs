using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPointLabs.Views;
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


namespace PowerPointLabs
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        public bool frameAnimationChecked = false;
        public bool backgroundZoomChecked = true;
        public bool multiSlideZoomChecked = false;
        public bool spotlightDelete = true;
        public float defaultDuration = 0.5f;
        public bool spotlightEnabled = false;
        public bool inSlideEnabled = false;
        public bool zoomButtonEnabled = false;
        public bool highlightBulletsEnabled = true;
        public bool addAutoMotionEnabled = true;
        public bool reloadAutoMotionEnabled = true;
        public bool reloadSpotlight = true;

        private List<PowerPoint.MsoAnimEffect> entryEffects = new List<PowerPoint.MsoAnimEffect>()
        {
            PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimEffect.msoAnimEffectBlinds, PowerPoint.MsoAnimEffect.msoAnimEffectBox,
            PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard, PowerPoint.MsoAnimEffect.msoAnimEffectCircle, PowerPoint.MsoAnimEffect.msoAnimEffectDiamond,
            PowerPoint.MsoAnimEffect.msoAnimEffectDissolve, PowerPoint.MsoAnimEffect.msoAnimEffectFly, PowerPoint.MsoAnimEffect.msoAnimEffectPeek, 
            PowerPoint.MsoAnimEffect.msoAnimEffectPlus, PowerPoint.MsoAnimEffect.msoAnimEffectRandomBars, PowerPoint.MsoAnimEffect.msoAnimEffectSplit,
            PowerPoint.MsoAnimEffect.msoAnimEffectStrips, PowerPoint.MsoAnimEffect.msoAnimEffectWedge, PowerPoint.MsoAnimEffect.msoAnimEffectWheel,
            PowerPoint.MsoAnimEffect.msoAnimEffectWipe, PowerPoint.MsoAnimEffect.msoAnimEffectExpand, PowerPoint.MsoAnimEffect.msoAnimEffectFade,
            PowerPoint.MsoAnimEffect.msoAnimEffectFadedSwivel, PowerPoint.MsoAnimEffect.msoAnimEffectFadedZoom, PowerPoint.MsoAnimEffect.msoAnimEffectZoom,
            PowerPoint.MsoAnimEffect.msoAnimEffectCenterRevolve, PowerPoint.MsoAnimEffect.msoAnimEffectFloat, PowerPoint.MsoAnimEffect.msoAnimEffectGrowAndTurn,
            PowerPoint.MsoAnimEffect.msoAnimEffectRiseUp, PowerPoint.MsoAnimEffect.msoAnimEffectSpinner, PowerPoint.MsoAnimEffect.msoAnimEffectSwivel,
            PowerPoint.MsoAnimEffect.msoAnimEffectBoomerang, PowerPoint.MsoAnimEffect.msoAnimEffectBounce, PowerPoint.MsoAnimEffect.msoAnimEffectCredits,
            PowerPoint.MsoAnimEffect.msoAnimEffectFlip, PowerPoint.MsoAnimEffect.msoAnimEffectFloat, PowerPoint.MsoAnimEffect.msoAnimEffectPinwheel,
            PowerPoint.MsoAnimEffect.msoAnimEffectSpiral, PowerPoint.MsoAnimEffect.msoAnimEffectWhip
        };

        private bool _allSlides;
        private bool _previewCurrentSlide;
        private bool _captionsAllSlides;

        private List<string> _voiceNames;

        private int _voiceSelected = 0;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            SetVoicesFromInstalledOptions();
            SetCoreVoicesToSelections();
        }

        public void RefreshRibbonControl(String controlID)
        {
            try
            {
                ribbon.InvalidateControl(controlID);
            }
            catch (Exception e)
            {
                LogException(e, "RefreshRibbonControl");
                throw;
            }
        }

        private void SetVoicesFromInstalledOptions()
        {
            var installedVoices = NotesToAudio.GetVoices().ToList();
            _voiceNames = installedVoices;
        }

        public void HighlightBulletsBackgroundButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kTextSelected;
                else
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kNoneSelected;

                HighlightBulletsBackground.AddHighlightBulletsBackground();
            }
            catch (Exception e)
            {
                LogException(e, "HighlightBulletsBackgroundButtonClick");
                throw;
            }
        }

        public void HighlightBulletsTextButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kTextSelected;
                else
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kNoneSelected;

                HighlightBulletsText.AddHighlightBulletsText();
            }
            catch (Exception e)
            {
                LogException(e, "HighlightBulletsTextButtonClick");
                throw;
            }
        }

        public void AddInSlideAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AnimateInSlide.isHighlightBullets = false;
                AnimateInSlide.AddAnimationInSlide();
            }
            catch (Exception e)
            {
                LogException(e, "AddInSlideAnimationButtonClick");
                throw;
            }
        }
        public void ReloadSpotlightButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Spotlight.ReloadSpotlightEffect();
            }
            catch (Exception e)
            {
                LogException(e, "ReloadSpotlightButtonClick");
                throw;
            }
        }
        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            try
            {
                Spotlight.AddSpotlightEffect();
            }
            catch (Exception e)
            {
                LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        //Button Click Callbacks
        private PowerPoint.Shape FindIdenticalShape(PowerPoint.Slide slideToSearch, PowerPoint.Shape shapeToSearch)
        {
            PowerPoint.Shape shapeToReturn = null;
            foreach (PowerPoint.Shape sh in slideToSearch.Shapes)
            {
                if (sh.Id == shapeToSearch.Id && sh.Name.Equals(shapeToSearch.Name))
                {
                    shapeToReturn = sh;
                    break;
                }
            }
            return shapeToReturn;
        }
       
        
        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.AddAutoAnimation();
            }
            catch (Exception e)
            {
                LogException(e, "AddAnimationButtonClick");
                throw;
            }
        }
        public void ReloadButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.ReloadAutoAnimation();
            }
            catch (Exception e)
            {
                LogException(e, "ReloadAnimationButtonClick");
                throw;
            }
        }
        private PowerPoint.Slide AddMagnifyingSlideWithBackground(PowerPoint.Slide currentSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = currentSlide.Duplicate()[1];
            MoveMotionAnimation(addedSlide);
            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (HasExitAnimation(addedSlide, tmp))
                    tmp.Delete();
            }
            addedSlide.Name = "PPTLabsMagnifyingSlide" + GetTimestamp(DateTime.Now);

            float centerX = selectedShape.Left + selectedShape.Width / 2;
            float centerY = selectedShape.Top + selectedShape.Height / 2;

            PowerPoint.Shape identicalShape = FindIdenticalShape(addedSlide, selectedShape);
            if (identicalShape != null)
            {
                identicalShape.Delete();
            }

            PowerPoint.Slide tempSlide = addedSlide.Duplicate()[1];
            addedSlide.Copy();
            PowerPoint.Shape duplicatePic = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];

            float scaleFactorX = presentation.PageSetup.SlideWidth / duplicatePic.Width;
            float scaleFactorY = presentation.PageSetup.SlideHeight / duplicatePic.Height;

            duplicatePic.LockAspectRatio = Office.MsoTriState.msoFalse;
            duplicatePic.Left = 0;
            duplicatePic.Top = 0;
            duplicatePic.Width = presentation.PageSetup.SlideWidth;
            duplicatePic.Height = presentation.PageSetup.SlideHeight;
            duplicatePic.Name = "PPTLabsMagnifyAreaSlide" + GetTimestamp(DateTime.Now);

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);

            PowerPoint.Shape cropShape = tempSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, presentation.PageSetup.SlideWidth - 1, presentation.PageSetup.SlideHeight - 1);
            cropShape.Select();
            //PowerPoint.Shape tempDuplicate = duplicatePic.Duplicate()[1];
            //tempDuplicate.Left = 0;
            //tempDuplicate.Top = 0;
            //tempDuplicate.Select();
            //foreach (PowerPoint.Shape sh in tempSlide.Shapes)
            //{
            //    if (sh.Visible == Office.MsoTriState.msoTrue)
            //    {
            //        PowerPoint.Shape dupShape = null;
            //        sh.Copy();
            //        if (sh.Type == Office.MsoShapeType.msoPlaceholder)
            //        {
            //            dupShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            //        }
            //        else
            //        {
            //            dupShape = addedSlide.Shapes.Paste()[1];
            //        }

            //        dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            //        dupShape.Width = sh.Width;
            //        dupShape.Height = sh.Height;
            //        dupShape.Left = sh.Left;
            //        dupShape.Top = sh.Top;
            //        dupShape.Select(Office.MsoTriState.msoFalse);
            //    }
            //}

            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape croppedShape = CropShapeToSlide(ref sel);

            //tempSlide.Delete();
            croppedShape.Cut();
            PowerPoint.Shape duplicatePic2 = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            duplicatePic2.LockAspectRatio = Office.MsoTriState.msoFalse;
            duplicatePic2.Left = 0;
            duplicatePic2.Top = 0;
            duplicatePic2.Width = presentation.PageSetup.SlideWidth;
            duplicatePic2.Height = presentation.PageSetup.SlideHeight;
            duplicatePic2.Name = "PPTLabsMagnifyAreaGroup" + GetTimestamp(DateTime.Now);

            duplicatePic2.PictureFormat.CropLeft += selectedShape.Left;
            duplicatePic2.PictureFormat.CropTop += selectedShape.Top;
            duplicatePic2.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width));
            duplicatePic2.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height));

            duplicatePic2.Left = centerX - (duplicatePic2.Width / 2);
            duplicatePic2.Top = centerY - (duplicatePic2.Height / 2);
            duplicatePic2.Visible = Office.MsoTriState.msoFalse;
            tempSlide.Delete();

            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectDisappear = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;

            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(duplicatePic) || tmp.Equals(indicatorShape) || tmp.Equals(duplicatePic2)))
                {
                    DeleteShapeAnnimations(addedSlide, tmp);
                    tmp.Visible = Office.MsoTriState.msoFalse;
                }
            }

            selectedShape.Copy();
            PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (magnifyShape.Width > magnifyShape.Height)
                magnifyShape.Width = presentation.PageSetup.SlideWidth;
            else
                magnifyShape.Height = presentation.PageSetup.SlideHeight;

            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

            float finalWidth = magnifyShape.Width;
            float initialWidth = selectedShape.Width;
            float finalHeight = magnifyShape.Height;
            float initialHeight = selectedShape.Height;

            float finalX = (magnifyShape.Left + (magnifyShape.Width) / 2) * (finalWidth / initialWidth);
            float initialX = (selectedShape.Left + (selectedShape.Width) / 2) * (finalWidth / initialWidth);
            float finalY = (magnifyShape.Top + (magnifyShape.Height) / 2) * (finalHeight / initialHeight);
            float initialY = (selectedShape.Top + (selectedShape.Height) / 2) * (finalHeight / initialHeight);

            magnifyShape.Delete();

            effectMotion = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
            effectMotion.Timing.Duration = 0.5f;
            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

            effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
            effectResize.Timing.Duration = 0.5f;
            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceTime = 0;
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape tmp in addedSlide.NotesPage.Shapes)
                {
                    if (tmp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        tmp.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (tmp.Type == Office.MsoShapeType.msoMedia)
                    tmp.Delete();
            }
            return addedSlide;
        }
        private PowerPoint.Slide AddMagnifiedSlideWithBackground(PowerPoint.Slide magnifyingSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = magnifyingSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsMagnifiedSlide" + GetTimestamp(DateTime.Now);
            PowerPoint.Shape duplicatePic = null;
            foreach (PowerPoint.Shape sh in magnifyingSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (!sh.Name.Contains("PPTLabsMagnifyAreaGroup"))
                {
                    tmp.Delete();
                }
                else
                {
                    duplicatePic = tmp;
                }
            }

            duplicatePic.Visible = Office.MsoTriState.msoTrue;
            DeleteShapeAnnimations(addedSlide, duplicatePic);
            duplicatePic.LockAspectRatio = Office.MsoTriState.msoTrue;
            //magnifyShape.Left = 0;
            //magnifyShape.Top = 0;
            if (duplicatePic.Width > duplicatePic.Height)
                duplicatePic.Width = presentation.PageSetup.SlideWidth;
            else
                duplicatePic.Height = presentation.PageSetup.SlideHeight;

            duplicatePic.Left = (presentation.PageSetup.SlideWidth / 2) - (duplicatePic.Width / 2);
            duplicatePic.Top = (presentation.PageSetup.SlideHeight / 2) - (duplicatePic.Height / 2);
            duplicatePic.PictureFormat.CropLeft = 0;
            duplicatePic.PictureFormat.CropTop = 0;
            duplicatePic.PictureFormat.CropRight = 0;
            duplicatePic.PictureFormat.CropBottom = 0;


            //selectedShape.Copy();
            //PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
            //magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            //if (magnifyShape.Width > magnifyShape.Height)
            //    magnifyShape.Width = presentation.PageSetup.SlideWidth;
            //else
            //    magnifyShape.Height = presentation.PageSetup.SlideHeight;

            //magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            //magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

            //float finalWidth = magnifyShape.Width;
            //float initialWidth = selectedShape.Width;
            //float finalHeight = magnifyShape.Height;
            //float initialHeight = selectedShape.Height;

            //selectedShape.Copy();
            //PowerPoint.Shape magnifyShape2 = addedSlide.Shapes.Paste()[1];
            //magnifyShape2.Left = selectedShape.Left;
            //magnifyShape2.Top = selectedShape.Top;
            //magnifyShape2.Width = selectedShape.Width;
            //magnifyShape2.Height = selectedShape.Height;

            //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            //duplicatePic.Select();
            //magnifyShape2.Select(Office.MsoTriState.msoFalse);
            //PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            //PowerPoint.Shape groupShape = selection.Group();

            //groupShape.Width *= (finalWidth / initialWidth);
            //groupShape.Height *= (finalHeight / initialHeight);
            //groupShape.Ungroup();
            //duplicatePic.Left += (magnifyShape.Left - magnifyShape2.Left);
            //duplicatePic.Top += (magnifyShape.Top - magnifyShape2.Top);
            //magnifyShape.Delete();
            //magnifyShape2.Delete();

            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
            return addedSlide;
        }
        private PowerPoint.Slide AddDeMagnifyingSlideWithBackground(PowerPoint.Slide magnifyingSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = magnifyingSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsDeMagnifyingSlide" + GetTimestamp(DateTime.Now);
            addedSlide.MoveTo(magnifyingSlide.SlideIndex + 2);

            PowerPoint.Shape duplicatePic = null;
            foreach (PowerPoint.Shape sh in magnifyingSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                DeleteShapeAnnimations(addedSlide, tmp);
                if (sh.Name.Contains("PPTLabsMagnifyAreaSlide"))
                {
                    duplicatePic = tmp;
                }
                if (sh.Name.Contains("PPIndicator") || sh.Name.Contains("PPTLabsMagnifyAreaGroup"))
                {
                    tmp.Delete();
                }
            }

            selectedShape.Copy();
            PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (magnifyShape.Width > magnifyShape.Height)
                magnifyShape.Width = presentation.PageSetup.SlideWidth;
            else
                magnifyShape.Height = presentation.PageSetup.SlideHeight;

            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

            float finalWidthMagnify = magnifyShape.Width;
            float initialWidthMagnify = selectedShape.Width;
            float finalHeightMagnify = magnifyShape.Height;
            float initialHeightMagnify = selectedShape.Height;

            selectedShape.Copy();
            PowerPoint.Shape magnifyShape2 = addedSlide.Shapes.Paste()[1];
            magnifyShape2.Left = selectedShape.Left;
            magnifyShape2.Top = selectedShape.Top;
            magnifyShape2.Width = selectedShape.Width;
            magnifyShape2.Height = selectedShape.Height;

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            duplicatePic.Select();
            magnifyShape2.Select(Office.MsoTriState.msoFalse);
            PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            PowerPoint.Shape groupShape = selection.Group();

            groupShape.Width *= (finalWidthMagnify / initialWidthMagnify);
            groupShape.Height *= (finalHeightMagnify / initialHeightMagnify);
            groupShape.Ungroup();
            duplicatePic.Left += (magnifyShape.Left - magnifyShape2.Left);
            duplicatePic.Top += (magnifyShape.Top - magnifyShape2.Top);
            magnifyShape.Delete();
            magnifyShape2.Delete();

            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectFade = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            float finalWidth = presentation.PageSetup.SlideWidth;
            float initialWidth = duplicatePic.Width;
            float finalHeight = presentation.PageSetup.SlideHeight;
            float initialHeight = duplicatePic.Height;

            float finalX = presentation.PageSetup.SlideWidth / 2;
            float initialX = (duplicatePic.Left + (duplicatePic.Width) / 2);
            float finalY = presentation.PageSetup.SlideHeight / 2;
            float initialY = (duplicatePic.Top + (duplicatePic.Height) / 2);

            int numFrames = 10;

            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
            //float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
            float incrementLeft = (finalX - initialX) / numFrames;
            float incrementTop = (finalY - initialY) / numFrames;

            PowerPoint.Shape lastShape = duplicatePic;
            for (int i = 1; i <= numFrames; i++)
            {
                PowerPoint.Shape dupShape = duplicatePic.Duplicate()[1];
                if (i != 1)
                    sequence[sequence.Count].Delete();

                dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                dupShape.Left = duplicatePic.Left;
                dupShape.Top = duplicatePic.Top;
                //dupShape.Rotation = groupShape.Rotation;

                if (incrementWidth != 0.0f)
                {
                    dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                }

                if (incrementHeight != 0.0f)
                {
                    dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                }

                //if (incrementRotation != 0.0f)
                //{
                //    dupShape.Rotation += (incrementRotation * i);
                //}

                if (incrementLeft != 0.0f)
                {
                    dupShape.Left += (incrementLeft * i);
                }

                if (incrementTop != 0.0f)
                {
                    dupShape.Top += (incrementTop * i);
                }

                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                //appear.Timing.Duration = 0.005f;
                appear.Timing.TriggerDelayTime = ((0.5f / numFrames) * i);

                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                disappear.Exit = Office.MsoTriState.msoTrue;
                //disappear.Timing.Duration = 0.005f;
                disappear.Timing.TriggerDelayTime = ((0.5f / numFrames) * i);

                lastShape = dupShape;
            }

            int j = 0;
            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(duplicatePic) || tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyShape")) && !(tmp.Name.Contains("PPTLabsMagnifyArea")))
                {
                    tmp.Visible = Office.MsoTriState.msoTrue;
                    if (j == 0)
                    {
                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    }
                    else
                    {
                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    }
                    effectFade.Timing.Duration = 0.01f;
                    j++;
                }
            }
            effectFade = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectFade.Exit = Office.MsoTriState.msoTrue;
            effectFade.Timing.Duration = 0.01f;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;

            return addedSlide;
        }
        private PowerPoint.Slide AddMagnifyingSlide(PowerPoint.Slide currentSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = currentSlide.Duplicate()[1];
            MoveMotionAnimation(addedSlide);
            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (HasExitAnimation(addedSlide, tmp))
                    tmp.Delete();
            }
            addedSlide.Name = "PPTLabsMagnifyingSlide" + GetTimestamp(DateTime.Now);

            float centerX = selectedShape.Left + selectedShape.Width / 2;
            float centerY = selectedShape.Top + selectedShape.Height / 2;
            float width = selectedShape.Width;
            float height = selectedShape.Height;

            PowerPoint.Shape identicalShape = FindIdenticalShape(addedSlide, selectedShape);
            if (identicalShape != null)
            {
                identicalShape.Delete();
            }

            PowerPoint.Slide tempSlide = addedSlide.Duplicate()[1];
            addedSlide.Copy();
            PowerPoint.Shape duplicatePic = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];

            float scaleFactorX = presentation.PageSetup.SlideWidth / duplicatePic.Width;
            float scaleFactorY = presentation.PageSetup.SlideHeight / duplicatePic.Height;

            duplicatePic.LockAspectRatio = Office.MsoTriState.msoFalse;
            duplicatePic.Left = 0;
            duplicatePic.Top = 0;
            duplicatePic.Width = presentation.PageSetup.SlideWidth;
            duplicatePic.Height = presentation.PageSetup.SlideHeight;
            duplicatePic.Name = "PPTLabsMagnifyAreaSlide" + GetTimestamp(DateTime.Now);

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);

            PowerPoint.Shape cropShape = tempSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, presentation.PageSetup.SlideWidth - 1, presentation.PageSetup.SlideHeight - 1);
            cropShape.Select();

            //PowerPoint.Shape tempDuplicate = duplicatePic.Duplicate()[1];
            //tempDuplicate.Left = 0;
            //tempDuplicate.Top = 0;
            //tempDuplicate.Select();
            //foreach (PowerPoint.Shape sh in tempSlide.Shapes)
            //{

            //    if (sh.Visible == Office.MsoTriState.msoTrue)
            //    {
            //        PowerPoint.Shape dupShape = null;
            //        sh.Copy();
            //        if (sh.Type == Office.MsoShapeType.msoPlaceholder)
            //        {
            //            dupShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            //        }
            //        else
            //        {
            //            dupShape = addedSlide.Shapes.Paste()[1];
            //        }

            //        dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            //        dupShape.Width = sh.Width;
            //        dupShape.Height = sh.Height;
            //        dupShape.Left = sh.Left;
            //        dupShape.Top = sh.Top;
            //        dupShape.Select(Office.MsoTriState.msoFalse);
            //    }
            //}

            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape croppedShape = CropShapeToSlide(ref sel);

            //tempSlide.Delete();
            croppedShape.Cut();

            PowerPoint.Shape duplicatePic2 = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            duplicatePic2.LockAspectRatio = Office.MsoTriState.msoFalse;
            duplicatePic2.Left = 0;
            duplicatePic2.Top = 0;
            duplicatePic2.Width = presentation.PageSetup.SlideWidth;
            duplicatePic2.Height = presentation.PageSetup.SlideHeight;
            duplicatePic2.Name = "PPTLabsMagnifyAreaGroup" + GetTimestamp(DateTime.Now);

            duplicatePic.PictureFormat.CropLeft += selectedShape.Left / scaleFactorX;
            duplicatePic.PictureFormat.CropTop += selectedShape.Top / scaleFactorY;
            duplicatePic.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width)) / scaleFactorX;
            duplicatePic.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height)) / scaleFactorY;

            duplicatePic2.PictureFormat.CropLeft += selectedShape.Left;
            duplicatePic2.PictureFormat.CropTop += selectedShape.Top;
            duplicatePic2.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width));
            duplicatePic2.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height));

            //duplicatePic.Cut();

            //currentSlide.Duplicate();
            //Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.CustomLayouts[7]);
            //PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);

            //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            //PowerPoint.Shape sh = addedSlide.Shapes.Paste()[1];
            duplicatePic.Left = centerX - (duplicatePic.Width / 2);
            duplicatePic.Top = centerY - (duplicatePic.Height / 2);
            duplicatePic2.Left = centerX - (duplicatePic.Width / 2);
            duplicatePic2.Top = centerY - (duplicatePic.Height / 2);
            duplicatePic2.Visible = Office.MsoTriState.msoFalse;
            tempSlide.Delete();

            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectFade = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;

            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(duplicatePic) || tmp.Equals(indicatorShape) || tmp.Equals(duplicatePic2)))
                {
                    DeleteShapeAnnimations(addedSlide, tmp);
                    effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.25f;
                }
            }

            duplicatePic.Copy();
            PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (magnifyShape.Width > magnifyShape.Height)
            {
                magnifyShape.Width = presentation.PageSetup.SlideWidth;
            }
            else
            {
                magnifyShape.Height = presentation.PageSetup.SlideHeight;
            }

            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

            float finalX = (magnifyShape.Left + (magnifyShape.Width) / 2);
            float initialX = (duplicatePic.Left + (duplicatePic.Width) / 2);
            float finalY = (magnifyShape.Top + (magnifyShape.Height) / 2);
            float initialY = (duplicatePic.Top + (duplicatePic.Height) / 2);

            float finalWidth = magnifyShape.Width;
            float initialWidth = duplicatePic.Width;
            float finalHeight = magnifyShape.Height;
            float initialHeight = duplicatePic.Height;

            magnifyShape.Delete();

            effectMotion = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
            effectMotion.Timing.Duration = 0.5f;
            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

            effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
            effectResize.Timing.Duration = 0.5f;
            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceTime = 0;
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape tmp in addedSlide.NotesPage.Shapes)
                {
                    if (tmp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        tmp.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (tmp.Type == Office.MsoShapeType.msoMedia)
                    tmp.Delete();
            }
            return addedSlide;
        }
        private PowerPoint.Slide AddMagnifiedSlide(PowerPoint.Slide magnifyingSlide)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = magnifyingSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsMagnifiedSlide" + GetTimestamp(DateTime.Now);
            PowerPoint.Shape magnifyShape = null;
            foreach (PowerPoint.Shape sh in magnifyingSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (!sh.Name.Contains("PPTLabsMagnifyAreaGroup"))
                {
                    tmp.Delete();
                }
                else
                {
                    magnifyShape = tmp;
                }
            }

            magnifyShape.Visible = Office.MsoTriState.msoTrue;
            DeleteShapeAnnimations(addedSlide, magnifyShape);
            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            //magnifyShape.Left = 0;
            //magnifyShape.Top = 0;
            if (magnifyShape.Width > magnifyShape.Height)
                magnifyShape.Width = presentation.PageSetup.SlideWidth;
            else
                magnifyShape.Height = presentation.PageSetup.SlideHeight;

            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);
            magnifyShape.PictureFormat.CropLeft = 0;
            magnifyShape.PictureFormat.CropTop = 0;
            magnifyShape.PictureFormat.CropRight = 0;
            magnifyShape.PictureFormat.CropBottom = 0;

            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
            return addedSlide;
        }
        private PowerPoint.Slide AddDeMagnifyingSlide(PowerPoint.Slide magnifyingSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = magnifyingSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsDeMagnifyingSlide" + GetTimestamp(DateTime.Now);
            addedSlide.MoveTo(magnifyingSlide.SlideIndex + 2);

            PowerPoint.Shape magnifyShape = null;
            foreach (PowerPoint.Shape sh in magnifyingSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                DeleteShapeAnnimations(addedSlide, tmp);
                if (sh.Name.Contains("PPTLabsMagnifyAreaSlide"))
                {
                    magnifyShape = tmp;
                }
                if (sh.Name.Contains("PPIndicator") || sh.Name.Contains("PPTLabsMagnifyAreaGroup"))
                {
                    tmp.Delete();
                }
            }

            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            if (magnifyShape.Width > magnifyShape.Height)
                magnifyShape.Width = presentation.PageSetup.SlideWidth;
            else
                magnifyShape.Height = presentation.PageSetup.SlideHeight;

            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectFade = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            float finalX = (selectedShape.Left + (selectedShape.Width) / 2);
            float initialX = (magnifyShape.Left + (magnifyShape.Width) / 2);
            float finalY = (selectedShape.Top + (selectedShape.Height) / 2);
            float initialY = (magnifyShape.Top + (magnifyShape.Height) / 2);

            float finalWidth = selectedShape.Width;
            float initialWidth = magnifyShape.Width;
            float finalHeight = selectedShape.Height;
            float initialHeight = magnifyShape.Height;

            effectMotion = sequence.AddEffect(magnifyShape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
            effectMotion.Timing.Duration = 0.5f;
            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

            effectResize = sequence.AddEffect(magnifyShape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
            effectResize.Timing.Duration = 0.5f;
            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

            int i = 0;
            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(magnifyShape) || tmp.Equals(indicatorShape)))
                {
                    if (i == 0)
                    {
                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    }
                    else
                    {
                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    }
                    effectFade.Timing.Duration = 0.25f;
                    i++;
                }
            }
            effectFade = sequence.AddEffect(magnifyShape, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectFade.Exit = Office.MsoTriState.msoTrue;
            effectFade.Timing.Duration = 0.25f;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;

            return addedSlide;
        }
        private PowerPoint.Slide AddMagnifiedPanSlide(PowerPoint.Slide slideToPanFrom, PowerPoint.Slide slideToPanTo, PowerPoint.Shape shapeFrom)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = slideToPanFrom.Duplicate()[1];
            addedSlide.Name = "PPTLabsMagnifiedPanSlide" + GetTimestamp(DateTime.Now);

            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectFade = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;

            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            PowerPoint.Shape magnifyShapeFrom = null;
            PowerPoint.Shape magnifyShapeTo = null;

            foreach (PowerPoint.Shape sh in slideToPanFrom.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (!sh.Name.Contains("PPTLabsMagnifyAreaGroup"))
                {
                    tmp.Delete();
                }
                else
                {
                    magnifyShapeFrom = tmp;
                }
            }

            foreach (PowerPoint.Shape sh in slideToPanTo.Shapes)
            {
                if (sh.Name.Contains("PPTLabsMagnifyAreaGroup"))
                {
                    magnifyShapeTo = sh;
                }
            }

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(magnifyShapeFrom) || tmp.Equals(indicatorShape)))
                {
                    DeleteShapeAnnimations(addedSlide, tmp);
                    effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.25f;
                }
            }

            //duplicatePic.Copy();
            //PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
            //magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            //if (magnifyShape.Width > magnifyShape.Height)
            //{
            //    magnifyShape.Width = presentation.PageSetup.SlideWidth;
            //}
            //else
            //{
            //    magnifyShape.Height = presentation.PageSetup.SlideHeight;
            //}

            //magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
            //magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);
            float scaleFactorX = presentation.PageSetup.SlideWidth / shapeFrom.Width;
            float scaleFactorY = presentation.PageSetup.SlideHeight / shapeFrom.Height;

            float finalX = (magnifyShapeTo.Left + (magnifyShapeTo.Width) / 2);
            float initialX = (magnifyShapeFrom.Left + (magnifyShapeFrom.Width) / 2);
            float finalY = (magnifyShapeTo.Top + (magnifyShapeTo.Height) / 2);
            float initialY = (magnifyShapeFrom.Top + (magnifyShapeFrom.Height) / 2);

            float finalWidth = magnifyShapeTo.Width;
            float initialWidth = magnifyShapeFrom.Width;
            float finalHeight = magnifyShapeTo.Height;
            float initialHeight = magnifyShapeFrom.Height;

            //magnifyShape.Delete();

            //effectMotion = sequence.AddEffect(magnifyShapeFrom, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            //PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
            //effectMotion.Timing.Duration = 3.0f;
            //float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            //float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            //float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            //float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            //float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
            //float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

            //motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
            //effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
            //effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

            //effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            //PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
            //effectResize.Timing.Duration = 0.5f;
            //resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
            //resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

            int numFrames = 10;

            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
            //float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
            float incrementLeft = (finalX - initialX) / numFrames;
            float incrementTop = (finalY - initialY) / numFrames;

            PowerPoint.Shape lastShape = magnifyShapeFrom;
            for (int i = 1; i <= numFrames; i++)
            {
                PowerPoint.Shape dupShape = magnifyShapeFrom.Duplicate()[1];
                if (i != 1)
                    sequence[sequence.Count].Delete();

                dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                dupShape.Left = magnifyShapeFrom.Left;
                dupShape.Top = magnifyShapeFrom.Top;
                //dupShape.Rotation = groupShape.Rotation;

                if (incrementWidth != 0.0f)
                {
                    dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                }

                if (incrementHeight != 0.0f)
                {
                    dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                }

                //if (incrementRotation != 0.0f)
                //{
                //    dupShape.Rotation += (incrementRotation * i);
                //}

                if (incrementLeft != 0.0f)
                {
                    dupShape.Left += (incrementLeft * i);
                }

                if (incrementTop != 0.0f)
                {
                    dupShape.Top += (incrementTop * i);
                }

                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                //appear.Timing.Duration = 0.005f;
                appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                disappear.Exit = Office.MsoTriState.msoTrue;
                //disappear.Timing.Duration = 0.005f;
                disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                lastShape = dupShape;
            }

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceTime = 0;
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape tmp in addedSlide.NotesPage.Shapes)
                {
                    if (tmp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        tmp.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (tmp.Type == Office.MsoShapeType.msoMedia)
                    tmp.Delete();
            }
            return addedSlide;

        }
        public void MultiSlideZoomToArea(PowerPoint.Slide currentSlide, List<PowerPoint.Shape> selectedShapes)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide magnifyingSlide = null;
                PowerPoint.Slide magnifiedSlide = null;
                PowerPoint.Slide magnifiedPanSlide = null;
                PowerPoint.Slide demagnifyingSlide = null;
                PowerPoint.Slide lastMagnifiedSlide = null;
                int count = 1;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    if (!backgroundZoomChecked)
                    {
                        magnifyingSlide = AddMagnifyingSlide(currentSlide, selectedShape);
                        magnifiedSlide = AddMagnifiedSlide(magnifyingSlide);
                        if (count != 1)
                            magnifiedPanSlide = AddMagnifiedPanSlide(lastMagnifiedSlide, magnifiedSlide, selectedShape);
                        if (count == selectedShapes.Count)
                        {
                            demagnifyingSlide = AddDeMagnifyingSlide(magnifyingSlide, selectedShape);
                        }
                    }
                    else
                    {
                        magnifyingSlide = AddMagnifyingSlideWithBackground(currentSlide, selectedShape);
                        magnifiedSlide = AddMagnifiedSlideWithBackground(magnifyingSlide, selectedShape);
                        if (count != 1)
                            magnifiedPanSlide = AddMagnifiedPanSlide(lastMagnifiedSlide, magnifiedSlide, selectedShape);
                        if (count == selectedShapes.Count)
                        {
                            demagnifyingSlide = AddDeMagnifyingSlideWithBackground(magnifyingSlide, selectedShape);
                        }
                    }
                    selectedShape.Delete();
                    if (count != 1)
                    {
                        magnifyingSlide.Delete();
                        magnifiedSlide.MoveTo(magnifiedPanSlide.SlideIndex);
                        if (demagnifyingSlide != null)
                            demagnifyingSlide.MoveTo(magnifiedSlide.SlideIndex);
                        lastMagnifiedSlide = magnifiedSlide;
                    }
                    else
                    {
                        lastMagnifiedSlide = magnifiedSlide;
                    }

                    count++;
                }
            }
            catch (Exception e)
            {
                LogException(e, "MultiSlideZoomToArea");
                throw;
            }
        }
        private PowerPoint.Slide AddZoomToAreaSlide(PowerPoint.Slide currentSlide, List<PowerPoint.Shape> selectedShapes)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Sequence sequence = null;

            PowerPoint.Slide addedSlide = currentSlide.Duplicate()[1];
            MoveMotionAnimation(addedSlide);
            foreach (PowerPoint.Shape sh in currentSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (HasExitAnimation(addedSlide, tmp))
                    tmp.Delete();
            }
            addedSlide.Name = "PPTLabsMagnifyingSlide" + GetTimestamp(DateTime.Now);


            foreach (PowerPoint.Shape sh in selectedShapes)
            {
                PowerPoint.Shape identicalShape = FindIdenticalShape(addedSlide, sh);
                if (identicalShape != null)
                {
                    identicalShape.Delete();
                }
            }

            PowerPoint.Slide tmpSlide = addedSlide.Duplicate()[1];
            addedSlide.Copy();
            PowerPoint.Shape magnifySlide = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            magnifySlide.Name = "PPTLabsZoomSlide" + GetTimestamp(DateTime.Now);

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tmpSlide.SlideIndex);

            PowerPoint.Shape cropShape = tmpSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, presentation.PageSetup.SlideWidth - 1, presentation.PageSetup.SlideHeight - 1);
            cropShape.Select();
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape croppedShape = CropShapeToSlide(ref sel);
            croppedShape.Cut();

            PowerPoint.Shape magnifyGroup = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            magnifyGroup.Name = "PPTLabsZoomGroup" + GetTimestamp(DateTime.Now);
            magnifyGroup.LockAspectRatio = Office.MsoTriState.msoFalse;
            magnifyGroup.Left = 0;
            magnifyGroup.Top = 0;
            magnifyGroup.Width = presentation.PageSetup.SlideWidth;
            magnifyGroup.Height = presentation.PageSetup.SlideHeight;
            tmpSlide.Delete();
            sequence = addedSlide.TimeLine.MainSequence;

            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Timing.Duration = 0;

            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(magnifySlide) || tmp.Equals(indicatorShape) || tmp.Equals(magnifyGroup)))
                {
                    DeleteShapeAnnimations(addedSlide, tmp);
                    tmp.Visible = Office.MsoTriState.msoFalse;
                }
            }
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape tmp in addedSlide.NotesPage.Shapes)
                {
                    if (tmp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        tmp.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (tmp.Type == Office.MsoShapeType.msoMedia)
                    tmp.Delete();
            }

            return addedSlide;
        }

        private void SingleSlideZoomToArea(PowerPoint.Slide currentSlide, List<PowerPoint.Shape> selectedShapes)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                int count = 1;
                PowerPoint.Shape lastMagnifiedShape = null;
                PowerPoint.Effect effectDisappear = null;
                PowerPoint.Effect effectAppear = null;
                PowerPoint.Effect effectFade = null;
                PowerPoint.Effect effectMotion = null;
                PowerPoint.Effect effectResize = null;
                PowerPoint.Shape indicatorShape = null;
                PowerPoint.Shape magnifySlide = null;
                PowerPoint.Shape magnifyGroup = null;

                PowerPoint.Slide addedSlide = AddZoomToAreaSlide(currentSlide, selectedShapes);
                PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
                foreach (PowerPoint.Shape sh in addedSlide.Shapes)
                {
                    if (sh.Name.Contains("PPIndicator"))
                        indicatorShape = sh;
                    else if (sh.Name.Contains("PPTLabsZoomSlide"))
                        magnifySlide = sh;
                    else if (sh.Name.Contains("PPTLabsZoomGroup"))
                        magnifyGroup = sh;
                }

                float scaleFactorX = presentation.PageSetup.SlideWidth / magnifySlide.Width;
                float scaleFactorY = presentation.PageSetup.SlideHeight / magnifySlide.Height;

                magnifySlide.LockAspectRatio = Office.MsoTriState.msoFalse;
                magnifySlide.Left = 0;
                magnifySlide.Top = 0;
                magnifySlide.Width = presentation.PageSetup.SlideWidth;
                magnifySlide.Height = presentation.PageSetup.SlideHeight;

                foreach (PowerPoint.Shape selectedShape in selectedShapes)
                {
                    if (!backgroundZoomChecked)
                    {
                        if (count == 1)
                        {
                            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
                            {
                                if (!(tmp.Equals(magnifySlide) || tmp.Equals(indicatorShape) || tmp.Equals(magnifyGroup)))
                                {
                                    tmp.Visible = Office.MsoTriState.msoTrue;
                                    effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    effectFade.Exit = Office.MsoTriState.msoTrue;
                                    effectFade.Timing.Duration = 0.25f;
                                }
                            }

                            PowerPoint.Shape duplicatePic = magnifySlide.Duplicate()[1];
                            duplicatePic.Name = "PPTLabsMagnifyingAreaSlide" + GetTimestamp(DateTime.Now);
                            duplicatePic.Left = magnifySlide.Left;
                            duplicatePic.Top = magnifySlide.Top;

                            duplicatePic.PictureFormat.CropLeft += selectedShape.Left / scaleFactorX;
                            duplicatePic.PictureFormat.CropTop += selectedShape.Top / scaleFactorY;
                            duplicatePic.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width)) / scaleFactorX;
                            duplicatePic.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height)) / scaleFactorY;

                            duplicatePic.Left = selectedShape.Left + selectedShape.Width / 2 - (duplicatePic.Width / 2);
                            duplicatePic.Top = selectedShape.Top + selectedShape.Height / 2 - (duplicatePic.Height / 2);

                            duplicatePic.Copy();

                            PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
                            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                            if (magnifyShape.Width > magnifyShape.Height)
                            {
                                magnifyShape.Width = presentation.PageSetup.SlideWidth;
                            }
                            else
                            {
                                magnifyShape.Height = presentation.PageSetup.SlideHeight;
                            }

                            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
                            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

                            float finalWidth = magnifyShape.Width;
                            float initialWidth = duplicatePic.Width;
                            float finalHeight = magnifyShape.Height;
                            float initialHeight = duplicatePic.Height;

                            float finalX = (magnifyShape.Left + (magnifyShape.Width) / 2);
                            float initialX = (duplicatePic.Left + (duplicatePic.Width) / 2);
                            float finalY = (magnifyShape.Top + (magnifyShape.Height) / 2);
                            float initialY = (duplicatePic.Top + (duplicatePic.Height) / 2);

                            magnifyShape.Delete();

                            effectMotion = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                            effectMotion.Timing.Duration = 0.5f;
                            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

                            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

                            effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                            effectResize.Timing.Duration = 0.5f;
                            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                            PowerPoint.Shape duplicatePic2 = magnifyGroup.Duplicate()[1];
                            duplicatePic2.Name = "PPTLabsMagnifyAreaGroupShape" + count + "-" + GetTimestamp(DateTime.Now);

                            duplicatePic2.PictureFormat.CropLeft += selectedShape.Left;
                            duplicatePic2.PictureFormat.CropTop += selectedShape.Top;
                            duplicatePic2.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width));
                            duplicatePic2.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height));
                            DeleteShapeAnnimations(addedSlide, duplicatePic2);
                            duplicatePic2.LockAspectRatio = Office.MsoTriState.msoTrue;

                            if (duplicatePic2.Width > duplicatePic2.Height)
                                duplicatePic2.Width = presentation.PageSetup.SlideWidth;
                            else
                                duplicatePic2.Height = presentation.PageSetup.SlideHeight;

                            duplicatePic2.Left = (presentation.PageSetup.SlideWidth / 2) - (duplicatePic2.Width / 2);
                            duplicatePic2.Top = (presentation.PageSetup.SlideHeight / 2) - (duplicatePic2.Height / 2);
                            duplicatePic2.PictureFormat.CropLeft = 0;
                            duplicatePic2.PictureFormat.CropTop = 0;
                            duplicatePic2.PictureFormat.CropRight = 0;
                            duplicatePic2.PictureFormat.CropBottom = 0;

                            effectDisappear = sequence.AddEffect(duplicatePic2, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            effectDisappear.Timing.Duration = 0;

                            effectDisappear = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectDisappear.Exit = Office.MsoTriState.msoTrue;
                            effectDisappear.Timing.Duration = 0;
                            lastMagnifiedShape = duplicatePic2;
                        }

                        if (count != 1)
                        {
                            PowerPoint.Shape duplicatePic2 = magnifyGroup.Duplicate()[1];
                            duplicatePic2.Name = "PPTLabsMagnifyAreaGroupShape" + count + "-" + GetTimestamp(DateTime.Now);

                            duplicatePic2.PictureFormat.CropLeft += selectedShape.Left;
                            duplicatePic2.PictureFormat.CropTop += selectedShape.Top;
                            duplicatePic2.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width));
                            duplicatePic2.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height));
                            DeleteShapeAnnimations(addedSlide, duplicatePic2);
                            duplicatePic2.LockAspectRatio = Office.MsoTriState.msoTrue;

                            if (duplicatePic2.Width > duplicatePic2.Height)
                                duplicatePic2.Width = presentation.PageSetup.SlideWidth;
                            else
                                duplicatePic2.Height = presentation.PageSetup.SlideHeight;

                            duplicatePic2.Left = (presentation.PageSetup.SlideWidth / 2) - (duplicatePic2.Width / 2);
                            duplicatePic2.Top = (presentation.PageSetup.SlideHeight / 2) - (duplicatePic2.Height / 2);
                            duplicatePic2.PictureFormat.CropLeft = 0;
                            duplicatePic2.PictureFormat.CropTop = 0;
                            duplicatePic2.PictureFormat.CropRight = 0;
                            duplicatePic2.PictureFormat.CropBottom = 0;

                            PowerPoint.Shape magnifyShapeFrom = lastMagnifiedShape;
                            PowerPoint.Shape magnifyShapeTo = duplicatePic2;

                            float finalX = (magnifyShapeTo.Left + (magnifyShapeTo.Width) / 2);
                            float initialX = (magnifyShapeFrom.Left + (magnifyShapeFrom.Width) / 2);
                            float finalY = (magnifyShapeTo.Top + (magnifyShapeTo.Height) / 2);
                            float initialY = (magnifyShapeFrom.Top + (magnifyShapeFrom.Height) / 2);

                            float finalWidth = magnifyShapeTo.Width;
                            float initialWidth = magnifyShapeFrom.Width;
                            float finalHeight = magnifyShapeTo.Height;
                            float initialHeight = magnifyShapeFrom.Height;

                            int numFrames = 10;

                            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                            //float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                            float incrementLeft = (finalX - initialX) / numFrames;
                            float incrementTop = (finalY - initialY) / numFrames;

                            PowerPoint.Shape lastShape = magnifyShapeFrom;
                            PowerPoint.MsoAnimTriggerType trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                            for (int i = 1; i <= numFrames; i++)
                            {
                                PowerPoint.Shape dupShape = magnifyShapeFrom.Duplicate()[1];
                                dupShape.Name = "PPTLabsMagnifyPanAreaGroup" + GetTimestamp(DateTime.Now);
                                DeleteShapeAnnimations(addedSlide, dupShape);

                                if (i != 1)
                                {
                                    //sequence[sequence.Count].Delete();
                                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                                }
                                else
                                {
                                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                                }


                                dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                                dupShape.Left = magnifyShapeFrom.Left;
                                dupShape.Top = magnifyShapeFrom.Top;
                                //dupShape.Rotation = groupShape.Rotation;

                                if (incrementWidth != 0.0f)
                                {
                                    dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                }

                                if (incrementHeight != 0.0f)
                                {
                                    dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                }

                                //if (incrementRotation != 0.0f)
                                //{
                                //    dupShape.Rotation += (incrementRotation * i);
                                //}

                                if (incrementLeft != 0.0f)
                                {
                                    dupShape.Left += (incrementLeft * i);
                                }

                                if (incrementTop != 0.0f)
                                {
                                    dupShape.Top += (incrementTop * i);
                                }

                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                                //appear.Timing.Duration = 0.005f;
                                appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                disappear.Exit = Office.MsoTriState.msoTrue;
                                //disappear.Timing.Duration = 0.005f;
                                disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                                lastShape = dupShape;
                            }
                            effectDisappear = sequence.AddEffect(duplicatePic2, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            effectDisappear.Timing.Duration = 0;
                            lastMagnifiedShape = duplicatePic2;

                            effectFade = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectFade.Exit = Office.MsoTriState.msoTrue;
                            effectFade.Timing.Duration = 0.01f;
                        }
                        if (count == selectedShapes.Count)
                        {
                            PowerPoint.Shape duplicatePic = magnifySlide.Duplicate()[1];
                            duplicatePic.Name = "PPTLabsDeMagnifyAreaSlide" + GetTimestamp(DateTime.Now);
                            duplicatePic.Left = magnifySlide.Left;
                            duplicatePic.Top = magnifySlide.Top;

                            duplicatePic.PictureFormat.CropLeft += selectedShape.Left / scaleFactorX;
                            duplicatePic.PictureFormat.CropTop += selectedShape.Top / scaleFactorY;
                            duplicatePic.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width)) / scaleFactorX;
                            duplicatePic.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height)) / scaleFactorY;

                            duplicatePic.Left = selectedShape.Left + selectedShape.Width / 2 - (duplicatePic.Width / 2);
                            duplicatePic.Top = selectedShape.Top + selectedShape.Height / 2 - (duplicatePic.Height / 2);

                            duplicatePic.LockAspectRatio = Office.MsoTriState.msoTrue;
                            if (duplicatePic.Width > duplicatePic.Height)
                                duplicatePic.Width = presentation.PageSetup.SlideWidth;
                            else
                                duplicatePic.Height = presentation.PageSetup.SlideHeight;

                            duplicatePic.Left = (presentation.PageSetup.SlideWidth / 2) - (duplicatePic.Width / 2);
                            duplicatePic.Top = (presentation.PageSetup.SlideHeight / 2) - (duplicatePic.Height / 2);

                            float finalX = (selectedShape.Left + (selectedShape.Width) / 2);
                            float initialX = (duplicatePic.Left + (duplicatePic.Width) / 2);
                            float finalY = (selectedShape.Top + (selectedShape.Height) / 2);
                            float initialY = (duplicatePic.Top + (duplicatePic.Height) / 2);

                            float finalWidth = selectedShape.Width;
                            float initialWidth = duplicatePic.Width;
                            float finalHeight = selectedShape.Height;
                            float initialHeight = duplicatePic.Height;

                            effectDisappear = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                            effectDisappear.Timing.Duration = 0;

                            effectDisappear = sequence.AddEffect(lastMagnifiedShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectDisappear.Timing.Duration = 0;
                            effectDisappear.Exit = Office.MsoTriState.msoTrue;

                            effectMotion = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                            effectMotion.Timing.Duration = 0.5f;
                            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

                            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

                            effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                            effectResize.Timing.Duration = 0.5f;
                            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                            int j = 0;
                            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
                            {
                                if (!(tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyAreaGroup")) && !(tmp.Name.Contains("PPTLabsMagnifyPanAreaGroup")) && !(tmp.Name.Contains("PPTLabsDeMagnifyAreaSlide")) && !(tmp.Name.Contains("PPTLabsMagnifyingAreaSlide")))
                                {
                                    tmp.Visible = Office.MsoTriState.msoTrue;
                                    if (j == 0)
                                    {
                                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                                    }
                                    else
                                    {
                                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    }
                                    effectFade.Timing.Duration = 0.25f;
                                    j++;
                                }
                            }
                            j = 0;
                            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
                            {
                                if (tmp.Name.Contains("PPTLabsMagnifyAreaGroup") || tmp.Name.Contains("PPTLabsMagnifyingAreaSlide") || tmp.Name.Contains("PPTLabsDeMagnifyAreaSlide"))
                                {
                                    if (tmp.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        if (j == 0)
                                        {
                                            effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                                        }
                                        else
                                        {
                                            effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                        }
                                        effectFade.Exit = Office.MsoTriState.msoTrue;
                                        effectFade.Timing.Duration = 0.01f;
                                        j++;
                                    }
                                }
                            }

                        }
                    }
                    else
                    {
                        if (count == 1)
                        {
                            PowerPoint.Shape duplicatePic = magnifySlide.Duplicate()[1];
                            duplicatePic.Name = "PPTLabsMagnifyingAreaSlide" + GetTimestamp(DateTime.Now);

                            duplicatePic.LockAspectRatio = Office.MsoTriState.msoFalse;
                            duplicatePic.Left = 0;
                            duplicatePic.Top = 0;
                            duplicatePic.Width = presentation.PageSetup.SlideWidth;
                            duplicatePic.Height = presentation.PageSetup.SlideHeight;

                            selectedShape.Copy();

                            PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
                            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                            if (magnifyShape.Width > magnifyShape.Height)
                            {
                                magnifyShape.Width = presentation.PageSetup.SlideWidth;
                            }
                            else
                            {
                                magnifyShape.Height = presentation.PageSetup.SlideHeight;
                            }

                            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
                            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

                            float finalWidth = magnifyShape.Width;
                            float initialWidth = selectedShape.Width;
                            float finalHeight = magnifyShape.Height;
                            float initialHeight = selectedShape.Height;

                            float finalX = (magnifyShape.Left + (magnifyShape.Width) / 2) * (finalWidth / initialWidth);
                            float initialX = (selectedShape.Left + (selectedShape.Width) / 2) * (finalWidth / initialWidth);
                            float finalY = (magnifyShape.Top + (magnifyShape.Height) / 2) * (finalHeight / initialHeight);
                            float initialY = (selectedShape.Top + (selectedShape.Height) / 2) * (finalHeight / initialHeight);

                            magnifyShape.Delete();

                            effectMotion = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                            effectMotion.Timing.Duration = 0.5f;
                            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

                            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

                            effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                            effectResize.Timing.Duration = 0.5f;
                            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                            PowerPoint.Shape duplicatePic2 = magnifyGroup.Duplicate()[1];
                            duplicatePic2.Name = "PPTLabsMagnifyAreaGroupShape" + count + "-" + GetTimestamp(DateTime.Now);

                            duplicatePic2.PictureFormat.CropLeft += selectedShape.Left;
                            duplicatePic2.PictureFormat.CropTop += selectedShape.Top;
                            duplicatePic2.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width));
                            duplicatePic2.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height));
                            DeleteShapeAnnimations(addedSlide, duplicatePic2);
                            duplicatePic2.LockAspectRatio = Office.MsoTriState.msoTrue;

                            if (duplicatePic2.Width > duplicatePic2.Height)
                                duplicatePic2.Width = presentation.PageSetup.SlideWidth;
                            else
                                duplicatePic2.Height = presentation.PageSetup.SlideHeight;

                            duplicatePic2.Left = (presentation.PageSetup.SlideWidth / 2) - (duplicatePic2.Width / 2);
                            duplicatePic2.Top = (presentation.PageSetup.SlideHeight / 2) - (duplicatePic2.Height / 2);
                            duplicatePic2.PictureFormat.CropLeft = 0;
                            duplicatePic2.PictureFormat.CropTop = 0;
                            duplicatePic2.PictureFormat.CropRight = 0;
                            duplicatePic2.PictureFormat.CropBottom = 0;

                            effectDisappear = sequence.AddEffect(duplicatePic2, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            effectDisappear.Timing.Duration = 0;

                            effectDisappear = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectDisappear.Exit = Office.MsoTriState.msoTrue;
                            effectDisappear.Timing.Duration = 0;
                            lastMagnifiedShape = duplicatePic2;
                        }

                        if (count != 1)
                        {
                            PowerPoint.Shape duplicatePic2 = magnifyGroup.Duplicate()[1];
                            duplicatePic2.Name = "PPTLabsMagnifyAreaGroupShape" + count + "-" + GetTimestamp(DateTime.Now);

                            duplicatePic2.PictureFormat.CropLeft += selectedShape.Left;
                            duplicatePic2.PictureFormat.CropTop += selectedShape.Top;
                            duplicatePic2.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width));
                            duplicatePic2.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height));
                            DeleteShapeAnnimations(addedSlide, duplicatePic2);
                            duplicatePic2.LockAspectRatio = Office.MsoTriState.msoTrue;

                            if (duplicatePic2.Width > duplicatePic2.Height)
                                duplicatePic2.Width = presentation.PageSetup.SlideWidth;
                            else
                                duplicatePic2.Height = presentation.PageSetup.SlideHeight;

                            duplicatePic2.Left = (presentation.PageSetup.SlideWidth / 2) - (duplicatePic2.Width / 2);
                            duplicatePic2.Top = (presentation.PageSetup.SlideHeight / 2) - (duplicatePic2.Height / 2);
                            duplicatePic2.PictureFormat.CropLeft = 0;
                            duplicatePic2.PictureFormat.CropTop = 0;
                            duplicatePic2.PictureFormat.CropRight = 0;
                            duplicatePic2.PictureFormat.CropBottom = 0;

                            PowerPoint.Shape magnifyShapeFrom = lastMagnifiedShape;
                            PowerPoint.Shape magnifyShapeTo = duplicatePic2;

                            float finalX = (magnifyShapeTo.Left + (magnifyShapeTo.Width) / 2);
                            float initialX = (magnifyShapeFrom.Left + (magnifyShapeFrom.Width) / 2);
                            float finalY = (magnifyShapeTo.Top + (magnifyShapeTo.Height) / 2);
                            float initialY = (magnifyShapeFrom.Top + (magnifyShapeFrom.Height) / 2);

                            float finalWidth = magnifyShapeTo.Width;
                            float initialWidth = magnifyShapeFrom.Width;
                            float finalHeight = magnifyShapeTo.Height;
                            float initialHeight = magnifyShapeFrom.Height;

                            int numFrames = 10;

                            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                            //float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                            float incrementLeft = (finalX - initialX) / numFrames;
                            float incrementTop = (finalY - initialY) / numFrames;

                            PowerPoint.Shape lastShape = magnifyShapeFrom;
                            PowerPoint.MsoAnimTriggerType trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                            for (int i = 1; i <= numFrames; i++)
                            {
                                PowerPoint.Shape dupShape = magnifyShapeFrom.Duplicate()[1];
                                dupShape.Name = "PPTLabsMagnifyPanAreaGroup" + GetTimestamp(DateTime.Now);
                                DeleteShapeAnnimations(addedSlide, dupShape);

                                if (i != 1)
                                {
                                    //sequence[sequence.Count].Delete();
                                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                                }
                                else
                                {
                                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                                }


                                dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                                dupShape.Left = magnifyShapeFrom.Left;
                                dupShape.Top = magnifyShapeFrom.Top;
                                //dupShape.Rotation = groupShape.Rotation;

                                if (incrementWidth != 0.0f)
                                {
                                    dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                }

                                if (incrementHeight != 0.0f)
                                {
                                    dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                }

                                //if (incrementRotation != 0.0f)
                                //{
                                //    dupShape.Rotation += (incrementRotation * i);
                                //}

                                if (incrementLeft != 0.0f)
                                {
                                    dupShape.Left += (incrementLeft * i);
                                }

                                if (incrementTop != 0.0f)
                                {
                                    dupShape.Top += (incrementTop * i);
                                }

                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                                //appear.Timing.Duration = 0.005f;
                                appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                disappear.Exit = Office.MsoTriState.msoTrue;
                                //disappear.Timing.Duration = 0.005f;
                                disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                                lastShape = dupShape;
                            }
                            effectDisappear = sequence.AddEffect(duplicatePic2, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                            effectDisappear.Timing.Duration = 0;
                            lastMagnifiedShape = duplicatePic2;

                            effectFade = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectFade.Exit = Office.MsoTriState.msoTrue;
                            effectFade.Timing.Duration = 0.01f;
                        }

                        if (count == selectedShapes.Count)
                        {
                            PowerPoint.Shape duplicatePic = magnifySlide.Duplicate()[1];
                            duplicatePic.Name = "PPTLabsDeMagnifyAreaSlide" + GetTimestamp(DateTime.Now);
                            duplicatePic.LockAspectRatio = Office.MsoTriState.msoFalse;
                            duplicatePic.Left = 0;
                            duplicatePic.Top = 0;
                            duplicatePic.Width = presentation.PageSetup.SlideWidth;
                            duplicatePic.Height = presentation.PageSetup.SlideHeight;

                            selectedShape.Copy();
                            PowerPoint.Shape magnifyShape = addedSlide.Shapes.Paste()[1];
                            magnifyShape.LockAspectRatio = Office.MsoTriState.msoTrue;
                            if (magnifyShape.Width > magnifyShape.Height)
                                magnifyShape.Width = presentation.PageSetup.SlideWidth;
                            else
                                magnifyShape.Height = presentation.PageSetup.SlideHeight;

                            magnifyShape.Left = (presentation.PageSetup.SlideWidth / 2) - (magnifyShape.Width / 2);
                            magnifyShape.Top = (presentation.PageSetup.SlideHeight / 2) - (magnifyShape.Height / 2);

                            float finalWidthMagnify = magnifyShape.Width;
                            float initialWidthMagnify = selectedShape.Width;
                            float finalHeightMagnify = magnifyShape.Height;
                            float initialHeightMagnify = selectedShape.Height;

                            selectedShape.Copy();
                            PowerPoint.Shape magnifyShape2 = addedSlide.Shapes.Paste()[1];
                            magnifyShape2.Left = selectedShape.Left;
                            magnifyShape2.Top = selectedShape.Top;
                            magnifyShape2.Width = selectedShape.Width;
                            magnifyShape2.Height = selectedShape.Height;

                            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
                            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
                            duplicatePic.Select();
                            magnifyShape2.Select(Office.MsoTriState.msoFalse);
                            PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                            PowerPoint.Shape groupShape = selection.Group();

                            groupShape.Width *= (finalWidthMagnify / initialWidthMagnify);
                            groupShape.Height *= (finalHeightMagnify / initialHeightMagnify);
                            groupShape.Ungroup();
                            duplicatePic.Left += (magnifyShape.Left - magnifyShape2.Left);
                            duplicatePic.Top += (magnifyShape.Top - magnifyShape2.Top);
                            magnifyShape.Delete();
                            magnifyShape2.Delete();

                            effectDisappear = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                            effectDisappear.Timing.Duration = 0;

                            float finalWidth = presentation.PageSetup.SlideWidth;
                            float initialWidth = duplicatePic.Width;
                            float finalHeight = presentation.PageSetup.SlideHeight;
                            float initialHeight = duplicatePic.Height;

                            float finalX = presentation.PageSetup.SlideWidth / 2;
                            float initialX = (duplicatePic.Left + (duplicatePic.Width) / 2);
                            float finalY = presentation.PageSetup.SlideHeight / 2;
                            float initialY = (duplicatePic.Top + (duplicatePic.Height) / 2);

                            int numFrames = 10;

                            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                            //float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                            float incrementLeft = (finalX - initialX) / numFrames;
                            float incrementTop = (finalY - initialY) / numFrames;

                            PowerPoint.Shape lastShape = duplicatePic;
                            for (int i = 1; i <= numFrames; i++)
                            {
                                PowerPoint.Shape dupShape = duplicatePic.Duplicate()[1];
                                DeleteShapeAnnimations(addedSlide, dupShape);
                                //if (i != 1)
                                //    sequence[sequence.Count].Delete();

                                dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                                dupShape.Left = duplicatePic.Left;
                                dupShape.Top = duplicatePic.Top;
                                //dupShape.Rotation = groupShape.Rotation;

                                if (incrementWidth != 0.0f)
                                {
                                    dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                }

                                if (incrementHeight != 0.0f)
                                {
                                    dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                }

                                //if (incrementRotation != 0.0f)
                                //{
                                //    dupShape.Rotation += (incrementRotation * i);
                                //}

                                if (incrementLeft != 0.0f)
                                {
                                    dupShape.Left += (incrementLeft * i);
                                }

                                if (incrementTop != 0.0f)
                                {
                                    dupShape.Top += (incrementTop * i);
                                }

                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                //appear.Timing.Duration = 0.005f;
                                appear.Timing.TriggerDelayTime = ((0.5f / numFrames) * i);

                                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                disappear.Exit = Office.MsoTriState.msoTrue;
                                //disappear.Timing.Duration = 0.005f;
                                disappear.Timing.TriggerDelayTime = ((0.5f / numFrames) * i);

                                lastShape = dupShape;
                            }

                            int j = 0;
                            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
                            {
                                if (!(tmp.Equals(indicatorShape)) && !(tmp.Name.Contains("PPTLabsMagnifyAreaGroup")) && !(tmp.Name.Contains("PPTLabsMagnifyPanAreaGroup")) && !(tmp.Name.Contains("PPTLabsDeMagnifyAreaSlide")) && !(tmp.Name.Contains("PPTLabsMagnifyingAreaSlide")))
                                {
                                    tmp.Visible = Office.MsoTriState.msoTrue;
                                    if (j == 0)
                                    {
                                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                                    }
                                    else
                                    {
                                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    }
                                    effectFade.Timing.Duration = 0.01f;
                                }
                                else if (tmp.Name.Contains("PPTLabsMagnifyAreaGroup") || tmp.Name.Contains("PPTLabsMagnifyingAreaSlide"))
                                {
                                    if (tmp.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        if (j == 0)
                                        {
                                            effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                                        }
                                        else
                                        {
                                            effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                        }
                                        effectFade.Exit = Office.MsoTriState.msoTrue;
                                        effectFade.Timing.Duration = 0.01f;
                                    }
                                }
                                j++;
                            }
                            effectFade = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectFade.Exit = Office.MsoTriState.msoTrue;
                            effectFade.Timing.Duration = 0.01f;
                        }
                    }
                    selectedShape.Delete();
                    indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    count++;
                }
                magnifySlide.Delete();
                magnifyGroup.Delete();
            }
            catch (Exception e)
            {
                LogException(e, "SingleSlideZoomToArea");
                throw;
            }
        }
        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            ZoomToArea.AddZoomToArea();
            //PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            //PowerPoint.Slide currentSlide = GetCurrentSlide();
            //if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count == 0)
            //    return;

            //if (currentSlide.Name.Contains("PPTLabsZoomToAreaSlide") && currentSlide.SlideIndex != presentation.Slides.Count)
            //{
            //    PowerPoint.Slide nextSlide = GetNextSlide(currentSlide);
            //    PowerPoint.Slide tempSlide = null;
            //    while ((nextSlide.Name.Contains("PPTLabsMagnifyingSlide") || (nextSlide.Name.Contains("PPTLabsMagnifiedSlide"))
            //           || (nextSlide.Name.Contains("PPTLabsDeMagnifyingSlide")) || (nextSlide.Name.Contains("PPTLabsMagnifiedPanSlide")))
            //           && nextSlide.SlideIndex < presentation.Slides.Count)
            //    {
            //        tempSlide = nextSlide;
            //        nextSlide = GetNextSlide(tempSlide);
            //        tempSlide.Delete();
            //    }
            //}
            //currentSlide.Name = "PPTLabsZoomToAreaSlide" + GetTimestamp(DateTime.Now);
            //PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            //List<PowerPoint.Shape> editedSelectedShapes = new List<PowerPoint.Shape>();
            //PowerPoint.Effect effectAppear = null;
            //PowerPoint.Effect effectDisappear = null;
            //int count = 1;

            //foreach (PowerPoint.Shape sh in selectedShapes)
            //{
            //    PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
            //    effectAppear = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            //    effectAppear.Timing.Duration = 0;

            //    effectDisappear = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            //    effectDisappear.Exit = Office.MsoTriState.msoTrue;
            //    effectDisappear.Timing.Duration = 0;
            //    sh.Visible = Office.MsoTriState.msoFalse;
            //}
            //foreach (PowerPoint.Shape selectedShape2 in selectedShapes)
            //{
            //    selectedShape2.Visible = Office.MsoTriState.msoTrue;
            //    selectedShape2.Name = "PPTLabsMagnifyShape" + GetTimestamp(DateTime.Now);
            //    selectedShape2.Copy();
            //    if (selectedShape2.HasTextFrame == Office.MsoTriState.msoTrue)
            //    {
            //        selectedShape2.TextFrame2.DeleteText();
            //        selectedShape2.TextFrame2.TextRange.Text = "Zoom Shape " + count;
            //        selectedShape2.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
            //        selectedShape2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xffffff;
            //        selectedShape2.TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            //        selectedShape2.Visible = Office.MsoTriState.msoFalse;
            //    }

            //    PowerPoint.Shape selectedShape = currentSlide.Shapes.Paste()[1];
            //    selectedShape.LockAspectRatio = Office.MsoTriState.msoFalse;

            //    if (selectedShape2.Width > selectedShape2.Height)
            //    {
            //        selectedShape.Width = selectedShape2.Width;
            //        selectedShape.Height = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight * selectedShape.Width / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth;
            //        selectedShape.Left = selectedShape2.Left + (selectedShape2.Width / 2) - (selectedShape.Width / 2);
            //        selectedShape.Top = selectedShape2.Top + (selectedShape2.Height / 2) - (selectedShape.Height / 2);
            //    }
            //    else
            //    {
            //        selectedShape.Height = selectedShape2.Height;
            //        selectedShape.Width = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth * selectedShape.Height / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight;
            //        selectedShape.Left = selectedShape2.Left + (selectedShape2.Width / 2) - (selectedShape.Width / 2);
            //        selectedShape.Top = selectedShape2.Top + (selectedShape2.Height / 2) - (selectedShape.Height / 2);
            //    }

            //    if (selectedShape.Width > presentation.PageSetup.SlideWidth)
            //        selectedShape.Width = presentation.PageSetup.SlideWidth;
            //    if (selectedShape.Height > presentation.PageSetup.SlideHeight)
            //        selectedShape.Height = presentation.PageSetup.SlideHeight;

            //    if (selectedShape.Left < 0)
            //        selectedShape.Left = 0;
            //    if (selectedShape.Left + selectedShape.Width > presentation.PageSetup.SlideWidth)
            //        selectedShape.Left = presentation.PageSetup.SlideWidth - selectedShape.Width;
            //    if (selectedShape.Top < 0)
            //        selectedShape.Top = 0;
            //    if (selectedShape.Top + selectedShape.Height > presentation.PageSetup.SlideHeight)
            //        selectedShape.Top = presentation.PageSetup.SlideHeight - selectedShape.Height;

            //    editedSelectedShapes.Add(selectedShape);
            //    count++;
            //}

            //if (!multiSlideZoomChecked)
            //{
            //    SingleSlideZoomToArea(currentSlide, editedSelectedShapes);
            //}
            //else
            //{
            //    MultiSlideZoomToArea(currentSlide, editedSelectedShapes);
            //}
            //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
            //foreach (PowerPoint.Shape sh in selectedShapes)
            //{
            //    sh.Visible = Office.MsoTriState.msoTrue;
            //    sh.Fill.ForeColor.RGB = 0xaaaaaa;
            //    sh.Fill.Transparency = 0.7f;
            //    sh.Line.ForeColor.RGB = 0x000000;
            //}
            //AddAckSlide();
        }
        public void AboutButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("          PowerPointLabs Plugin Version 1.7.2 [Release date: 3 Apr 2014]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website http://PowerPointLabs.info", "About PowerPointLabs");
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/docs.html";
                System.Diagnostics.Process.Start(myURL);
            }
            catch (Exception e)
            {
                LogException(e, "HelpButtonClick");
                throw;
            }
        }
        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/contact.html";
                System.Diagnostics.Process.Start(myURL);
            }
            catch (Exception e)
            {
                LogException(e, "FeedbackButtonClick");
                throw;
            }
        }
        public void AddZoomInButtonClick(Office.IRibbonControl control)
        {
            AutoZoom.AddDrillDownAnimation();
        }
        public void AddZoomOutButtonClick(Office.IRibbonControl control)
        {
            AutoZoom.AddStepBackAnimation();
        }

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.Application.
                ActiveWindow.Selection;
            CropShape(ref selection);
        }

        public System.Drawing.Bitmap GetAddAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAnimation);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimation);
            }
            catch (Exception e)
            {
                LogException(e, "GetReloadAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Spotlight);
            }
            catch (Exception e)
            {
                LogException(e, "GetSpotlightImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadSpotlight);
            }
            catch (Exception e)
            {
                LogException(e, "GetReloadSpotlightImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsTextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightText);
            }
            catch (Exception e)
            {
                LogException(e, "GetHighlightBulletsTextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsBackgroundImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightBackground);
            }
            catch (Exception e)
            {
                LogException(e, "GetHighlightBulletsBackgroundImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetHighlightBulletsTextContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightTextContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetHighlightBulletsTextContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsBackgroundContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.HighlightBackgroundContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetHighlightBulletsBackgroundContextImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetZoomInImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomIn);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomInImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetZoomOutImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomOut);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomOutImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomToAreaImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomToArea);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomToAreaImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomToAreaContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomToAreaContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomToAreaContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetCropShapeImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.CutOutShape);
            }
            catch (Exception e)
            {
                LogException(e, "GetCropShapeImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.About);
            }
            catch (Exception e)
            {
                LogException(e, "GetAboutImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Help);
            }
            catch (Exception e)
            {
                LogException(e, "GetHelpImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Feedback);
            }
            catch (Exception e)
            {
                LogException(e, "GetFeedbackImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAudio);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAudioImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetRemoveAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RemoveAudio);
            }
            catch (Exception e)
            {
                LogException(e, "GetRemoveAudioImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddCaption);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddCaptionsImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetRemoveCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RemoveCaption);
            }
            catch (Exception e)
            {
                LogException(e, "GetRemoveCaptionsImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetAddAudioContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddNarrationContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAudioContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetPreviewNarrationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.SpeakTextContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetPreviewNarrationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetInSlideAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.InSlideAnimation);
            }
            catch (Exception e)
            {
                LogException(e, "GetInSlideAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAnimationContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimationContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetReloadAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddSpotlightContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddSpotlightContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddSpotlightContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetEditNameContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.EditNameContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetEditNameContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetInSlideAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.InSlideContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetInSlideAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomInContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomInContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomInContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomOutContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomOutContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                if (pressed)
                {
                    backgroundZoomChecked = true;
                }
                else
                {
                    backgroundZoomChecked = false;
                }
            }
            catch (Exception e)
            {
                LogException(e, "ZoomStyleChanged");
                throw;
            }
        }
        public bool ZoomStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return backgroundZoomChecked;
            }
            catch (Exception e)
            {
                LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }
        //Control Enabled Callbacks
        public bool OnGetEnabledSpotlight(Office.IRibbonControl control)
        {
            return spotlightEnabled;
        }
        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return reloadSpotlight;
        }
        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return addAutoMotionEnabled;
        }
        public bool OnGetEnabledReloadAutoMotion(Office.IRibbonControl control)
        {
            return reloadAutoMotionEnabled;
        }
        public bool OnGetEnabledAddInSlide(Office.IRibbonControl control)
        {
            return inSlideEnabled;
        }
        public bool OnGetEnabledZoomButton(Office.IRibbonControl control)
        {
            return zoomButtonEnabled;
        }
        public bool OnGetEnabledHighlightBullets(Office.IRibbonControl control)
        {
            return highlightBulletsEnabled;
        }

        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                Form1 editForm = new Form1(this, selectedShape.Name);
                editForm.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "NameEditBtnClick");
                throw;
            }
        }
        public void ShapeNameEdited(String newName)
        {
            try
            {
                PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        public void AutoAnimateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimateDialogBox dialog = new AutoAnimateDialogBox(defaultDuration, frameAnimationChecked);
                dialog.SettingsHandler += AnimationPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "AutoAnimateDialogButtonPressed");
                throw;
            }
        }

        public void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            try
            {
                defaultDuration = newDuration;
                frameAnimationChecked = newFrameChecked;
                AnimateInSlide.defaultDuration = newDuration;
                AnimateInSlide.frameAnimationChecked = newFrameChecked;
                AutoAnimate.defaultDuration = newDuration;
                AutoAnimate.frameAnimationChecked = newFrameChecked;
            }
            catch (Exception e)
            {
                LogException(e, "AnimationPropertiesEdited");
                throw;
            }
        }

        public void AutoZoomDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                AutoZoomDialogBox dialog = new AutoZoomDialogBox(backgroundZoomChecked, multiSlideZoomChecked);
                dialog.SettingsHandler += ZoomPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "AutoZoomDialogButtonPressed");
                throw;
            }
        }

        public void ZoomPropertiesEdited(bool backgroundChecked, bool multiSlideChecked)
        {
            try
            {
                backgroundZoomChecked = backgroundChecked;
                multiSlideZoomChecked = multiSlideChecked;
                AutoZoom.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.multiSlideZoomChecked = multiSlideChecked;
            }
            catch (Exception e)
            {
                LogException(e, "ZoomPropertiesEdited");
                throw;
            }
        }

        public void SpotlightDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                SpotlightDialogBox dialog = new SpotlightDialogBox(Spotlight.defaultTransparency, Spotlight.defaultSoftEdges);
                dialog.SettingsHandler += SpotlightPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "SpotlightDialogButtonPressed");
                throw;
            }
        }

        public void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge)
        {
            try
            {
                Spotlight.defaultTransparency = newTransparency;
                Spotlight.defaultSoftEdges = newSoftEdge;
            }
            catch (Exception e)
            {
                LogException(e, "SpotlightPropertiesEdited");
                throw;
            }
        }

        public void HighlightBulletsPropertiesEdited(Color newHighlightColor, Color newDefaultColor, Color newBackgroundColor)
        {
            try
            {
                HighlightBulletsText.highlightColor = newHighlightColor;
                HighlightBulletsText.defaultColor = newDefaultColor;
                HighlightBulletsBackground.backgroundColor = newBackgroundColor;

            }
            catch (Exception e)
            {
                LogException(e, "HighlightBulletsPropertiesEdited");
                throw;
            }
        }
        public void HighlightBulletsDialogBoxPressed(Office.IRibbonControl control)
        {
            try
            {
                HighlightBulletsDialogBox dialog = new HighlightBulletsDialogBox(HighlightBulletsText.highlightColor, HighlightBulletsText.defaultColor, HighlightBulletsBackground.backgroundColor);
                dialog.SettingsHandler += HighlightBulletsPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "HighlightBulletsDialogBoxPressed");
                throw;
            }
        }

        #region NotesToAudio Button Callbacks
        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }

        public void RemoveAudioClick(Office.IRibbonControl control)
        {
            if (_allSlides)
            {
                NotesToAudio.RemoveAudioFromAllSlides();
            }
            else
            {
                NotesToAudio.RemoveAudioFromCurrentSlide();
            }
        }

        public void AddAudioClick(Office.IRibbonControl control)
        {
            if (_allSlides)
            {
                NotesToAudio.EmbedAllSlideNotes();
            }
            else
            {
                NotesToAudio.EmbedCurrentSlideNotes();
            }

            PreviewAnimationsIfChecked();
        }

        public void ContextAddAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.EmbedCurrentSlideNotes();
            PreviewAnimationsIfChecked();
        }
        #endregion

        #region NotesToCaptions Button Callbacks

        public void AddCaptionClick(Office.IRibbonControl control)
        {
            if (_captionsAllSlides)
            {
                NotesToCaptions.EmbedCaptionsOnAllSlides();
            }
            else
            {
                NotesToCaptions.EmbedCaptionsOnCurrentSlide();
            }
        }

        public void RemoveCaptionClick(Office.IRibbonControl control)
        {
            if (_captionsAllSlides)
            {
                NotesToCaptions.RemoveCaptionsFromAllSlides();
            }
            else
            {
                NotesToCaptions.RemoveCaptionsFromCurrentSlide();
            }
        }

        public void ContextReplaceAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.ReplaceSelectedAudio();
        }

        #endregion

        #region NotesToAudio/Caption Helpers
        public void AutoNarrateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoNarrateDialogBox(_voiceSelected, _voiceNames, _allSlides,
                    _previewCurrentSlide);
                dialog.SettingsHandler += AutoNarrateSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "AutoNarrateDialogButtonPressed");
                throw;
            }
        }

        public void AutoNarrateSettingsChanged(String voiceName, bool allSlides, bool previewCurrentSlide)
        {
            _allSlides = allSlides;
            _previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                _voiceSelected = _voiceNames.IndexOf(voiceName);
            }
        }

        public void AutoCaptionDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoCaptionDialogBox(_captionsAllSlides);
                dialog.SettingsHandler += AutoCaptionSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                LogException(e, "AutoCaptionDialogButtonPressed");
                throw;
            }
        }

        public void AutoCaptionSettingsChanged(bool allSlides)
        {
            _captionsAllSlides = allSlides;
        }

        private void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
        }

        #endregion

        private void SetCoreVoicesToSelections()
        {
            string defaultVoice = GetSelectedVoiceOrNull();
            NotesToAudio.SetDefaultVoice(defaultVoice);
        }

        private string GetSelectedVoiceOrNull()
        {
            string selectedVoice = null;
            try
            {
                selectedVoice = _voiceNames.ToArray()[_voiceSelected];
            }
            catch (IndexOutOfRangeException)
            {
                // No voices are installed.
                // (It should be impossible for the index to be out of range otherwise.)
            }
            return selectedVoice;
        }

        #endregion

        #region feature: Fit To Slide | Fit To Width | Fit To Height

        private const int SelectedShapeIndex = 1;
        private const int TopMost = 0;
        private const int LeftMost = 0;

        public void FitToWidthClick(Office.IRibbonControl control)
        {
            DoFitToWidth();
        }

        public void FitToHeightClick(Office.IRibbonControl control)
        {
            DoFitToHeight();
        }

        private void DoFitToHeight()
        {
            var pageSetup = GetPageSetup();
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape.Height, selectedShape.Width);
            float factor = GetRotationFactorForFitToHeight(ref selectedShape);
            //fit to height
            selectedShape.Height = (pageSetup.SlideHeight / factor);
            selectedShape.Width = selectedShape.Height / shapeSizeRatio;
            //move to centre
            selectedShape.Left = (pageSetup.SlideWidth - selectedShape.Width) / 2;
            selectedShape.Top = TopMost;
            AdjustMoveToCentreForFitToHeight(ref selectedShape);
        }

        private void AdjustMoveToCentreForFitToHeight(ref PowerPoint.Shape shape)
        {
            float adjustLength;
            float rotation = SetupRotationValueForAdjustment(ref shape);
            //Pythagorean theorem
            float diagonal = (float)Math.Sqrt(Math.Pow(shape.Width / 2, 2)
                + Math.Pow(shape.Height / 2, 2));
            //Law of cosines
            float oppositeSideLength = (float)Math.Sqrt((Math.Pow(diagonal, 2) * 2
                * (1 - Math.Cos(rotation))));
            float angle1 = (float)Math.Atan(shape.Width / shape.Height);
            float angle2 = (float)((Math.PI - rotation) / 2);
            //case 1:
            if ((shape.Rotation >= 0 && shape.Rotation <= 90)
                || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                float targetAngle = (float)(Math.PI - angle1 - angle2);
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle));
            }
            //case 2:
            else/* case: 90 < rotation < 270)*/
            {
                float targetAngle = angle1 - angle2;
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle)) - shape.Height;
            }
            shape.Top += adjustLength;
        }

        private float GetRotationFactorForFitToHeight(ref PowerPoint.Shape shape)
        {
            //set up rotation value
            float rotation = SetupRotationValueForRotateFactor(ref shape);
            //calculate factor for Fit to Height
            float shapeSizeRatio = GetSizeRatio(shape.Height, shape.Width);
            float factor = (float)(Math.Cos(rotation) + Math.Sin(rotation) / shapeSizeRatio);
            return factor;
        }

        private void DoFitToWidth()
        {
            var pageSetup = GetPageSetup();
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape.Height, selectedShape.Width);
            //fit to width
            float factor = GetRotationFactorForFitToWidth(ref selectedShape);
            selectedShape.Height = pageSetup.SlideWidth / factor;
            selectedShape.Width = selectedShape.Height / shapeSizeRatio;
            //move to middle
            selectedShape.Top = (pageSetup.SlideHeight - selectedShape.Height) / 2;
            selectedShape.Left = LeftMost;
            //adjustment for rotation
            AdjustMoveToCentreForFitToWidth(ref selectedShape);
        }

        private void AdjustMoveToCentreForFitToWidth(ref PowerPoint.Shape shape)
        {
            float adjustLength;
            float rotation = SetupRotationValueForAdjustment(ref shape);
            //Pythagorean theorem
            float diagonal = (float)Math.Sqrt(Math.Pow(shape.Width / 2, 2)
                + Math.Pow(shape.Height / 2, 2));
            //Law of cosines
            float oppositeSideLength = (float)Math.Sqrt((Math.Pow(diagonal, 2) * 2
                * (1 - Math.Cos(rotation))));
            float angle1 = (float)Math.Atan(shape.Height / shape.Width);
            float angle2 = (float)((Math.PI - rotation) / 2);
            //case 1:
            if ((shape.Rotation >= 0 && shape.Rotation <= 90)
                || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                float targetAngle = (float)(Math.PI - angle1 - angle2);
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle));
            }
            //case 2:
            else/* case: 90 < rotation < 270)*/
            {
                float targetAngle = angle1 - angle2;
                adjustLength = (float)(oppositeSideLength * Math.Cos(targetAngle)) - shape.Width;
            }
            shape.Left += adjustLength;
        }

        private float SetupRotationValueForAdjustment(ref PowerPoint.Shape shape)
        {
            float rotation = shape.Rotation;
            if (shape.Rotation > 180 && shape.Rotation <= 360)
            {
                rotation = 360 - shape.Rotation;
            }
            return ConvertDegToRad(rotation);
        }

        private float GetRotationFactorForFitToWidth(ref PowerPoint.Shape shape)
        {
            float rotation = SetupRotationValueForRotateFactor(ref shape);
            //calculate factor for Fit to Height
            float shapeSizeRatio = GetSizeRatio(shape.Height, shape.Width);
            float factor = (float)(Math.Sin(rotation) + Math.Cos(rotation) / shapeSizeRatio);
            return factor;
        }

        private float SetupRotationValueForRotateFactor(ref PowerPoint.Shape shape)
        {
            float rotation;
            if (shape.Rotation == 90.0)
            {
                rotation = shape.Rotation;
            }
            else if (shape.Rotation == 270.0)
            {
                rotation = 360 - shape.Rotation;
            }
            else if ((shape.Rotation > 90 && shape.Rotation <= 180)
                     || (shape.Rotation > 270 && shape.Rotation <= 360))
            {
                rotation = (360 - shape.Rotation) % 90;
            }
            else
            {
                rotation = shape.Rotation % 90;
            }
            return ConvertDegToRad(rotation);
        }

        private float ConvertDegToRad(float rotation)
        {
            rotation = (float)((rotation) * Math.PI / 180); return rotation;
        }

        private float GetSizeRatio(float height, float width)
        {
            return height / width;
        }

        private PowerPoint.PageSetup GetPageSetup()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup;
        }

        private PowerPoint.Shape GetSelectedShape()
        {
            return Globals.ThisAddIn.Application.
                   ActiveWindow.Selection.ShapeRange[SelectedShapeIndex];
        }

        public System.Drawing.Bitmap GetFitToWidthImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.FitToWidth);
            }
            catch (Exception e)
            {
                LogException(e, "GetFitToWidthImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetFitToHeightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.FitToHeight);
            }
            catch (Exception e)
            {
                LogException(e, "GetFitToHeightImage");
                throw;
            }
        }

        #endregion

        #region feature: Crop to Shape
        public PowerPoint.Shape CropShapeToSlide(ref PowerPoint.Selection selection)
        {
            try
            {
                IsValidSelectionForCropToShape(ref selection);
                var shape = FormShapeForCropToShape(ref selection);
                TakeScreenshotWithoutShape(ref shape);
                return FillInShapeWithScreenshotOfSlide(shape);
                //return shape;
            }
            catch (Exception)
            {
                MessageBox.Show(GetCropToShapeErrorMessage(), "Unable to crop");
                CropToShapeErrorCode = -1;
                return null;
            }
        }

        private void CropShape(ref PowerPoint.Selection selection)
        {
            try
            {
                IsValidSelectionForCropToShape(ref selection);
                var shape = FormShapeForCropToShape(ref selection);
                TakeScreenshotWithoutShape(ref shape);
                FillInShapeWithScreenshot(shape);
            }
            catch (Exception)
            {
                MessageBox.Show(GetCropToShapeErrorMessage(), "Unable to crop");
                CropToShapeErrorCode = -1;
            }
        }

        private int CropToShapeErrorCode = -1;
        private const int ERROR_SELECTION_COUNT_ZERO = 0;
        private const int ERROR_SELECTION_NON_SHAPE = 1;
        private const int ERROR_EXCEED_SLIDE_BOUND = 2;
        private const int ERROR_ROTATION_NON_ZERO = 3;
        const string OfficeVersion2013 = "15.0";
        const string OfficeVersion2010 = "14.0";

        private string GetCropToShapeErrorMessage()
        {
            switch (CropToShapeErrorCode)
            {
                case ERROR_SELECTION_COUNT_ZERO:
                    return "To start 'Crop To Shape', please select at least one shape.";
                case ERROR_SELECTION_NON_SHAPE:
                    return "'Crop To Shape' only supports shape objects.";
                case ERROR_EXCEED_SLIDE_BOUND:
                    return "Please ensure your shape is within the slide.";
                case ERROR_ROTATION_NON_ZERO:
                    return "In the current version, the 'Crop To Shape' feature does not work if the shape is rotated";
                default:
                    return "Undefined error.";
            }
        }

        private PowerPoint.Shape FillInShapeWithScreenshotOfSlide(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup)
            {
                ProduceFillInBackground(ref shape);
                shape.Fill.UserPicture(GetPathForFillInBackground());
            }
            else
            {
                using (Bitmap slideImage = (Bitmap)Bitmap.FromFile(GetPathToStore()))
                {
                    foreach (var sh in shape.GroupItems)
                    {
                        var shapeGroupItem = sh as PowerPoint.Shape;
                        ProduceFillInBackground(ref shapeGroupItem, slideImage);
                        shapeGroupItem.Fill.UserPicture(GetPathForFillInBackground());
                    }
                }
            }
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Copy();
            PowerPoint.Shape returnShape = GetCurrentSlide().Shapes.Paste()[1];
            shape.Delete();
            return returnShape;
        }

        private void FillInShapeWithScreenshot(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoGroup)
            {
                ProduceFillInBackground(ref shape);
                shape.Fill.UserPicture(GetPathForFillInBackground());
            }
            else
            {
                using (Bitmap slideImage = (Bitmap)Bitmap.FromFile(GetPathToStore()))
                {
                    foreach (var sh in shape.GroupItems)
                    {
                        var shapeGroupItem = sh as PowerPoint.Shape;
                        ProduceFillInBackground(ref shapeGroupItem, slideImage);
                        shapeGroupItem.Fill.UserPicture(GetPathForFillInBackground());
                    }
                }
            }
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Copy();
            GetCurrentSlide().Shapes.Paste();
            shape.Delete();
        }

        private void AdjustFillEffect(ref PowerPoint.Shape shape)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                var range = shape.Ungroup();
                foreach (var o in range)
                {
                    var sh = o as PowerPoint.Shape;
                    var tmpFillEffect = sh.Fill;
                    tmpFillEffect.TextureOffsetX = -sh.Left;
                    tmpFillEffect.TextureOffsetY = -sh.Top;
                }
                shape = range.Group();
            }
        }

        private void ProduceFillInBackground(ref PowerPoint.Shape shape)
        {
            using (Bitmap slideImage = (Bitmap)Bitmap.FromFile(GetPathToStore()))
            {
                float horizontalRatio =
                    (float)(GetDesiredExportWidth() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth);
                float verticalRatio =
                    (float)(GetDesiredExportHeight() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight);
                Bitmap croppedImage = KiCut(slideImage,
                    shape.Left * horizontalRatio,
                    shape.Top * verticalRatio,
                    shape.Width * horizontalRatio,
                    shape.Height * verticalRatio);
                croppedImage.Save(GetPathForFillInBackground(), ImageFormat.Png);
            }
        }

        private void ProduceFillInBackground(ref PowerPoint.Shape shape, Bitmap slideImage)
        {
            float horizontalRatio =
                (float)(GetDesiredExportWidth() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth);
            float verticalRatio =
                (float)(GetDesiredExportHeight() / Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight);
            Bitmap croppedImage = KiCut(slideImage,
                shape.Left * horizontalRatio,
                shape.Top * verticalRatio,
                shape.Width * horizontalRatio,
                shape.Height * verticalRatio);
            croppedImage.Save(GetPathForFillInBackground(), ImageFormat.Png);
        }

        public static Bitmap KiCut(Bitmap original, float startX, float startY, float width, float height)
        {
            if (original == null)
            {
                return null;
            }
            if (startX >= original.Width || startY >= original.Height)
            {
                return null;
            }
            try
            {
                Bitmap outputImage = new Bitmap((int)width, (int)height, PixelFormat.Format24bppRgb);

                Graphics inputGraphics = Graphics.FromImage(outputImage);
                inputGraphics.DrawImage(original,
                    new Rectangle(0, 0, (int)width, (int)height),
                    new Rectangle((int)startX, (int)startY, (int)width, (int)height),
                    GraphicsUnit.Pixel);
                inputGraphics.Dispose();

                return outputImage;
            }
            catch
            {
                return null;
            }
        }

        private string GetPathForFillInBackground()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            return path + "\\currentFillInBg.png";
        }

        private string GetPathToStore()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            return path + "\\currentSlide.png";
        }

        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.Application.Version == OfficeVersion2010;
        }

        private PowerPoint.Shape ConvertToPicture(ref PowerPoint.Shape shape)
        {
            float rotation = 0;
            try
            {
                rotation = shape.Rotation;
                shape.Rotation = 0;
            }
            catch (Exception)
            {
                //chart cannot be rotated
            }
            shape.Copy();
            float x = shape.Left;
            float y = shape.Top;
            float width = shape.Width;
            float height = shape.Height;
            shape.Delete();
            var pic = GetCurrentSlide().Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pic.Left = x + (width - pic.Width) / 2;
            pic.Top = y + (height - pic.Height) / 2;
            pic.Rotation = rotation;
            pic.Select();
            return pic;
        }

        private bool IsFirstOneOverlapWithSecond(PowerPoint.Shape first, PowerPoint.Shape second)
        {
            if (first == second)
                return false;
            if (!(first.Left > second.Left + second.Width ||
                  first.Left + first.Width < second.Left ||
                  first.Top > second.Top + second.Height ||
                  first.Top + first.Height < second.Top))
            {
                return true;
            }
            return false;
        }

        private void RemoveOverlapItems(PowerPoint.Shape shape)
        {
            var shapeRange = GetCurrentSlide().Shapes;
            for (int i = 1; i < shapeRange.Count; )
            {
                if (IsFirstOneOverlapWithSecond(shapeRange[i], shape))
                    shapeRange[i].Delete();
                else
                    i++;
            }
        }

        private void TakeScreenshotWithoutShape(ref PowerPoint.Shape shape)
        {
            shape.Visible = Office.MsoTriState.msoFalse;
            TakeScreenshot();
            shape.Visible = Office.MsoTriState.msoTrue;
        }

        private void TakeScreenshot()
        {
            GetCurrentSlide().Export(GetPathToStore(), "PNG",
                (int)GetDesiredExportWidth(), (int)GetDesiredExportHeight());
        }

        private double GetDesiredExportWidth()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth / 72.0 * 96.0;
        }

        private double GetDesiredExportHeight()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight / 72.0 * 96.0;
        }

        private PowerPoint.Shape FormShapeForCropToShape(ref PowerPoint.Selection selection)
        {
            var oldRange = selection.ShapeRange;
            //some shapes in the nameList cannot be used due to 
            //Powerpoint's 'Delete-Undo' bug: when a shape got deleted or cut, then undo,
            //then we can only read its name/width/height/left/top.. for others, it aborts
            //var nameList = ProduceNameList(GetCurrentSlide().Shapes.Range());
            //'Cut-Paste' is a normal workaround method for the bug mentioned above
            oldRange.Cut();
            oldRange = GetCurrentSlide().Shapes.Paste();
            var copyRange = MakeCopyWithSamePosition(oldRange);
            var ungroupedCopyRange = UngroupAllShapes(copyRange);
            PowerPoint.Shape mergedShape = ungroupedCopyRange[1];
            //if (Globals.ThisAddIn.Application.Version == OfficeVersion2013)
            //{
            //    mergedShape = MergeAllShapes(ref ungroupedCopyRange, nameList);
            //    RemoveRotation(ref mergedShape, nameList);
            //}
            //else 
            //if (Globals.ThisAddIn.Application.Version == OfficeVersion2010)
            //{
            if (ungroupedCopyRange.Count > 1)
            {
                mergedShape = ungroupedCopyRange.Group();
            }
            //}
            if (IsWithinSlide(mergedShape))
            {
                oldRange.Delete();
            }
            else
            {
                mergedShape.Delete();
                ThrowExceptionFromCropToShape(ERROR_EXCEED_SLIDE_BOUND);
            }
            return mergedShape;
        }

        private PowerPoint.ShapeRange MakeCopyWithSamePosition(PowerPoint.ShapeRange oldRange)
        {
            //setup shapes' names in oldRange, so that shapes' names in copyRange is the same
            ModifyNames(oldRange);
            oldRange.Copy();
            var copyRange = GetCurrentSlide().Shapes.Paste();
            //adjust position
            Dictionary<string, Tuple<float, float>> dict = new Dictionary<string, Tuple<float, float>>();
            foreach (var sh in oldRange)
            {
                var shape = sh as PowerPoint.Shape;
                dict.Add(shape.Name, new Tuple<float, float>(shape.Left, shape.Top));
            }
            foreach (var sh in copyRange)
            {
                var shape = sh as PowerPoint.Shape;
                shape.Left = dict[shape.Name].Item1;
                shape.Top = dict[shape.Name].Item2;
            }
            //append something in case we want to differentiate them
            ModifyNames(copyRange, "_Copy");
            return copyRange;
        }

        private List<string> ProduceNameList(PowerPoint.ShapeRange range)
        {
            List<string> output = new List<string>();
            foreach (var sh in range)
            {
                output.Add((sh as PowerPoint.Shape).Name);
            }
            return output;
        }

        private void ModifyNames(PowerPoint.ShapeRange range, string appendString = "")
        {
            if (appendString != "")
            {
                foreach (var sh in range)
                {
                    (sh as PowerPoint.Shape).Name += appendString;
                }
            }
            else
            {
                foreach (var sh in range)
                {
                    (sh as PowerPoint.Shape).Name = Guid.NewGuid().ToString();
                }
            }
        }

        private bool IsWithinSlide(PowerPoint.Shape shape)
        {
            //-1 and +1 for better user experience
            bool cond1 = shape.Left >= -1;
            bool cond2 = shape.Top >= -1;
            bool cond3 = shape.Left + shape.Width <= Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth + 1;
            bool cond4 = shape.Top + shape.Height <= Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight + 1;
            return cond1 && cond2 && cond3 && cond4;
        }

        private PowerPoint.ShapeRange UngroupAllShapes(PowerPoint.ShapeRange range)
        {
            List<string> ungroupedShapes = new List<string>();
            Queue<PowerPoint.Shape> selectedShapes = new Queue<PowerPoint.Shape>();

            foreach (var item in range)
            {
                selectedShapes.Enqueue(item as PowerPoint.Shape);
            }
            while (selectedShapes.Count != 0)
            {
                var shape = selectedShapes.Dequeue();
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (var item in subRange)
                    {
                        selectedShapes.Enqueue(item as PowerPoint.Shape);
                    }
                }
                else if (/*Globals.ThisAddIn.Application.Version == OfficeVersion2010
                    && */shape.Rotation != 0.0)
                {
                    //remove copies before throwing exception
                    shape.Delete();
                    if (ungroupedShapes.Count > 0)
                    {
                        GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray()).Delete();
                    }
                    while (selectedShapes.Count != 0)
                    {
                        selectedShapes.Dequeue().Delete();
                    }
                    ThrowExceptionFromCropToShape(ERROR_ROTATION_NON_ZERO);
                }
                else if (!IsShape(shape))
                {
                    //remove copies before throwing exception
                    shape.Delete();
                    if (ungroupedShapes.Count > 0)
                    {
                        GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray()).Delete();
                    }
                    while (selectedShapes.Count != 0)
                    {
                        selectedShapes.Dequeue().Delete();
                    }
                    ThrowExceptionFromCropToShape(ERROR_SELECTION_NON_SHAPE);
                }
                else
                {
                    shape.Name = Guid.NewGuid().ToString();
                    ungroupedShapes.Add(shape.Name);
                }
            }
            return GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray());
        }

        //private PowerPoint.Shape MergeAllShapes(ref PowerPoint.ShapeRange range, List<string> nameList)
        //{
        //    PowerPoint.Shape shape = range[1];
        //    if (range.Count > 1)
        //    {
        //        int KaisBirthday = 0x900209;
        //        range[1].Line.ForeColor.RGB = KaisBirthday;
        //        //merged shape will inherit the foreColor of range[1]'s line,
        //        //which is Xie Kai's birthday :p
        //        //a better choice can be: random number or timestamp
        //        range.MergeShapes(Office.MsoMergeCmd.msoMergeUnion, range[1]);
        //        //find the merged shape
        //        var newRange = GetCurrentSlide().Shapes.Range();
        //        foreach (var sh in newRange)
        //        {
        //            var mergedShape = sh as PowerPoint.Shape;
        //            if (!nameList.Contains(mergedShape.Name))
        //            {
        //                if(KaisBirthday == mergedShape.Line.ForeColor.RGB)
        //                {
        //                    shape = mergedShape;
        //                }
        //            }
        //        }
        //    }
        //    return shape;
        //}

        //private void RemoveRotation(ref PowerPoint.Shape shape, List<string> nameList)
        //{
        //    var helperShape = GetCurrentSlide().Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, -100, -100, 1, 1);
        //    int KaisBirthday = 0x900209;
        //    helperShape.Line.ForeColor.RGB = KaisBirthday;
        //    //merged shape will inherit the foreColor of range[1]'s line,
        //    //which is Xie Kai's birthday :p
        //    //a better choice can be: random number or timestamp
        //    var range = GetCurrentSlide().
        //        Shapes.Range(new List<string> { helperShape.Name, shape.Name }.ToArray());
        //    //Separate the shapes and make rotation back to zero
        //    range.MergeShapes(Office.MsoMergeCmd.msoMergeFragment, helperShape);
        //    //find those resulted shapes
        //    var newRange = GetCurrentSlide().Shapes.Range();
        //    List<string> list = new List<string>();
        //    foreach (var sh in newRange)
        //    {
        //        var mergedShape = sh as PowerPoint.Shape;
        //        if (mergedShape.Left == -100
        //            && mergedShape.Top == -100
        //            && mergedShape.Width == 1
        //            && mergedShape.Height == 1)
        //        {
        //            helperShape = mergedShape;
        //        }
        //        else if (!nameList.Contains(mergedShape.Name))
        //        {
        //            if(KaisBirthday == mergedShape.Line.ForeColor.RGB)
        //            {
        //                list.Add(mergedShape.Name);
        //            }
        //        }
        //    }
        //    helperShape.Delete();
        //    range = GetCurrentSlide().Shapes.Range(list.ToArray());
        //    if (list.Count > 1)
        //    {
        //        range.MergeShapes(Office.MsoMergeCmd.msoMergeUnion);
        //        //find out the merged shape
        //        newRange = GetCurrentSlide().Shapes.Range();
        //        foreach (var sh in newRange)
        //        {
        //            var mergedShape = sh as PowerPoint.Shape;
        //            if (!nameList.Contains(mergedShape.Name))
        //            {
        //                if (KaisBirthday == mergedShape.Line.ForeColor.RGB)
        //                {
        //                    shape = mergedShape;
        //                }
        //            }
        //        }
        //    }
        //    else
        //    {
        //        shape = range[1];
        //    }
        //}

        private void ThrowExceptionFromCropToShape(int typeOfError)
        {
            CropToShapeErrorCode = typeOfError;
            throw new Exception("Error: " + typeOfError.ToString());
        }

        private void IsValidSelectionForCropToShape(ref PowerPoint.Selection selection)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (selection.ShapeRange.Count < 1)
                {
                    ThrowExceptionFromCropToShape(ERROR_SELECTION_COUNT_ZERO);
                }
                if (!IsShape(selection))
                {
                    ThrowExceptionFromCropToShape(ERROR_SELECTION_NON_SHAPE);
                }
            }
            else
            {
                ThrowExceptionFromCropToShape(ERROR_SELECTION_COUNT_ZERO);
            }
        }

        private bool IsShape(PowerPoint.Selection sel)
        {
            var shapeRange = sel.ShapeRange;
            foreach (var o in shapeRange)
            {
                var shape = o as PowerPoint.Shape;
                if (!IsShape(shape))
                    return false;
            }
            return true;
        }

        private bool IsShape(PowerPoint.Shape shape)
        {
            if (shape.Type != Office.MsoShapeType.msoAutoShape
                    && shape.Type != Office.MsoShapeType.msoFreeform
                    && shape.Type != Office.MsoShapeType.msoGroup)
            {
                return false;
            }
            return true;
        }

        public System.Drawing.Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.CutOutShapeMenu);
            }
            catch (Exception e)
            {
                LogException(e, "GetCutOutShapeMenuImage");
                throw;
            }
        }

        #endregion

        #region feature: Convert to Picture

        public void ConvertToPictureButtonClick(Office.IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.Application.
                ActiveWindow.Selection;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var shape = selection.ShapeRange[1];
                if (selection.ShapeRange.Count > 1)
                {
                    shape = selection.ShapeRange.Group();
                }
                shape.Cut();
                shape = GetCurrentSlide().Shapes.Paste()[1];
                ConvertToPicture(ref shape);
            }
            else
            {
                MessageBox.Show("Convert to Picture only supports Shapes and Charts.", "Unable to Convert to Picture");
            }
        }

        public System.Drawing.Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ConvertToPicture);
            }
            catch (Exception e)
            {
                LogException(e, "GetConvertToPicMenuImage");
                throw;
            }
        }

        #endregion

        #region Helpers
        private bool DeleteShapeAnnimations(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            try
            {
                PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
                bool flag = true;
                for (int x = sequence.Count; x >= 1; x--)
                {
                    PowerPoint.Effect effect = sequence[x];
                    if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    {
                        if (effect.Exit == Office.MsoTriState.msoTrue)
                            flag = false;
                        effect.Delete();
                    }
                }
                return flag;
            }
            catch (Exception e)
            {
                LogException(e, "DeleteShapeAnimations");
                throw;
            }
        }
        private bool HasExitAnimation(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            try
            {
                PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                bool flag = false;
                for (int x = sequence.Count; x >= 1; x--)
                {
                    PowerPoint.Effect effect = sequence[x];
                    if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    {
                        if (effect.Exit == Office.MsoTriState.msoTrue)
                        {
                            flag = true;
                            break;
                        }
                    }
                }

                return flag;
            }
            catch (Exception e)
            {
                LogException(e, "HasEntryAnimation");
                throw;
            }
        }
        private void MoveMotionAnimation(PowerPoint.Slide slide)
        {
            try
            {
                PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                foreach (PowerPoint.Effect eff in slide.TimeLine.MainSequence)
                {
                    if ((eff.EffectType >= PowerPoint.MsoAnimEffect.msoAnimEffectPathCircle && eff.EffectType <= PowerPoint.MsoAnimEffect.msoAnimEffectPathRight) || eff.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectCustom)
                    {
                        //sh.Delete();
                        PowerPoint.AnimationBehavior motion = eff.Behaviors[1];
                        if (motion.Type == PowerPoint.MsoAnimType.msoAnimTypeMotion)
                        {
                            PowerPoint.Shape sh = eff.Shape;
                            string motionPath = motion.MotionEffect.Path.Trim();
                            if (motionPath.Last() < 'A' || motionPath.Last() > 'Z')
                                motionPath += " X";
                            string[] path = motionPath.Split(' ');
                            int count = path.Length;
                            float xVal = Convert.ToSingle(path[count - 3]);
                            float yVal = Convert.ToSingle(path[count - 2]);
                            sh.Left += (xVal * presentation.PageSetup.SlideWidth);
                            sh.Top += (yVal * presentation.PageSetup.SlideHeight);
                        }
                    }
                }
            }
            catch (Exception e)
            {
            }
        }
        //Other Helpers
        private void AddAckSlide()
        {
            try
            {
                PowerPoint.Slide tempSlide = (Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActivePresentation.Slides.Count]);
                if (!(tempSlide.Name.Contains("PPAck") && tempSlide.Name.Substring(0, 5).Equals("PPAck")))
                {
                    PowerPoint.Slide ackSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides.Count + 1, GetCurrentSlide().CustomLayout);
                    //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(ackSlide.SlideIndex);
                    PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                    String tempFileName = Path.GetTempFileName();
                    Properties.Resources.Acknowledgement.Save(tempFileName);
                    float width = presentation.PageSetup.SlideWidth * 0.858f;
                    float height = presentation.PageSetup.SlideHeight * (5.33f / 7.5f);
                    PowerPoint.Shape ackShape = ackSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, ((presentation.PageSetup.SlideWidth - width) / 2), ((presentation.PageSetup.SlideHeight - height) / 2), width, height);
                    ackSlide.SlideShowTransition.Hidden = Office.MsoTriState.msoTrue;
                    ackSlide.Name = "PPAck" + GetTimestamp(DateTime.Now);
                }
            }
            catch (Exception e)
            {
                LogException(e, "AddAckSlide");
                throw;
            }
        }
        private void Log(string logText, string type)
        {
            if (type.Equals("Info"))
                Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            else if (type.Equals("Error"))
                Trace.TraceError(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            else if (type.Equals("Warning"))
                Trace.TraceWarning(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
        }
        private void LogException(Exception e, string methodName)
        {
            Log(methodName + ": " + e.Message + ": " + e.StackTrace, "Error");
        }
        private PowerPoint.Slide GetCurrentSlide()
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                return Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            }
            catch (Exception e)
            {
                LogException(e, "GetCurrentSlide");
                throw;
            }
        }
        private PowerPoint.Slide GetNextSlide(PowerPoint.Slide currentSlide)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                int slideIndex = currentSlide.SlideIndex;
                return presentation.Slides[slideIndex + 1];
            }
            catch (Exception e)
            {
                LogException(e, "GetNextSlide");
                throw;
            }
        }
        public String GetTimestamp(DateTime value)
        {
            try
            {
                return value.ToString("yyyyMMddHHmmssffff");
            }
            catch (Exception e)
            {
                LogException(e, "GetTimestamp");
                throw;
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
