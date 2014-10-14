using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using ImageProcessor.Imaging.Filters;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ImageProcessor;

// Follow these steps to enable the Ribbon (XML) item:

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
        private Office.IRibbonUI _ribbon;
        
        public bool FrameAnimationChecked = false;
        public bool BackgroundZoomChecked = true;
        public bool MultiSlideZoomChecked = true;
        public bool SpotlightDelete = true;
        public float DefaultDuration = 0.5f;

        public bool SpotlightEnabled = false;
        public bool InSlideEnabled = false;
        public bool ZoomButtonEnabled = false;
        public bool HighlightBulletsEnabled = true;
        public bool AddAutoMotionEnabled = true;
        public bool ReloadAutoMotionEnabled = true;
        public bool ReloadSpotlight = true;
        public bool RemoveCaptionsEnabled = true;
        public bool RemoveAudioEnabled = true;

        public bool HighlightTextFragmentsEnabled = true;

        public bool EmbedAudioVisible = true;
        public bool RecorderPaneVisible = false;

        private bool _previewCurrentSlide;
        
        private List<string> _voiceNames;

        private int _voiceSelected;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void RibbonLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;

            SetVoicesFromInstalledOptions();
            SetCoreVoicesToSelections();
        }

        public void RefreshRibbonControl(String controlId)
        {
            try
            {
                _ribbon.InvalidateControl(controlId);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "RefreshRibbonControl");
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
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsBackgroundButtonClick");
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
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsTextButtonClick");
                throw;
            }
        }
        public void HighlightTextFragmentsButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kTextSelected;
                else
                    HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kNoneSelected;

                HighlightTextFragments.AddHighlightedTextFragments();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightTextFragmentsButtonClick");
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
                PowerPointLabsGlobals.LogException(e, "AddInSlideAnimationButtonClick");
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
                PowerPointLabsGlobals.LogException(e, "ReloadSpotlightButtonClick");
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
                PowerPointLabsGlobals.LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        # region Supertips
        public string GetAddAnimationButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationButtonSupertip;
        }
        public string GetReloadButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ReloadButtonSupertip;
        }
        public string GetInSlideAnimateButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.InSlideAnimateButtonSupertip;
        }
        
        public string GetAddZoomInButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomInButtonSupertip;
        }
        public string GetAddZoomOutButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomOutButtonSupertip;
        }
        public string GetZoomToAreaButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ZoomToAreaButtonSupertip;
        }
        
        public string GetMoveCropShapeButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.MoveCropShapeButtonSupertip;
        }
        
        public string GetAddSpotlightButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddSpotlightButtonSupertip;
        }
        public string GetReloadSpotlightButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ReloadSpotlightButtonSupertip;
        }
        
        public string GetAddAudioButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddAudioButtonSupertip;
        }
        public string GetGenerateRecordButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.GenerateRecordButtonSupertip;
        }
        public string GetAddRecordButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddRecordButtonSupertip;
        }
        public string GetRemoveAudioButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAudioButtonSupertip;
        }
        
        public string GetAddCaptionsButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddCaptionsButtonSupertip;
        }
        public string GetRemoveCaptionsButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.RemoveCaptionsButtonSupertip;
        }
        public string GetRemoveAllNotesButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAllNotesButtonSupertip;
        }
        
        public string GetHighlightBulletsTextButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextButtonSupertip;
        }
        public string GetHighlightBulletsBackgroundButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundButtonSupertip;
        }

        public string GetHighlightTextFragmentsButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightTextFragmentsButtonSupertip;
        }
        
        public string GetColorPickerButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ColorPickerButtonSupertip;
        }
        
        public string GetCustomeShapeButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.CustomeShapeButtonSupertip;
        }
        
        public string GetHelpButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HelpButtonSupertip;
        }
        public string GetFeedbackButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.FeedbackButtonSupertip;
        }
        public string GetAboutButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AboutButtonSupertip;
        }
        # endregion

        # region Button Labels
        public string GetPowerPointLabsAddInsTabLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsAddInsTabLabel;
        }

        public string GetCombineShapesLabel(Office.IRibbonControl control)
        {
            return TextCollection.CombineShapesLabel;
        }

        public string GetAutoAnimateGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.AutoAnimateGroupLabel;
        }
        public string GetAddAnimationButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationButtonLabel;
        }
        public string GetReloadButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationReloadButtonLabel;
        }
        public string GetInSlideAnimateButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationInSlideAnimateButtonLabel;
        }

        public string GetAutoZoomGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.AutoZoomGroupLabel;
        }
        public string GetAddZoomInButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomInButtonLabel;
        }
        public string GetAddZoomOutButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomOutButtonLabel;
        }
        public string GetZoomToAreaButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomToAreaButtonLabel;
        }

        public string GetAutoCropGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.AutoCropGroupLabel;
        }
        public string GetMoveCropShapeButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.MoveCropShapeButtonLabel;
        }

        public string GetSpotLightGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.SpotLightGroupLabel;
        }
        public string GetAddSpotlightButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddSpotlightButtonLabel;
        }
        public string GetReloadSpotlightButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.ReloadSpotlightButtonLabel;
        }

        public string GetEmbedAudioGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.EmbedAudioGroupLabel;
        }
        public string GetAddAudioButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAudioButtonLabel;
        }
        public string GetGenerateRecordButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.GenerateRecordButtonLabel;
        }
        public string GetAddRecordButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddRecordButtonLabel;
        }
        public string GetRemoveAudioButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAudioButtonLabel;
        }

        public string GetEmbedCaptionGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.EmbedCaptionGroupLabel;
        }
        public string GetAddCaptionsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddCaptionsButtonLabel;
        }
        public string GetRemoveCaptionsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.RemoveCaptionsButtonLabel;
        }
        public string GetRemoveAllNotesButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAllNotesButtonLabel;
        }

        public string GetHighlightBulletsGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsGroupLabel;
        }
        public string GetHighlightBulletsTextButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextButtonLabel;
        }
        public string GetHighlightBulletsBackgroundButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundButtonLabel;
        }
        public string GetHighlightTextFragmentsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightTextFragmentsButtonLabel;
        }
        
        public string GetLabsGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.LabsGroupLabel;
        }
        public string GetColorPickerButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.ColorPickerButtonLabel;
        }
        public string GetCustomeShapeButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.CustomeShapeButtonLabel;
        }

        public string GetPPTLabsHelpGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.PPTLabsHelpGroupLabel;
        }
        public string GetHelpButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HelpButtonLabel;
        }
        public string GetFeedbackButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.FeedbackButtonLabel;
        }
        public string GetAboutButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AboutButtonLabel;
        }

        public string GetNameEditShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.NameEditShapeLabel;
        }
        public string GetSpotlightShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightShapeLabel;
        }
        public string GetZoomInContextMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomInContextMenuLabel;
        }
        public string GetZoomOutContextMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomOutContextMenuLabel;
        }
        public string GetZoomToAreaContextMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomToAreaContextMenuLabel;
        }
        public string GetHighlightBulletsMenuShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsMenuShapeLabel;
        }
        public string GetHighlightBulletsBackgroundShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundShapeLabel;
        }
        public string GetHighlightBulletsTextShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextShapeLabel;
        }
        public string GetConvertToPictureShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.ConvertToPictureShapeLabel;
        }
        public string GetAddCustomShapeShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
        public string GetCutOutShapeShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.CutOutShapeShapeLabel;
        }
        public string GetFitToWidthShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.FitToWidthShapeLabel;
        }
        public string GetFitToHeightShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.FitToHeightShapeLabel;
        }
        public string GetInSlideAnimateGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.InSlideAnimateGroupLabel;
        }
        public string GetApplyAutoMotionThumbnailLabel(Office.IRibbonControl control)
        {
            return TextCollection.ApplyAutoMotionThumbnailLabel;
        }
        public string GetContextSpeakSelectedTextLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextSpeakSelectedTextLabel;
        }
        public string GetContextAddCurrentSlideLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextAddCurrentSlideLabel;
        }
        public string GetContextReplaceAudioLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextReplaceAudioLabel;
        }
        # endregion

        //Button Click Callbacks        
        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.AddAutoAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddAnimationButtonClick");
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
                PowerPointLabsGlobals.LogException(e, "ReloadAnimationButtonClick");
                throw;
            }
        }
        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            ZoomToArea.AddZoomToArea();
        }
        public void AboutButtonClick(Office.IRibbonControl control)
        {
            MessageBox.Show(TextCollection.AboutInfo, TextCollection.AboutInfoTitle);
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Process.Start(TextCollection.HelpDocumentUrl);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HelpButtonClick");
                throw;
            }
        }
        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Process.Start(TextCollection.FeedbackUrl);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "FeedbackButtonClick");
                throw;
            }
        }
        public void AddZoomInButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoZoom.AddDrillDownAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomInButtonClick");
                throw;
            }
        }
        public void AddZoomOutButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoZoom.AddStepBackAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomOutButtonClick");
                throw;
            }
        }

        # region Icon Getters
        public Bitmap GetAddAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAnimationImage");
                throw;
            }
        }
        public Bitmap GetReloadAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadAnimationImage");
                throw;
            }
        }
        public Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Spotlight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetSpotlightImage");
                throw;
            }
        }
        public Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadSpotlight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadSpotlightImage");
                throw;
            }
        }
        public Bitmap GetHighlightBulletsTextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightText);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsTextImage");
                throw;
            }
        }
        public Bitmap GetHighlightBulletsBackgroundImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightBackground);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsBackgroundImage");
                throw;
            }
        }

        public Bitmap GetHighlightWordsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightWords);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightWordsImage");
                throw;
            }
        }

        public Bitmap GetHighlightBulletsTextContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightTextContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsTextContextImage");
                throw;
            }
        }
        public Bitmap GetHighlightBulletsBackgroundContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightBackgroundContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsBackgroundContextImage");
                throw;
            }
        }

        public Bitmap GetZoomInImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomIn);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomInImage");
                throw;
            }
        }

        public Bitmap GetZoomOutImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomOut);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutImage");
                throw;
            }
        }
        public Bitmap GetZoomToAreaImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomToArea);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomToAreaImage");
                throw;
            }
        }
        public Bitmap GetZoomToAreaContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomToAreaContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomToAreaContextImage");
                throw;
            }
        }
        public Bitmap GetCropShapeImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.CutOutShape);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetCropShapeImage");
                throw;
            }
        }

        public Bitmap GetShapesLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ShapesLab);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetShapesLabImage");
                throw;
            }
        }
        public Bitmap GetColorsLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ColorsLab);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetColorsLabImage");
                throw;
            }
        }

        public Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.About);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAboutImage");
                throw;
            }
        }
        public Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Help);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHelpImage");
                throw;
            }
        }
        public Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Feedback);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetFeedbackImage");
                throw;
            }
        }
        public Bitmap GetAddAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAudio);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAudioImage");
                throw;
            }
        }
        public Bitmap GetRemoveAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.RemoveAudio);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetRemoveAudioImage");
                throw;
            }
        }
        public Bitmap GetAddCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddCaption);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddCaptionsImage");
                throw;
            }
        }
        public Bitmap GetRemoveCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.RemoveCaption);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetRemoveCaptionsImage");
                throw;
            }
        }

        public Bitmap GetAddAudioContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddNarrationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAudioContextImage");
                throw;
            }
        }
        public Bitmap GetPreviewNarrationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.SpeakTextContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetPreviewNarrationContextImage");
                throw;
            }
        }
        public Bitmap GetInSlideAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.InSlideAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetInSlideAnimationImage");
                throw;
            }
        }
        public Bitmap GetAddAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAnimationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAnimationContextImage");
                throw;
            }
        }
        public Bitmap GetReloadAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadAnimationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadAnimationContextImage");
                throw;
            }
        }
        public Bitmap GetAddSpotlightContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddSpotlightContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddSpotlightContextImage");
                throw;
            }
        }
        public Bitmap GetEditNameContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.EditNameContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetEditNameContextImage");
                throw;
            }
        }
        public Bitmap GetInSlideAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.InSlideContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetInSlideAnimationContextImage");
                throw;
            }
        }
        public Bitmap GetZoomInContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomInContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomInContextImage");
                throw;
            }
        }
        public Bitmap GetZoomOutContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomOutContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        public Bitmap GetAddToCustomShapeContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddToCustomShapes);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        # endregion

        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                BackgroundZoomChecked = pressed;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomStyleChanged");
                throw;
            }
        }
        public bool ZoomStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return BackgroundZoomChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }

        # region Control Enable
        public bool OnGetEnabledSpotlight(Office.IRibbonControl control)
        {
            return SpotlightEnabled;
        }
        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return ReloadSpotlight;
        }
        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return AddAutoMotionEnabled;
        }
        public bool OnGetEnabledReloadAutoMotion(Office.IRibbonControl control)
        {
            return ReloadAutoMotionEnabled;
        }
        public bool OnGetEnabledAddInSlide(Office.IRibbonControl control)
        {
            return InSlideEnabled;
        }
        public bool OnGetEnabledZoomButton(Office.IRibbonControl control)
        {
            return ZoomButtonEnabled;
        }
        public bool OnGetEnabledHighlightBullets(Office.IRibbonControl control)
        {
            return HighlightBulletsEnabled;
        }
        public bool OnGetEnabledRemoveCaptions(Office.IRibbonControl control)
        {
            return RemoveCaptionsEnabled;
        }
        public bool OnGetEnabledRemoveAudio(Office.IRibbonControl control)
        {
            return RemoveAudioEnabled;
        }
        public bool OnGetEnabledHighlightTextFragments(Office.IRibbonControl control)
        {
            return HighlightTextFragmentsEnabled;
        }
        # endregion

        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                var editForm = new Form1(this, selectedShape.Name);
                editForm.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "NameEditBtnClick");
                throw;
            }
        }
        public void ShapeNameEdited(String newName)
        {
            try
            {
                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        public void AutoAnimateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoAnimateDialogBox(DefaultDuration, FrameAnimationChecked);
                dialog.SettingsHandler += AnimationPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoAnimateDialogButtonPressed");
                throw;
            }
        }

        public void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            try
            {
                DefaultDuration = newDuration;
                FrameAnimationChecked = newFrameChecked;
                AnimateInSlide.defaultDuration = newDuration;
                AnimateInSlide.frameAnimationChecked = newFrameChecked;
                AutoAnimate.defaultDuration = newDuration;
                AutoAnimate.frameAnimationChecked = newFrameChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AnimationPropertiesEdited");
                throw;
            }
        }

        public void AutoZoomDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoZoomDialogBox(BackgroundZoomChecked, MultiSlideZoomChecked);
                dialog.SettingsHandler += ZoomPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoZoomDialogButtonPressed");
                throw;
            }
        }

        public void ZoomPropertiesEdited(bool backgroundChecked, bool multiSlideChecked)
        {
            try
            {
                BackgroundZoomChecked = backgroundChecked;
                MultiSlideZoomChecked = multiSlideChecked;
                AutoZoom.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.multiSlideZoomChecked = multiSlideChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomPropertiesEdited");
                throw;
            }
        }

        public void SpotlightDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new SpotlightDialogBox(Spotlight.defaultTransparency, Spotlight.defaultSoftEdges);
                dialog.SettingsHandler += SpotlightPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightDialogButtonPressed");
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
                PowerPointLabsGlobals.LogException(e, "SpotlightPropertiesEdited");
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
                HighlightTextFragments.backgroundColor = newBackgroundColor;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsPropertiesEdited");
                throw;
            }
        }
        public void HighlightBulletsDialogBoxPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new HighlightBulletsDialogBox(HighlightBulletsText.highlightColor, HighlightBulletsText.defaultColor, HighlightBulletsBackground.backgroundColor);
                dialog.SettingsHandler += HighlightBulletsPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsDialogBoxPressed");
                throw;
            }
        }

        public bool GetEmbedAudioVisiblity(Office.IRibbonControl control)
        {
            return EmbedAudioVisible;
        }

        private bool IsValidPresentation(PowerPoint.Presentation pres)
        {
            if (!Globals.ThisAddIn.VerifyVersion(pres))
            {
                MessageBox.Show(TextCollection.VersionNotCompatibleErrorMsg);
                return false;
            }

            if (!Globals.ThisAddIn.VerifyOnLocal(pres))
            {
                MessageBox.Show(TextCollection.OnlinePresentationNotCompatibleErrorMsg);
                return false;
            }

            return true;
        }

        private void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
        }

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

        #region Feature: Fit To Slide | Fit To Width | Fit To Height

        public void FitToWidthClick(Office.IRibbonControl control)
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToWidth(selectedShape);
        }

        public void FitToHeightClick(Office.IRibbonControl control)
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToHeight(selectedShape);
        }

        public Bitmap GetFitToWidthImage(Office.IRibbonControl control)
        {
            return FitToSlide.GetFitToWidthImage(control);
        }

        public Bitmap GetFitToHeightImage(Office.IRibbonControl control)
        {
            return FitToSlide.GetFitToHeightImage(control);
        }

        #endregion

        #region Feature: Crop to Shape

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            CropToShape.Crop(selection);
        }

        public Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            return CropToShape.GetCutOutShapeMenuImage(control);
        }

        #endregion

        #region Feature: Convert to Picture

        public void ConvertToPictureButtonClick(Office.IRibbonControl control)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            ConvertToPicture.Convert(selection);
        }

        public Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            return ConvertToPicture.GetConvertToPicMenuImage(control);
        }

        #endregion

        # region Feature: Combine Shapes
        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            const string officeVersion2010 = "14.0";
            return Globals.ThisAddIn.Application.Version == officeVersion2010;
        }
        # endregion

        # region Feature: Auto Narration
        public void AddAudioClick(Office.IRibbonControl control)
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

            if (PowerPointCurrentPresentationInfo.SelectedSlides.Any(slide => slide.NotesPageText.Trim() != ""))
            {
                RemoveAudioEnabled = true;
                RefreshRibbonControl("RemoveAudioButton");
            }

            var allAudioFiles = NotesToAudio.EmbedSelectedSlideNotes();

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null) return;

            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null) return;

            // initialize selected slides' audio
            recorder.InitializeAudioAndScript(PowerPointCurrentPresentationInfo.SelectedSlides.ToList(),
                                                  allAudioFiles, true);

            // if current list is visible, update the pane immediately
            if (recorderPane.Visible)
            {
                recorder.UpdateLists(currentSlide.ID);
            }

            PreviewAnimationsIfChecked();
        }

        public void AutoNarrateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoNarrateDialogBox(_voiceSelected, _voiceNames,
                    _previewCurrentSlide);
                dialog.SettingsHandler += AutoNarrateSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoNarrateDialogButtonPressed");
                throw;
            }
        }

        public void AutoNarrateSettingsChanged(String voiceName, bool previewCurrentSlide)
        {
            _previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                _voiceSelected = _voiceNames.IndexOf(voiceName);
            }
        }

        public void ContextAddAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.EmbedCurrentSlideNotes();
            PreviewAnimationsIfChecked();
        }

        public void ContextReplaceAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.ReplaceSelectedAudio();
        }

        public void RecManagementClick(Office.IRibbonControl control)
        {
            var currentPresentation = Globals.ThisAddIn.Application.ActivePresentation;

            if (!IsValidPresentation(currentPresentation))
            {
                return;
            }

            // prepare media files
            var tempPath = Globals.ThisAddIn.PrepareTempFolder(currentPresentation);
            Globals.ThisAddIn.PrepareMediaFiles(currentPresentation, tempPath);

            Globals.ThisAddIn.RegisterRecorderPane(currentPresentation.Windows[1], tempPath);

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            // if currently the pane is hidden, show the pane
            if (recorder != null && !recorderPane.Visible)
            {
                // fire the pane visble change event
                recorderPane.Visible = true;

                // reload the pane
                recorder.RecorderPaneReload();
            }
        }

        public void RemoveAudioClick(Office.IRibbonControl control)
        {
            try
            {
                NotesToAudio.RemoveAudioFromSelectedSlides();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null) return;

            recorder.ClearRecordDataListForSelectedSlides();

            // if current list is visible, update the pane immediately
            if (recorderPane.Visible)
            {
                foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides)
                {
                    recorder.UpdateLists(slide.ID);
                }
            }

            RemoveAudioEnabled = false;
            RefreshRibbonControl("RemoveAudioButton");
        }

        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }
        # endregion

        # region Feature: Auto Caption
        public void AddCaptionClick(Office.IRibbonControl control)
        {
            foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                if (slide.NotesPageText.Trim() != "")
                {
                    RemoveCaptionsEnabled = true;
                    break;
                }
            }

            NotesToCaptions.EmbedCaptionsOnSelectedSlides();
            RefreshRibbonControl("RemoveCaptionsButton");
        }

        public void RemoveCaptionClick(Office.IRibbonControl control)
        {
            RemoveCaptionsEnabled = false;
            RefreshRibbonControl("RemoveCaptionsButton");
            NotesToCaptions.RemoveCaptionsFromSelectedSlides();
        }

        public void RemoveAllNotesClick(Office.IRibbonControl control)
        {
            foreach (var slide in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                slide.NotesPageText = string.Empty;
            }
        }
        # endregion

        # region Feature: Shapes Lab
        public void CustomShapeButtonClick(Office.IRibbonControl control)
        {
            var prensentation = Globals.ThisAddIn.Application.ActivePresentation;
            
            Globals.ThisAddIn.InitializeShapesLabConfig();
            Globals.ThisAddIn.InitializeShapeGallery();
            Globals.ThisAddIn.RegisterShapesLabPane(prensentation);

            var customShapePane = Globals.ThisAddIn.GetActivePane(typeof(CustomShapePane));

            if (customShapePane == null || !(customShapePane.Control is CustomShapePane))
            {
                return;
            }

            var customShape = customShapePane.Control as CustomShapePane;

            Trace.TraceInformation(
                "Before Visible: " +
                string.Format("Pane Width = {0}, Pane Height = {1}, Control Width = {2}, Control Height {3}",
                              customShapePane.Width, customShapePane.Height, customShape.Width, customShape.Height));

            // if currently the pane is hidden, show the pane
            if (customShapePane.Visible)
            {
                return;
            }

            customShapePane.Visible = true;

            customShape.Width = customShapePane.Width - 16;
            customShape.PaneReload();
        }

        public void AddShapeButtonClick(Office.IRibbonControl control)
        {
            var prensentation = Globals.ThisAddIn.Application.ActivePresentation;

            Globals.ThisAddIn.InitializeShapesLabConfig();
            Globals.ThisAddIn.InitializeShapeGallery();
            Globals.ThisAddIn.RegisterShapesLabPane(prensentation);

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            var customShapePane = Globals.ThisAddIn.GetActivePane(typeof(CustomShapePane));

            if (customShapePane == null || !(customShapePane.Control is CustomShapePane))
            {
                return;
            }

            var customShape = customShapePane.Control as CustomShapePane;

            // show pane if not visible
            if (!customShapePane.Visible)
            {
                customShapePane.Visible = true;

                customShape.Width = customShapePane.Width - 16;
                customShape.PaneReload();
            }

            // first of all we check if the shape gallery has been opened correctly
            if (!Globals.ThisAddIn.ShapePresentation.Opened)
            {
                MessageBox.Show(TextCollection.ShapeGalleryInitErrorMsg);
                return;
            }

            // to determine if a presentation needs to be saved, we check 3 things:
            // 1. if the presentation is readonly;
            // 2. if the presentation contains a valid saving path;
            // 3. if the presentation has been saved.
            //
            // The only case we can save the presentation is:
            // The presentation is writable (readonly = false), contains a valid saving
            // path (valid path = true), and it has been saved (therefore all programmatical
            // changes can be saved without triggering a save dialog).
            var presentationSaved = prensentation.ReadOnly == Office.MsoTriState.msoFalse &&
                                    prensentation.Path != string.Empty &&
                                    prensentation.Saved == Office.MsoTriState.msoTrue;

            var shapeName = customShape.NextDefaultNameWithoutExtension;
            var shapeFullName = customShape.NextDefaultFullName;

            // add shape into shape gallery first to reduce flicker
            Globals.ThisAddIn.ShapePresentation.AddShape(selection, shapeName);

            // add the selection into pane and save it as .png locally
            ConvertToPicture.ConvertAndSave(selection, shapeFullName);

            // sync the shape among all opening panels
            Globals.ThisAddIn.SyncShapeAdd(shapeName, shapeFullName, customShape.CurrentCategory);

            // since we group and then ungroup the shape, document has been modified.
            // if the presentation has been saved before the group->ungroup, we can save
            // the file; else we leave it.
            if (presentationSaved)
            {
                Globals.ThisAddIn.Application.ActivePresentation.Save();
            }

            // finally, add the shape into the panel and waiting for name editing
            customShape.AddCustomShape(shapeName, shapeFullName, true);
        }
        # endregion

        #region Feature: Colors Lab
        public void ColorPickerButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.RegisterColorPane(Globals.ThisAddIn.Application.ActivePresentation);

                var colorPane = Globals.ThisAddIn.GetActivePane(typeof(ColorPane));

                // if currently the pane is hidden, show the pane
                if (!colorPane.Visible)
                {
                    // fire the pane visble change event
                    colorPane.Visible = true;
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Color Picker Failed", e.Message, e);
                PowerPointLabsGlobals.LogException(e, "ColorPickerButtonClicked");
                throw;
            }
        }
        #endregion

        # region Feature: Effects Lab
        public void MagnifyGlassEffectClick(Office.IRibbonControl control)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            if (selection.ShapeRange.Count > 1)
            {
                MessageBox.Show("Only one magnify area is allowed.");
                
                return;
            }

            var croppedShape = CropToShape.Crop(selection);

            croppedShape.Left -= 12;
            croppedShape.Top -= 12;

            croppedShape.ThreeD.BevelTopType = Office.MsoBevelType.msoBevelCircle;
            croppedShape.ThreeD.BevelBottomInset = 12;
            croppedShape.ThreeD.BevelBottomDepth = 3;
            croppedShape.ThreeD.BevelBottomType = Office.MsoBevelType.msoBevelNone;
            croppedShape.ThreeD.PresetLighting = Office.MsoLightRigType.msoLightRigBalanced;
            croppedShape.ThreeD.LightAngle = 145;

            croppedShape.ScaleHeight(1.4f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
            croppedShape.ScaleWidth(1.4f, Office.MsoTriState.msoTrue, Office.MsoScaleFrom.msoScaleFromMiddle);
        }

        public void BlurBackgroundEffectClick(Office.IRibbonControl control)
        {
            BackgroundManipulation(null);
        }

        public void GreyScaleBackgroundEffectClick(Office.IRibbonControl control)
        {
            BackgroundManipulation(MatrixFilters.GreyScale);
        }

        public void BlackWhiteBackgroundEffectClick(Office.IRibbonControl control)
        {
            BackgroundManipulation(MatrixFilters.BlackWhite);
        }

        public void GohamBackgroundEffectClick(Office.IRibbonControl control)
        {
            BackgroundManipulation(MatrixFilters.Gotham);
        }

        public void SepiaBackgroundEffectClick(Office.IRibbonControl control)
        {
            BackgroundManipulation(MatrixFilters.Sepia);
        }

        public void TransparentEffectClick(Office.IRibbonControl control)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            TransparentEffect(selection.ShapeRange);
        }

        private void BackgroundManipulation(IMatrixFilter filter)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            
            // soften cropped shape's edge
            selection.ShapeRange.SoftEdge.Type = Office.MsoSoftEdgeType.msoSoftEdgeType5;

            var croppedShape = CropToShape.Crop(selection);

            if (croppedShape == null) return;

            croppedShape.Left -= 12;
            croppedShape.Top -= 12;

            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;

            var picSaveTempPath = Path.Combine(Path.GetTempPath(), "slide.png");

            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory.Load(picSaveTempPath);
                
                if (filter == null)
                {
                    image = image.GaussianBlur(20);
                }
                else
                {
                    image = image.Filter(filter);
                }

                image.Save(picSaveTempPath);
            }

            var newPic = shapes.AddPicture(picSaveTempPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                           0, 0,
                                           PowerPointCurrentPresentationInfo.SlideWidth,
                                           PowerPointCurrentPresentationInfo.SlideHeight);

            while (newPic.ZOrderPosition > croppedShape.ZOrderPosition)
            {
                newPic.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }
        }

        private void TransparentEffect(PowerPoint.ShapeRange shapeRange)
        {
            foreach (PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subShapeRange = shape.Ungroup();
                    TransparentEffect(subShapeRange);
                    subShapeRange.Group();
                } else
                if (shape.Type == Office.MsoShapeType.msoPicture)
                {
                    PictureTransparencyHandler(shape);
                } else
                if (IsTransparentableShape(shape))
                {
                    ShapeTransparencyHandler(shape);
                }
            }
        }

        private bool IsTransparentableShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape ||
                   shape.Type == Office.MsoShapeType.msoFreeform;

        }

        private void PictureTransparencyHandler(PowerPoint.Shape picture)
        {
            var tempPicPath = Path.Combine(Path.GetTempPath(), "tempPic.png");

            picture.Export(tempPicPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0,
                           PowerPoint.PpExportMode.ppScaleXY);

            var shapeHolder =
                PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    picture.Left,
                    picture.Top,
                    picture.Width,
                    picture.Height);

            var oriZOrder = picture.ZOrderPosition;

            picture.Delete();

            // move shape holder to original z-order
            while (shapeHolder.ZOrderPosition > oriZOrder)
            {
                shapeHolder.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }

            shapeHolder.Line.Visible = Office.MsoTriState.msoFalse;
            shapeHolder.Fill.UserPicture(tempPicPath);
            shapeHolder.Fill.Transparency = 0.5f;

            File.Delete(tempPicPath);
        }

        private void ShapeTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Fill.Transparency = 0.5f;
            shape.Line.Transparency = 0.5f;
        }

        # endregion

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
    }
}
