﻿using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Factory;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.PictureSlidesLab.Views;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


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
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "Migration to Action Framework")]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        #region Action Framework Factory
        private ActionHandlerFactory ActionHandlerFactory { get; set; }

        private EnabledHandlerFactory EnabledHandlerFactory { get; set; }

        private LabelHandlerFactory LabelHandlerFactory { get; set; }

        private ImageHandlerFactory ImageHandlerFactory { get; set; }

        private SupertipHandlerFactory SupertipHandlerFactory { get; set; }

        private ContentHandlerFactory ContentHandlerFactory { get; set; }

        private PressedHandlerFactory PressedHandlerFactory { get; set; }

        private CheckBoxActionHandlerFactory CheckBoxActionHandlerFactory { get; set; }
        #endregion

        #region Action Framework entry point

        public void OnAction(Office.IRibbonControl control)
        {
            var actionHandler = ActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            actionHandler.Execute(control.Id);
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            var enabledHandler = EnabledHandlerFactory.CreateInstance(control.Id, control.Tag);
            return enabledHandler.Get(control.Id);
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            var labelHandler = LabelHandlerFactory.CreateInstance(control.Id, control.Tag);
            return labelHandler.Get(control.Id);
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            var supertipHandler = SupertipHandlerFactory.CreateInstance(control.Id, control.Tag);
            return supertipHandler.Get(control.Id);
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            var imageHandler = ImageHandlerFactory.CreateInstance(control.Id, control.Tag);
            return imageHandler.Get(control.Id);
        }

        public string GetContent(Office.IRibbonControl control)
        {
            var contentHandler = ContentHandlerFactory.CreateInstance(control.Id, control.Tag);
            return contentHandler.Get(control.Id);
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            var pressedHandler = PressedHandlerFactory.CreateInstance(control.Id, control.Tag);
            return pressedHandler.Get(control.Id);
        }

        public void OnCheckBoxAction(Office.IRibbonControl control, bool pressed)
        {
            var checkBoxActionHandler = CheckBoxActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            checkBoxActionHandler.Execute(control.Id, pressed);
        }

        #endregion

        #region Deprecated. Please only use Action Framework to support the feature.

#pragma warning disable 0618
        private Office.IRibbonUI _ribbon;

        public bool FrameAnimationChecked = false;
        public bool SpotlightDelete = true;
        public float DefaultDuration = 0.5f;

        public bool HighlightBulletsEnabled = true;
        public bool AddAutoMotionEnabled = true;
        public bool ReloadAutoMotionEnabled = true;
        public bool ReloadSpotlight = true;
        public bool RemoveCaptionsEnabled = true;
        public bool RemoveAudioEnabled = true;

        public bool HighlightTextFragmentsEnabled = true;

        public bool EmbedAudioVisible = true;
        public bool RecorderPaneVisible = false;

        public bool _previewCurrentSlide;

        public List<string> _voiceNames;

        public int _voiceSelected;

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
            ActionHandlerFactory = new ActionHandlerFactory();
            EnabledHandlerFactory = new EnabledHandlerFactory();
            LabelHandlerFactory = new LabelHandlerFactory();
            SupertipHandlerFactory = new SupertipHandlerFactory();
            ImageHandlerFactory = new ImageHandlerFactory();
            ContentHandlerFactory = new ContentHandlerFactory();
            PressedHandlerFactory = new PressedHandlerFactory();
            CheckBoxActionHandlerFactory = new CheckBoxActionHandlerFactory();

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
                Logger.LogException(e, "RefreshRibbonControl");
                throw;
            }
        }

        private void SetVoicesFromInstalledOptions()
        {
            var installedVoices = NotesToAudio.GetVoices().ToList();
            _voiceNames = installedVoices;
        }

        #region Button Labels
        public string GetPowerPointLabsAddInsTabLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsAddInsTabLabel;
        }

        public string GetCombineShapesLabel(Office.IRibbonControl control)
        {
            return TextCollection.CombineShapesLabel;
        }

        public string GetPowerPointLabsMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsMenuLabel;
        }
        # endregion

        #region Control Enable
        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return ReloadSpotlight;
        }
        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return AddAutoMotionEnabled;
        }
        # endregion

        //Edit Name Callbacks
        public bool GetEmbedAudioVisiblity(Office.IRibbonControl control)
        {
            return EmbedAudioVisible;
        }

        public bool IsValidPresentation(PowerPoint.Presentation pres)
        {
            if (!Globals.ThisAddIn.VerifyVersion(pres))
            {
                MessageBox.Show(TextCollection.VersionNotCompatibleErrorMsg);
                return false;
            }

            return true;
        }

        public void PreviewAnimationsIfChecked()
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
            catch (IndexOutOfRangeException e)
            {
                // No voices are installed (It should be impossible for the index to be out of range otherwise.)
                Logger.LogException(e, "GetSelectedVoiceOrNull");
            }
            return selectedVoice;
        }

        public Bitmap GetContextMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.PptlabsContextMenu);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetContextMenuImage");
                throw;
            }
        }

        #endregion

        #region Feature: Picture Slides Lab

        public PictureSlidesLabWindow PictureSlidesLabWindow { get; set; }

        #endregion

        #region Feature: Combine Shapes
        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            const string officeVersion2010 = "14.0";
            return Globals.ThisAddIn.Application.Version == officeVersion2010;
        }
        #endregion

        #region Feature: Narrations Lab
        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }
        #endregion

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; i++)
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
