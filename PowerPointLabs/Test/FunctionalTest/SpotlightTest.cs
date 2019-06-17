using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.EffectsLab.Views;
using PowerPointLabs.FunctionalTestInterface.Windows;
using Test.Util;
using TestInterface.Windows;

namespace Test.FunctionalTest
{
    [TestClass]
    public class SpotlightTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "EffectsLab\\Spotlight.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_Spotlight()
        {
            VariousMultipleShapesSuccessfully();
            SettingsAndSingleShapeSuccessfully();

        }

        private void SettingsAndSingleShapeSuccessfully()
        {
            // old code partially replaced
            //PplFeatures.SetSpotlightProperties(0.01f, 50f, Color.FromArgb(0x00FF00));
            PplFeatures.SetSpotlightProperties(1.0f, 10f, Color.FromArgb(0x00FF00));
            VerifySpotlightSettingsDialogBoxWPF();

            PpOperations.SelectSlide(4);
            PpOperations.SelectShape("Spotlight Me");
            PplFeatures.Spotlight();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide1 = PpOperations.SelectSlide(4);
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide2 = PpOperations.SelectSlide(5);
            Microsoft.Office.Interop.PowerPoint.Slide expSlide1 = PpOperations.SelectSlide(6);
            Microsoft.Office.Interop.PowerPoint.Slide expSlide2 = PpOperations.SelectSlide(7);
            SlideUtil.IsSameLooking(expSlide1, actualSlide1);
            SlideUtil.IsSameLooking(expSlide2, actualSlide2);
        }

        private void VariousMultipleShapesSuccessfully()
        {
            PpOperations.SelectSlide(8);
            PpOperations.SelectShapes(new List<String>
            {
                "Rectangle 3",
                "Flowchart: Document 5",
                "Freeform 17",
                "Group 9",
                "Line Callout 1 (Border and Accent Bar) 11",
                "Freeform 1",
                "Flowchart: Alternate Process 16",
                "Rectangle 4"
            });

            PplFeatures.Spotlight();

            Microsoft.Office.Interop.PowerPoint.Slide actualSlide1 = PpOperations.SelectSlide(8);
            Microsoft.Office.Interop.PowerPoint.Slide actualSlide2 = PpOperations.SelectSlide(9);
            Microsoft.Office.Interop.PowerPoint.Slide expSlide1 = PpOperations.SelectSlide(10);
            PpOperations.SelectShape("text 3")[1].Delete();
            Microsoft.Office.Interop.PowerPoint.Slide expSlide2 = PpOperations.SelectSlide(11);
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameLooking(expSlide1, actualSlide1);
            SlideUtil.IsSameLooking(expSlide2, actualSlide2);
        }

        private void VerifySpotlightSettingsDialogBoxWPF()
        {
            string spotlightSettingsWindowTitle = "Spotlight Settings";
            IMarshalWindow window = WindowStackManager.WaitAndPush<SpotlightSettingsDialogBox>(
                PplFeatures.OpenSpotlightDialog,
                spotlightSettingsWindowTitle);

            window.LeftClick<SpotlightSettingsDialogBox>("spotlightTransparencyInput");
            window.SelectAll<SpotlightSettingsDialogBox>("spotlightTransparencyInput");
            window.TypeUsingKeyboard<SpotlightSettingsDialogBox>("spotlightTransparencyInput", "1");

            // scrolling down doesn't work for a dropdown, it's not a combobox!
            ThreadUtil.WaitFor(1000);
            window.NativeClick<SpotlightSettingsDialogBox>("softEdgesSelectionInput");
            ThreadUtil.WaitFor(1000);
            window.NativeClickList<SpotlightSettingsDialogBox>("softEdgesSelectionInput", 7);
            // The WPF element listens for mouse down/up events so this is required.
            window.LeftClick<SpotlightSettingsDialogBox>("softEdgesSelectionInput");
            ThreadUtil.WaitFor(1000);
            window.NativeClick<SpotlightSettingsDialogBox>("okButton");

            // spotlightColorRect is not supported as it uses WinForms
        }

        [Obsolete]
        private void VerifySpotlightSettingsDialogBox()
        {
            string spotlightSettingsWindowTitle = "Spotlight Settings";

            DialogUtil.WaitForDialogBox(PplFeatures.OpenSpotlightDialog, null, spotlightSettingsWindowTitle);
            IntPtr spotlightDialog = NativeUtil.FindWindow(null, spotlightSettingsWindowTitle);
            Assert.AreNotEqual(IntPtr.Zero, spotlightDialog, "Failed to find " + spotlightSettingsWindowTitle + ".");

            // In Win7, it's "25 %", but in Win10, it's "25%"
            IntPtr transparencyDialog = NativeUtil.FindWindowEx(spotlightDialog, IntPtr.Zero, null, "25 %");
            if (transparencyDialog == IntPtr.Zero)
            {
                transparencyDialog = NativeUtil.FindWindowEx(spotlightDialog, IntPtr.Zero, null, "25%");
            }
            Assert.AreNotEqual(IntPtr.Zero, transparencyDialog, "Failed to find Text Dialog.");

            // Set text
            NativeUtil.SendMessage(transparencyDialog, 0x000C /*WM_SETTEXT*/, IntPtr.Zero, "1");

            // try to get class's build id
            StringBuilder actualContentBuilder = new StringBuilder(1024);
            NativeUtil.GetClassName(spotlightDialog, actualContentBuilder, 1024);
            string classBuildId = actualContentBuilder.ToString().Split('.').Last();

            IntPtr fadeComboBox = NativeUtil.FindWindowEx(spotlightDialog, IntPtr.Zero, "WindowsForms10.COMBOBOX.app.0." + classBuildId, null);
            Assert.AreNotEqual(IntPtr.Zero, fadeComboBox, "Failed to find Fade Dialog.");

            StringBuilder sb = new StringBuilder(256, 256);
            NativeUtil.SendMessage(fadeComboBox, 0x0148 /*CB_GETLBTEXT*/, (IntPtr)2, sb);

            // Set combo box
            NativeUtil.SendMessage(fadeComboBox, 0x014E /*CB_SETCURSEL*/, IntPtr.Zero, sb.ToString());
            NativeUtil.SendMessage(fadeComboBox, 0x100 /*WM_KEYDOWN*/, (IntPtr)Keys.Down, IntPtr.Zero);
            NativeUtil.SendMessage(fadeComboBox, 0x100 /*WM_KEYDOWN*/, (IntPtr)Keys.Down, IntPtr.Zero);
            NativeUtil.SendMessage(fadeComboBox, 0x100 /*WM_KEYDOWN*/, (IntPtr)Keys.Down, IntPtr.Zero);
            NativeUtil.SendMessage(fadeComboBox, 0x100 /*WM_KEYDOWN*/, (IntPtr)Keys.Down, IntPtr.Zero);
            NativeUtil.SendMessage(fadeComboBox, 0x100 /*WM_KEYDOWN*/, (IntPtr)Keys.Down, IntPtr.Zero);
            NativeUtil.SendMessage(fadeComboBox, 0x100 /*WM_KEYDOWN*/, (IntPtr)Keys.Down, IntPtr.Zero);

            DialogUtil.CloseDialogBox(spotlightDialog, "OK");
        }
    }
}
