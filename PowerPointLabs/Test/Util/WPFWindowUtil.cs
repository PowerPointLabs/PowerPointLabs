﻿using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ZoomLab.Views;
using TestInterface;
using TestInterface.Windows;

namespace Test.Util
{
    static class WPFWindowUtil
    {
        public static string After(this string original, string searchTerm)
        {
            int index;
            return (index = original.LastIndexOf(searchTerm)) != -1 ? original.Substring(index + searchTerm.Length) : "";
        }

        public static IMarshalWindow WaitAndPush<T>(this IWindowStackManager windowStackManager,
            Action action, string name, int timeout = 5000)
            where T : DispatcherObject
        {
            IntPtr handle = WindowWatcher.Push(name, action, timeout);
            IMarshalWindow window = windowStackManager.Push(handle);
            Assert.IsNotNull(window);
            Assert.IsTrue(window.IsType<T>());
            Assert.IsTrue(name == null || name == window.Title);
            return window;
        }

        public static void NativeClick<T>(this IMarshalWindow window, string name)
        {
            Point p = window.GetPosition<T>(name);
            MouseUtil.SendMouseLeftClick((int)p.X, (int)p.Y);
        }

        public static void NativeClickList<T>(this IMarshalWindow window, string name, int index)
        {
            Point p = window.GetListElementPosition<T>(name, index);
            MouseUtil.SendMouseLeftClick((int)p.X, (int)p.Y);
        }

        public static void SetZoomProperties(this IPowerPointLabsFeatures PplFeatures, IWindowStackManager windowStackManager,
            bool backgroundChecked, bool multiSlideChecked)
        {
            string zoomLabSettingsWindowTitle = "Zoom Lab Settings";
            IMarshalWindow window = windowStackManager.WaitAndPush<ZoomLabSettingsDialogBox>(
                PplFeatures.OpenZoomLabSettings,
                zoomLabSettingsWindowTitle);
            window.SetCheckBox<ZoomLabSettingsDialogBox>("slideBackgroundCheckbox", backgroundChecked);
            window.SetCheckBox<ZoomLabSettingsDialogBox>("separateSlidesCheckbox", multiSlideChecked);
            window.NativeClick<ZoomLabSettingsDialogBox>("okButton");
        }

        public static void SetCheckBox<T>(this IMarshalWindow window, string name, bool isChecked)
        {
            if (window.IsChecked<T>(name) == isChecked)
            {
                return;
            }
            window.NativeClick<ZoomLabSettingsDialogBox>(name);
        }

        public static bool IsRunning(this Process process)
        {
            if (process == null)
                throw new ArgumentNullException("process");

            try
            {
                Process.GetProcessById(process.Id);
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
            return true;
        }
    }
}
