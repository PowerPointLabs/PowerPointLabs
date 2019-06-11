using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.FunctionalTestInterface.Windows;
using PowerPointLabs.ZoomLab.Views;
using TestInterface;
using TestInterface.Windows;

namespace Test.Util
{
    static class WPFWindowUtil
    {
        public static IMarshalWindow WaitAndPush<T>(this IWindowStackManager windowStackManager,
            Action action, uint processId, string name = null, int timeout = 5000)
            where T : DispatcherObject
        {
            WindowOpenTrigger trigger = new WindowOpenTrigger(false);
            AutomationEventHandler handler = GetOpenWindowHandler(processId, name, trigger);

            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Children,
                handler);

            Task task = new Task(action);
            task.Start();
            trigger.Wait(timeout);

            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                handler);

            Assert.IsTrue(trigger.IsSet, $"Timeout of {timeout}ms has been reached.");
            Assert.AreNotEqual(trigger.resultingWindow, IntPtr.Zero, "Found null window handle");
            IMarshalWindow window = windowStackManager.Push(trigger.resultingWindow);
            Assert.IsNotNull(window);
            Assert.IsTrue(window.IsType<T>());
            Assert.IsTrue(name == null || name == window.Title);
            trigger.Dispose();
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

        public static void SetZoomProperties(this IPowerPointLabsFeatures PplFeatures, IPowerPointOperations PpOperations,
            bool backgroundChecked, bool multiSlideChecked)
        {
            string zoomLabSettingsWindowTitle = "Zoom Lab Settings";
            IMarshalWindow window = PpOperations.WindowStackManager.WaitAndPush<ZoomLabSettingsDialogBox>(
                PplFeatures.OpenZoomLabSettings,
                PpOperations.ProcessId,
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

        private static AutomationEventHandler GetOpenWindowHandler(uint processId, string name, WindowOpenTrigger trigger)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);
                if ((uint)element.Current.ProcessId == processId
                         && (name == null || windowName == name))
                {
                    trigger.resultingWindow = handle;
                    trigger.Set();
                }
            };
        }

    }
}
