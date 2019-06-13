using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.FunctionalTestInterface.Windows;

namespace Test.Util
{
    public static class WindowWatcher
    {
        private static Dictionary<string, WindowOpenTrigger> windowTriggers;
        private static AutomationEventHandler handler;

        public static void Setup(uint processId)
        {
            windowTriggers = new Dictionary<string, WindowOpenTrigger>();
            handler = GetOpenWindowHandler(processId);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Children,
                handler);
        }

        public static void Teardown()
        {
            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                handler);
            handler = null;
            windowTriggers = null;
        }

        public static IntPtr Push(string name, Action action, int timeout = 5000)
        {
            WindowOpenTrigger trigger = new WindowOpenTrigger(false);
            windowTriggers.Add(name, trigger);
            Task task = new Task(action);
            task.Start();
            trigger.Wait(timeout);
            windowTriggers.Remove(name);
            Assert.IsTrue(trigger.IsSet, $"Timeout of {timeout}ms has been reached.");
            Assert.AreNotEqual(trigger.resultingWindow, IntPtr.Zero, "Found null window handle");
            return trigger.resultingWindow;
        }

        private static AutomationEventHandler GetOpenWindowHandler(uint processId)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                if ((uint)element.Current.ProcessId != processId) { return; }

                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);

                WindowOpenTrigger resultTrigger = windowTriggers.FirstOrDefault(o => !o.Value.IsSet && o.Key == windowName).Value;
                if (resultTrigger == null)
                {
                    WindowUtil.CloseWindow(handle);
                    return;
                }
                resultTrigger.resultingWindow = handle;
                resultTrigger.Set();
            };
        }
    }
}
