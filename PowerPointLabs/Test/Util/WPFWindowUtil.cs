using System;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util.Windows;
using TestInterface.Windows;

namespace Test.Util
{
    static class WPFWindowUtil
    {
        public static MarshalWindow WaitAndPush<T>(this IWindowStackManager windowStackManager,
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
            MarshalWindow window = windowStackManager.Push(trigger.resultingWindow);
            Assert.IsNotNull(window);
            Assert.IsTrue(window.IsType<T>());
            Assert.IsTrue(name == null || name == window.Title);
            trigger.Dispose();
            return window;
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
                    trigger.name = windowName;
                    trigger.Set();
                }
            };
        }

    }
}
