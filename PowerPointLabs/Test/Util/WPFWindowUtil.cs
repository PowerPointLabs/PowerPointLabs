using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util.Windows;
using TestInterface.Windows;

namespace Test.Util
{
    static class WPFWindowUtil
    {
        public static IMarshalWPF WaitAndPush<T>(this IWindowStackManager windowStack,
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

            if (!trigger.IsSet)
            {
                trigger.Dispose();
                Assert.Fail($"Timeout of {timeout}ms has been reached.");
                return null;
            }
            trigger.Dispose();
            if (new IntPtr(trigger.resultingWindow) == IntPtr.Zero)
            {
                Assert.Fail("Found null window handle");
                return null;
            }
            IMarshalWPF result = windowStack.Push<T>(trigger.resultingWindow, trigger.name);
            return result;
        }

        // do a wait and pop later
        // can also do negatives

        private static AutomationEventHandler GetOpenWindowHandler(uint processId, string name, WindowOpenTrigger trigger)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                int handle = element.Current.NativeWindowHandle;
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
