using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using Test.Util.Windows;

namespace Test.Util
{
    public class WindowStackManager
    {
        private static Stack<Window> windowStack;

        public static void Setup()
        {
            windowStack = new Stack<Window>();
        }

        public static void Teardown()
        {
            Automation.RemoveAllEventHandlers();
            while (Peek() != null)
            {
                Pop();
            }
        }

        /// <summary>
        /// Opens a new window asynchronously and waits for the get the window handle.
        /// </summary>
        public static Window Push(Action action, uint processId, string name = null, int timeout = 5000)
        {
            WindowOpenTrigger trigger = new WindowOpenTrigger(false);
            Automation.AddAutomationEventHandler(
            WindowPattern.WindowOpenedEvent,
            AutomationElement.RootElement,
            TreeScope.Children, GetOpenWindowHandler<Window>(processId, name, trigger));
            Task task = new Task(action);
            task.Start();
            trigger.Wait(timeout);
            if (!trigger.IsSet)
            {
                return null; //can throw exception when receive null
            }
            windowStack.Push(trigger.resultingWindow);
            return trigger.resultingWindow;
        }

        public static Window Peek()
        {
            return windowStack.Peek();
        }

        public static void Pop(bool close = true)
        {
            Window w = windowStack.Pop();
            if (close)
            {
                w.Close();
            }
        }

        private static AutomationEventHandler GetOpenWindowHandler<T>(uint processId, string name, WindowOpenTrigger trigger) where T : Window
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                if ((uint)element.Current.ProcessId == processId
                         && (name == null || WindowUtil.GetWindowTitle(element.Current.NativeWindowHandle) == name))
                {
                    return;
                }
                trigger.resultingWindow = null;
                trigger.Set();
            };
        }
    }
}
