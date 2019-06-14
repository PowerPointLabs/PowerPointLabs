using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.FunctionalTestInterface.Windows;

namespace Test.Util
{
    public static class WindowWatcher
    {
        private static Dictionary<string, WindowOpenTrigger> windowTriggers;
        private static AutomationEventHandler handler;
        private static HashSet<string> whitelist;
        private static SortedDictionary<string, WindowOpenTrigger> whitelistInstances;
        private static string lastOpenWindowName = "";

        public static void AddToWhitelist(string name)
        {
            if (!whitelist.Contains(name))
            {
                whitelist.Add(name);
            }
        }

        public static void Setup(Process process, Process childProcess, string startWindowName, int timeout = 10000)
        {
            if (process.IsRunning())
            {
                WaitForPresentation(process.Id, childProcess, startWindowName, timeout);
            }
            else
            {
                WaitForStartup(process, startWindowName, timeout);
            }

            windowTriggers = new Dictionary<string, WindowOpenTrigger>();
            whitelist = new HashSet<string>();
            whitelistInstances = new SortedDictionary<string, WindowOpenTrigger>();
            handler = GetOpenWindowHandler(process.Id);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Descendants,
                handler);
        }

        public static void Teardown()
        {
            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                handler);
            whitelistInstances = null;
            whitelist = null;
            handler = null;
            windowTriggers = null;
        }

        public static IntPtr Push(string name, Action action, int timeout = 5000)
        {
            WindowOpenTrigger whitelistTrigger = whitelistInstances.FirstOrDefault(o => o.Key == name).Value;
            if (whitelistTrigger != null)
            {
                whitelistInstances.Remove(name);
                return whitelistTrigger.resultingWindow;
            }
            WindowOpenTrigger trigger = new WindowOpenTrigger(false);
            windowTriggers.Add(name, trigger);
            Task task = new Task(action);
            task.Start();
            trigger.Wait(timeout);
            windowTriggers.Remove(name);
            Assert.IsTrue(trigger.IsSet, $"Timeout of {timeout}ms has been reached.{lastOpenWindowName}");
            Assert.AreNotEqual(trigger.resultingWindow, IntPtr.Zero, "Found null window handle");
            return trigger.resultingWindow;
        }

        private static void WaitForPresentation(int processId, Process childProcess, string startWindowName, int timeout)
        {
            WindowOpenTrigger trigger = new WindowOpenTrigger(false);
            AutomationPropertyChangedEventHandler handler = GetPresentationOpenHandler(processId, startWindowName, trigger);
            Automation.AddAutomationPropertyChangedEventHandler(
                AutomationElement.RootElement,
                TreeScope.Descendants,
                handler,
                AutomationElement.NameProperty);
            childProcess.Start();
            trigger.Wait(timeout);
            Automation.RemoveAutomationPropertyChangedEventHandler(
                AutomationElement.RootElement,
                handler);
            Assert.IsTrue(trigger.IsSet, $"Failed to load presentation in {timeout}ms.");
        }

        private static void WaitForStartup(Process process, string startWindowName, int timeout)
        {
            WindowOpenTrigger openTrigger = new WindowOpenTrigger(false);
            AutomationEventHandler openHandler = GetStartingWindowHandler(process, startWindowName, openTrigger);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Descendants,
                openHandler);

            process.Start();

            openTrigger.Wait(timeout);
            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                openHandler);
            Assert.IsTrue(openTrigger.IsSet, $"Failed to find starting window {timeout}ms.");
        }

        private static void WaitForStartupOld(Process process, string loaderName, int timeout)
        {
            if (process.IsRunning()) return;

            WindowOpenTrigger openTrigger = new WindowOpenTrigger(false);
            AutomationEventHandler openHandler = GetLoadWindowOpenHandler(process, loaderName, openTrigger);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Descendants,
                openHandler);

            process.Start();

            openTrigger.Wait(timeout);
            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                openHandler);
            Assert.IsTrue(openTrigger.IsSet, $"Failed to find loader in {timeout}ms. {lastOpenWindowName}");
            openTrigger.Reset();
            AutomationEventHandler closedHandler = GetLoadWindowClosedHandler(openTrigger);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowClosedEvent,
                AutomationElement.RootElement,
                TreeScope.Subtree,
                closedHandler);
            openTrigger.Wait(timeout);
            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowClosedEvent,
                AutomationElement.RootElement,
                closedHandler);
            Assert.IsTrue(openTrigger.IsSet, $"Loader failed to close in {timeout}ms.");
        }

        private static AutomationPropertyChangedEventHandler GetPresentationOpenHandler(int processId, string startWindowName, WindowOpenTrigger trigger)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                if (element.Current.ProcessId != processId) { return; }

                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);

                if (windowName == startWindowName)
                {
                    trigger.runtimeId = element.GetRuntimeId();
                    trigger.Set();
                }
            };
        }

        private static AutomationEventHandler GetStartingWindowHandler(Process process, string startWindowName, WindowOpenTrigger trigger)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                if (element.Current.ProcessId != process.Id) { return; }

                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);
                lastOpenWindowName = windowName;

                if (windowName == startWindowName)
                {
                    trigger.runtimeId = element.GetRuntimeId();
                    trigger.Set();
                }
            };
        }

        private static AutomationEventHandler GetLoadWindowOpenHandler(Process process, string loaderName, WindowOpenTrigger trigger)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                try
                {
                    if (element.Current.ProcessId != process.Id) { return; }
                }
                catch (Exception)
                {
                    // for now dont do anything
                    return;
                }
                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);
                //MessageBox.Show(windowName);
                lastOpenWindowName = windowName;

                if (windowName == loaderName)
                {
                    trigger.runtimeId = element.GetRuntimeId();
                    trigger.Set();
                }
            };
        }

        private static AutomationEventHandler GetLoadWindowClosedHandler(WindowOpenTrigger trigger)
        {
            return (sender, e) =>
            {
                WindowClosedEventArgs windowEventArgs = (WindowClosedEventArgs)e;
                if (Automation.Compare(trigger.runtimeId, windowEventArgs.GetRuntimeId()))
                {
                    trigger.Set();
                }
            };
        }

        private static AutomationEventHandler GetOpenWindowHandler(int processId)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                if (element.Current.ProcessId != processId &&
                Process.GetProcessById(element.Current.ProcessId).ProcessName != "POWERPNT") { return; }

                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);
                lastOpenWindowName = windowName;

                WindowOpenTrigger resultTrigger = GetWindowTrigger(windowName);
                if (resultTrigger == null)
                {
                    WindowUtil.CloseWindow(handle);
                    return;
                }
                resultTrigger.resultingWindow = handle;
                resultTrigger.Set();
            };
        }

        private static WindowOpenTrigger GetWindowTrigger(string windowName)
        {
            WindowOpenTrigger trigger = windowTriggers.FirstOrDefault(o => !o.Value.IsSet && o.Key == windowName).Value;
            if (trigger == null && whitelist.Contains(windowName))
            {
                trigger = new WindowOpenTrigger(false);
                whitelistInstances.Add(windowName, trigger);
            }
            return trigger;
        }
    }
}
