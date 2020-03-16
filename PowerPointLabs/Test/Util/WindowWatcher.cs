using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ActionFramework.Common.Log;
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
        private static Process process;
        private static string processName;
        private static Stack<Task> tasks;

        public static void AddToWhitelist(string name)
        {
            if (!whitelist.Contains(name))
            {
                whitelist.Add(name);
            }
        }

        public static void RevalidateApp()
        {
            if (process == null)
            {
                // headless
                return;
            }
            try
            {
                process.WaitForInputIdle();
            }
            catch
            {
                string tempName = processName;
                Teardown(false);
                process = GetProcess(tempName);
                Setup(process, process, tempName);
            }
        }

        public static void Setup(Process process, PPTProcessWrapper childProcessWrapper, string processName, int timeout = 10000)
        {
            SetupWhitelist();
            Process childProcess = childProcessWrapper.Start();
            if (process == null)
            {
                process = childProcess;
            }
            StartProcessAndStartWindowWatching(process, childProcess, processName);
        }

        public static void Setup(Process process, Process childProcess, string processName, int timeout = 10000)
        {
            SetupWhitelist();
            childProcess.Start();
            StartProcessAndStartWindowWatching(process, childProcess, processName);
        }

        private static void StartProcessAndStartWindowWatching(Process process, Process childProcess, string processName)
        {
            childProcess.WaitForInputIdle();
            WindowWatcher.process = process;
            WindowWatcher.processName = processName;
            RevalidateApp();

            handler = GetOpenWindowHandler(process.Id);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Descendants,
                handler);
        }

        private static Process GetProcess(string processName)
        {
            Process process = null;
            int retries = 5;
            while (process == null && retries > 0)
            {
                Process[] p = Process.GetProcessesByName(processName);
                if (p.Count() != 0)
                {
                    process = p[0];
                    break;
                }
                retries--;
            }
            process.WaitForInputIdle();
            return process;
        }

        private static void SetupWhitelist()
        {
            windowTriggers = new Dictionary<string, WindowOpenTrigger>();
            if (whitelist == null)
            {
                whitelist = new HashSet<string>();
            }
            whitelistInstances = new SortedDictionary<string, WindowOpenTrigger>();
            tasks = new Stack<Task>();
        }

        public static void HeadlessSetup(int processId)
        {
            SetupWhitelist();
            handler = GetOpenWindowHandler(processId);
            Automation.AddAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                TreeScope.Descendants,
                handler);
        }

        public static void Teardown(bool eraseWhiteList = true)
        {
            Automation.RemoveAutomationEventHandler(
                WindowPattern.WindowOpenedEvent,
                AutomationElement.RootElement,
                handler);
            whitelistInstances = null;
            if (eraseWhiteList)
            {
                whitelist = null;
            }
            tasks = null;
            handler = null;
            windowTriggers = null;
            process = null;
            processName = null;
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
            tasks.Push(task);
            return trigger.resultingWindow;
        }

        public static void Pop(Action action = null)
        {
            action?.Invoke();
            tasks.Pop().Wait();
            RevalidateApp();
        }

        private static AutomationEventHandler GetOpenWindowHandler(int processId)
        {
            return (sender, e) =>
            {
                AutomationElement element = sender as AutomationElement;
                string processName = Process.GetProcessById(element.Current.ProcessId).ProcessName;
                if (element.Current.ProcessId != processId &&
                    processName != Constants.pptProcess) { return; }
                IntPtr handle = new IntPtr(element.Current.NativeWindowHandle);
                string windowName = WindowUtil.GetWindowTitle(handle);
                if (windowName == "")
                {
                    Logger.Log("Titleless window in windowwatcher: " +
                        $"pID: {element.Current.ProcessId}\n" +
                        $"process name: {processName}\n" +
                        $"HWND: {handle}");
                    return;
                }
                lastOpenWindowName = windowName;

                WindowOpenTrigger resultTrigger = GetWindowTrigger(windowName);
                if (resultTrigger == null)
                {
                    WindowUtil.CloseWindow(handle);
                    //MessageBox.Show(windowName);
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
