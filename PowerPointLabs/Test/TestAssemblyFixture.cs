using System;
using System.Windows;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.Utils;

namespace Test
{
    [TestClass]
    public class TestAssemblyFixture
    {
        [AssemblyInitialize]
        public static void AssemblySetup(TestContext context)
        {
            TempPath.CreateTempTestFolder();
        }

        [AssemblyCleanup]
        public static void AssemblyTeardown()
        {
            TempPath.DeleteTempTestFolder();
        }
    }
}
