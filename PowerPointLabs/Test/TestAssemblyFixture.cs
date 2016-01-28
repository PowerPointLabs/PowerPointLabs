using System;
using System.IO;
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
            if (!TempPath.IsExistingTempFolder())
            {
                TempPath.CreateTempTestFolder();
            }
        }

        [AssemblyCleanup]
        public static void AssemblyTeardown()
        {
            TempPath.DeleteTempTestFolder();
        }
    }
}
