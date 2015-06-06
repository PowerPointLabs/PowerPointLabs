
using System;
using Microsoft.QualityTools.Testing.Fakes;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [TestClass]
    public class BaseUnitTest
    {
        private IDisposable _shimsContext;

        [TestInitialize]
        public void Setup()
        {
            _shimsContext = ShimsContext.Create();
        }

        [TestCleanup]
        public void TearDown()
        {
            _shimsContext.Dispose();
        }
    }
}
