using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.HighlightLab
{
    [TestClass]
    public class BaseHighlightLabTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "HighlightPoints.pptx";
        }
    }
}
