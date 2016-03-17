using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ResizeLab;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class BasePositionsLabTest : ResizeLab.BaseResizeLabTest
    {
        protected override string GetTestingSlideName()
        {
            return "PositionsLab.pptx";
        } 
    }
}
