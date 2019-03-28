using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ELearningLabTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "ELearningLab\\ELearningLabTest.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CreateSelfExplanationTest()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShapes(new List<string> { "Rectangle 2", "Rectangle 5", "Rectangle 6" });
        }
        }
    }
