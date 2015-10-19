﻿using System;
using System.Collections.Generic;
using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AnimateInSlideTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AnimateInSlide.pptx";
        }

        [TestMethod]
        public void FT_AnimateInSlideTest()
        {
            PpOperations.SelectSlide(4);
            PpOperations.SelectShapes(new List<String> { "Rectangle 2", "Rectangle 5", "Rectangle 6" });

            PplFeatures.AnimateInSlide();

            var actualSlide = PpOperations.SelectSlide(4);
            var expSlide = PpOperations.SelectSlide(5);

            // remove text "Expected"
            PpOperations.SelectShape("text 3")[1].Delete();
            SlideUtil.IsSameAnimations(expSlide, actualSlide);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }
    }
}
