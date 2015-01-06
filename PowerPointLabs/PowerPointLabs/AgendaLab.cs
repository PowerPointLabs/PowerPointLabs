using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    internal static class AgendaLab
    {
        # region Enum
        public enum AgendaType
        {
            Bullet,
            Visual
        };
        # endregion

        # region API
        public static void GenerateAgenda(AgendaType type)
        {
            switch (type)
            {
                case AgendaType.Bullet:
                    GenerateBulletAgenda();
                    break;
                case AgendaType.Visual:
                    GenerateVisualAgenda();
                    break;
            }
        }
        # endregion

        # region Helper Functions
        private static void GenerateBulletAgenda()
        {
            var newSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.Add(1, PpSlideLayout.ppLayoutText);

            var contentPlaceHolder = newSlide.Shapes.Placeholders[2];

            contentPlaceHolder.TextFrame.TextRange.Text = PowerPointCurrentPresentationInfo.Sections
                                                                                           .Skip(1)
                                                                                           .Aggregate((current, next) => current + "\n" + next);
        }

        private static void GenerateVisualAgenda()
        {

        }
        # endregion
    }
}
