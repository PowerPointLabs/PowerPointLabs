using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    partial class ResizeLabMain
    {
        // To be used for error handling
        internal const int Match_MinNoOfShapesRequired = 1;
        internal const string Match_FeatureName = "Match";
        internal const string Match_ShapeSupport = "object";
        internal static readonly string[] Match_ErrorParameters =
        {
            Match_FeatureName,
            Match_MinNoOfShapesRequired.ToString(),
            Match_ShapeSupport
        };

        /// <summary>
        /// Resize the height of selected shapes to match their width.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void MatchWidth(PowerPoint.ShapeRange selectedShapes)
        {
            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                shape.Height = shape.Width;
            }
        }

        /// <summary>
        /// Resize the width of selected shapes to match their height.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void MatchHeight(PowerPoint.ShapeRange selectedShapes)
        {
            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                shape.Width = shape.Height;
            }
        }
    }
}
