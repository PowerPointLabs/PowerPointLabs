using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ResizeLab;
using PowerPointLabs.Utils;
using System;
using PowerPointLabs.PositionsLab;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Diagnostics;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class BasePositionsLabTest : ResizeLab.BaseResizeLabTest
    {
        private const int Left = 0;
        private const int Top = 1;
        private readonly Dictionary<string, string> _originalShapeName = new Dictionary<string, string>();
        protected const string ErrorInvalidShapesSelected = "Invalid Shapes Selected.";

        protected override string GetTestingSlideName()
        {
            return "PositionsLab.pptx";
        }

        protected void SyncShapes(PowerPoint.ShapeRange selected, PowerPoint.ShapeRange simulatedShapes)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                var selectedShape = selected[i];
                var simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(Graphics.GetCenterPoint(simulatedShape).X - Graphics.GetCenterPoint(selectedShape).X);
                selectedShape.IncrementTop(Graphics.GetCenterPoint(simulatedShape).Y - Graphics.GetCenterPoint(selectedShape).Y);
                selectedShape.Rotation = simulatedShape.Rotation;
            }
        }

        protected void SyncShapes(PowerPoint.ShapeRange selected, PowerPoint.ShapeRange simulatedShapes, float[,] originalPositions)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                var selectedShape = selected[i];
                var simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(Graphics.GetCenterPoint(simulatedShape).X - originalPositions[i - 1, Left]);
                selectedShape.IncrementTop(Graphics.GetCenterPoint(simulatedShape).Y - originalPositions[i - 1, Top]);
            }
        }

        protected PowerPoint.ShapeRange DuplicateShapes(PowerPoint.ShapeRange range)
        {
            int totalShapes = PpOperations.GetCurrentSlide().Shapes.Count;
            int[] duplicatedShapeIndices = new int[range.Count];

            for (int i = 1; i <= range.Count; i++)
            {
                var shape = range[i];
                var duplicated = shape.Duplicate()[1];
                duplicated.Name = shape.Id + "";
                duplicated.Left = shape.Left;
                duplicated.Top = shape.Top;
                duplicatedShapeIndices[i - 1] = totalShapes + i;
            }

            return PpOperations.GetCurrentSlide().Shapes.Range(duplicatedShapeIndices);
        }

        protected float[,] SaveOriginalPositions(List<PPShape> shapes)
        {
            var initialPositions = new float[shapes.Count, 2];
            for (var i = 0; i < shapes.Count; i++)
            {
                var s = shapes[i];
                var pt = s.VisualCenter;
                initialPositions[i, Left] = pt.X;
                initialPositions[i, Top] = pt.Y;
            }

            return initialPositions;
        }

        protected List<PPShape> ConvertShapeRangeToPPShapeList(PowerPoint.ShapeRange range, int index)
        {
            var shapes = new List<PPShape>();

            for (var i = index; i <= range.Count; i++)
            {
                shapes.Add(new PPShape(range[i]));
            }

            return shapes;
        }

        protected List<PowerPoint.Shape> ConvertShapeRangeToShapeList(PowerPoint.ShapeRange range, int index)
        {
            var shapes = new List<PowerPoint.Shape>();

            for (var i = index; i <= range.Count; i++)
            {
                shapes.Add(range[i]);
            }

            return shapes;
        }

        protected void ExecutePositionsAction(Action<PowerPoint.ShapeRange> positionsAction, PowerPoint.ShapeRange selectedShapes,
            bool isConvertPPShape = true)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);

                if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes);
                }
                else if (isConvertPPShape)
                {
                    var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                    positionsAction.Invoke(simulatedShapes);

                    SyncShapes(selectedShapes, simulatedShapes, initialPositions);
                }
                else
                {
                    positionsAction.Invoke(simulatedShapes);

                    SyncShapes(selectedShapes, simulatedShapes);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        // Align right, bottom, vertical center, horizontal center
        protected void ExecutePositionsAction(Action<PowerPoint.ShapeRange, float> positionsAction, PowerPoint.ShapeRange selectedShapes, float dimension)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension);
                }
                else
                {
                    var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                    positionsAction.Invoke(simulatedShapes, dimension);

                    SyncShapes(selectedShapes, simulatedShapes, initialPositions);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        // Align center
        protected void ExecutePositionsAction(Action<PowerPoint.ShapeRange, float, float> positionsAction, PowerPoint.ShapeRange selectedShapes, float dimension1, float dimension2)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                if (PositionsLabMain.AlignReference == PositionsLabMain.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension1, dimension2);
                }
                else
                {
                    var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                    positionsAction.Invoke(simulatedShapes, dimension1, dimension2);

                    SyncShapes(selectedShapes, simulatedShapes, initialPositions);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        protected void ExecutePositionsAction(Action<List<PPShape>> positionsAction, PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes);

                SyncShapes(selectedShapes, simulatedShapes, initialPositions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        protected void ExecutePositionsAction(Action<List<PPShape>, bool> positionsAction, PowerPoint.ShapeRange selectedShapes, bool booleanVal)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, booleanVal);

                SyncShapes(selectedShapes, simulatedShapes, initialPositions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        protected void ExecutePositionsAction(Action<List<PPShape>, float> positionsAction, PowerPoint.ShapeRange selectedShapes, float dimension)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, dimension);

                SyncShapes(selectedShapes, simulatedShapes, initialPositions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        protected void ExecutePositionsAction(Action<List<PPShape>, float, float> positionsAction, PowerPoint.ShapeRange selectedShapes, float dimension1, float dimension2)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, dimension1, dimension2);

                SyncShapes(selectedShapes, simulatedShapes, initialPositions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        protected void ExecutePositionsAction(Action<List<PPShape>, int, int> positionsAction, PowerPoint.ShapeRange selectedShapes, int dimension1, int dimension2)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            PowerPoint.ShapeRange simulatedShapes = null;

            try
            {
                simulatedShapes = DuplicateShapes(selectedShapes);
                var simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                var initialPositions = SaveOriginalPositions(simulatedPPShapes);

                positionsAction.Invoke(simulatedPPShapes, dimension1, dimension2);

                SyncShapes(selectedShapes, simulatedShapes, initialPositions);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (simulatedShapes != null)
                {
                    simulatedShapes.Delete();
                    GC.Collect();
                }
            }
        }

        protected void ExecutePositionsAction(Action<IList<Shape>> positionsAction, PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count == 0)
            {
                throw new Exception(ErrorInvalidShapesSelected);
            }

            try
            {
                var shapes = ConvertShapeRangeToShapeList(selectedShapes, 1);

                positionsAction.Invoke(shapes);

                GC.Collect();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
