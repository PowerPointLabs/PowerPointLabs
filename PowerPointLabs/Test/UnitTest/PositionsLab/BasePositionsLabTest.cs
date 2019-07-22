using System;
using System.Collections.Generic;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PositionsLab;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

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

        // TODO: Reuse from PositionsPaneWPF.xaml.cs
        protected void SyncShapes(PowerPoint.ShapeRange selected, PowerPoint.ShapeRange simulatedShapes)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                Shape selectedShape = selected[i];
                Shape simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(simulatedShape.GetCenterPoint().X - selectedShape.GetCenterPoint().X);
                selectedShape.IncrementTop(simulatedShape.GetCenterPoint().Y - selectedShape.GetCenterPoint().Y);
                selectedShape.Rotation = simulatedShape.Rotation;
            }
        }

        protected void SyncShapes(PowerPoint.ShapeRange selected, PowerPoint.ShapeRange simulatedShapes, float[,] originalPositions)
        {
            for (int i = 1; i <= selected.Count; i++)
            {
                Shape selectedShape = selected[i];
                Shape simulatedShape = simulatedShapes[i];

                selectedShape.IncrementLeft(simulatedShape.GetCenterPoint().X - originalPositions[i - 1, Left]);
                selectedShape.IncrementTop(simulatedShape.GetCenterPoint().Y - originalPositions[i - 1, Top]);
                SwapZOrder(selectedShape, simulatedShape);
            }
        }

        protected PowerPoint.ShapeRange DuplicateShapes(PowerPoint.ShapeRange range)
        {
            int totalShapes = PpOperations.GetCurrentSlide().Shapes.Count;
            int[] duplicatedShapeIndices = new int[range.Count];

            for (int i = 1; i <= range.Count; i++)
            {
                Shape shape = range[i];
                Shape duplicated = shape.Duplicate()[1];
                duplicated.Name = shape.Id + "";
                duplicated.Left = shape.Left;
                duplicated.Top = shape.Top;
                duplicatedShapeIndices[i - 1] = totalShapes + i;
            }

            return PpOperations.GetCurrentSlide().Shapes.Range(duplicatedShapeIndices);
        }

        protected float[,] SaveOriginalPositions(List<PPShape> shapes)
        {
            float[,] initialPositions = new float[shapes.Count, 2];
            for (int i = 0; i < shapes.Count; i++)
            {
                PPShape s = shapes[i];
                System.Drawing.PointF pt = s.VisualCenter;
                initialPositions[i, Left] = pt.X;
                initialPositions[i, Top] = pt.Y;
            }

            return initialPositions;
        }

        protected List<PPShape> ConvertShapeRangeToPPShapeList(PowerPoint.ShapeRange range, int index)
        {
            List<PPShape> shapes = new List<PPShape>();

            for (int i = index; i <= range.Count; i++)
            {
                shapes.Add(new PPShape(range[i]));
            }

            return shapes;
        }

        protected List<PowerPoint.Shape> ConvertShapeRangeToShapeList(PowerPoint.ShapeRange range, int index)
        {
            List<Shape> shapes = new List<PowerPoint.Shape>();

            for (int i = index; i <= range.Count; i++)
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

                if (PositionsLabSettings.AlignReference == PositionsLabSettings.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes);
                }
                else if (isConvertPPShape)
                {
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                if (PositionsLabSettings.AlignReference == PositionsLabSettings.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension);
                }
                else
                {
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                if (PositionsLabSettings.AlignReference == PositionsLabSettings.AlignReferenceObject.PowerpointDefaults)
                {
                    positionsAction.Invoke(selectedShapes, dimension1, dimension2);
                }
                else
                {
                    List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                    float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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

                // set the zOrder
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    SwapZOrder(simulatedShapes[i], selectedShapes[i]);
                }

                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                List<PPShape> simulatedPPShapes = ConvertShapeRangeToPPShapeList(simulatedShapes, 1);
                float[,] initialPositions = SaveOriginalPositions(simulatedPPShapes);

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
                List<Shape> shapes = ConvertShapeRangeToShapeList(selectedShapes, 1);

                positionsAction.Invoke(shapes);

                GC.Collect();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void UpdateZOrder(Shape original, int zOrder)
        {
            while (original.ZOrderPosition != zOrder)
            {
                if (original.ZOrderPosition < zOrder)
                {
                    original.ZOrder(MsoZOrderCmd.msoBringForward);
                }
                else if (original.ZOrderPosition > zOrder)
                {
                    original.ZOrder(MsoZOrderCmd.msoSendBackward);
                }
            }
        }

        private void SwapZOrder(Shape original, Shape target)
        {
            int originalZOrder = original.ZOrderPosition;
            int targetZOrder = target.ZOrderPosition;

            UpdateZOrder(original, targetZOrder);
            UpdateZOrder(target, originalZOrder);
        }
    }
}
