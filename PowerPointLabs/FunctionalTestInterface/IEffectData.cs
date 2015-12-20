using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace FunctionalTestInterface
{
    public interface IEffectData
    {
        MsoAnimEffect EffectType { get; }
        MsoShapeType ShapeType { get; }
        float ShapeRotation { get; }
        float ShapeWidth { get; }
        float ShapeHeight { get; }
        float ShapeLeft { get; }
        float ShapeTop { get; }
        float TimingTriggerDelayTime { get; }
    }
}
