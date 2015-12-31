namespace FunctionalTestInterface
{
    public interface IEffectData
    {
        int EffectType { get; }
        int ShapeType { get; }
        float ShapeRotation { get; }
        float ShapeWidth { get; }
        float ShapeHeight { get; }
        float ShapeLeft { get; }
        float ShapeTop { get; }
        float TimingTriggerDelayTime { get; }
    }
}
