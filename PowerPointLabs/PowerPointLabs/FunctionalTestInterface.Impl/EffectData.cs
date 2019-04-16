using System;

using Microsoft.Office.Interop.PowerPoint;

using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    public class EffectData : IEffectData
    {
        public int EffectType { get; private set; }
        public int ShapeType { get; private set; }
        public float ShapeRotation { get; private set; }
        public float ShapeWidth { get; private set; }
        public float ShapeHeight { get; private set; }
        public float ShapeLeft { get; private set; }
        public float ShapeTop { get; private set; }
        public float TimingTriggerDelayTime { get; private set; }

        private EffectData(Effect effect)
        {
            EffectType = effect.EffectType.GetHashCode();
            ShapeType = effect.Shape.Type.GetHashCode();
            ShapeRotation = effect.Shape.Rotation;
            ShapeWidth = effect.Shape.Width;
            ShapeHeight = effect.Shape.Height;
            ShapeLeft = effect.Shape.Left;
            ShapeTop = effect.Shape.Top;
            TimingTriggerDelayTime = effect.Timing.TriggerDelayTime;
        }

        public static IEffectData FromEffect(Effect effect)
        {
            return new EffectData(effect);
        }
    }
}
