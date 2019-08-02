namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class CustomEffect: AbstractEffect
    {
        public string shapeId;
        public AnimationType type;
        public CustomEffect(string shapeName, string shapeId, AnimationType type): base(shapeName)
        {
            this.shapeId = shapeId;
            this.type = type;
        }
    }
}
