namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public abstract class AbstractEffect
    {
        public string shapeName;
        protected AbstractEffect(string shapeName)
        {
            this.shapeName = shapeName;
        }
    }
}
