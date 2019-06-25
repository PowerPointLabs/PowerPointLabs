namespace TestInterface
{
    public interface IELearningLabController
    {
        void OpenPane();
        void CreateTemplateExplanations(params IExplanationItem[] items);
        void AddSelfExplanationItem();
        void Sync();
        void Reorder();
        void Delete();
    }
}
