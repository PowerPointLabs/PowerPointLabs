namespace TestInterface
{
    public interface IELearningLabController
    {
        string DefaultVoiceLabel { get; }

        void OpenPane();
        void CreateTemplateExplanations(params ExplanationItemTemplate[] items);
        ExplanationItemTemplate[] GetExplanations();
        void AddSelfExplanationItem();
        void AddAbove(int index);
        void AddBelow(int index);
        void AddAtBottom();
        void Sync();
        void Reorder();
        void Delete();
    }
}
