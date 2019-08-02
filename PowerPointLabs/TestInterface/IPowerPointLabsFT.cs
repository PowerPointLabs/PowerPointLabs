using TestInterface.Windows;

namespace TestInterface
{
    public interface IPowerPointLabsFT
    {
        IPowerPointLabsFeatures GetFeatures();
        IPowerPointOperations GetOperations();
        IWindowStackManager GetWindowStackManager();
    }
}
