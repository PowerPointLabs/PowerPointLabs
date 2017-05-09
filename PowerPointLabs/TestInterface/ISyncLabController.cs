using Microsoft.Office.Interop.PowerPoint;

namespace TestInterface
{
    public interface ISyncLabController
    {
        void OpenPane();

        void Copy();

        void Sync(int index);
    }
}
