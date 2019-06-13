using TestInterface.Windows;

namespace TestInterface
{
    public interface ISyncLabController
    {
        IMarshalWindow Dialog { get; }

        void OpenPane();

        void Copy();

        void Sync(int index);

        void DialogSelectItem(int categoryIndex, int itemIndex);

        void DialogClickOk();

        bool GetCopyButtonEnabledStatus();
    }
}
