﻿using Microsoft.Office.Interop.PowerPoint;

namespace TestInterface
{
    public interface ISyncLabController
    {
        void OpenPane();

        void Copy();

        void Sync(int index);

        void DialogSelectItem(int categoryIndex, int itemIndex);

        void DialogClickOk();
    }
}
