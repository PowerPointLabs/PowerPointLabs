using Microsoft.Office.Core;

namespace PowerPointLabs.Utils
{
    class BoolUtil
    {
        public static bool ToBool(string state)
        {
            if (StringUtil.IsEmpty(state))
            {
                return false;
            }

            try
            {
                return bool.Parse(state);
            }
            catch
            {
                return false;
            }
        }

        public static bool ToBool(MsoTriState state)
        {
            return state == MsoTriState.msoTrue;
        }

        public static MsoTriState ToMsoTriState(bool boolean)
        {
            return boolean ? MsoTriState.msoTrue : MsoTriState.msoFalse;
        }
    }
}
