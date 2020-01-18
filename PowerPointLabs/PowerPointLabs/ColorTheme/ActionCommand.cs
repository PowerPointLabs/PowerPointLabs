using System;
using System.Windows.Input;

namespace PowerPointLabs.ColorThemes
{
    public class ActionCommand : ICommand
    {
        private Action _action;
        public ActionCommand(Action action)
        {
            _action = action;
        }

        event EventHandler ICommand.CanExecuteChanged
        {
            add
            {
                throw new NotImplementedException();
            }

            remove
            {
                throw new NotImplementedException();
            }
        }

        bool ICommand.CanExecute(object parameter)
        {
            return true;
        }

        void ICommand.Execute(object parameter)
        {
            _action();
        }
    }
}
