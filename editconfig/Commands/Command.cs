using System;
using System.Windows.Input;

namespace editconfig
{
    internal class Command : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public Command(Action execute, Func<bool> canExecute)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return _canExecute();
        }

        public void OnCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Execute(object parameter)
        {
            _execute();
        }
    }
}
