using System;
using System.Windows.Input;

namespace editconfig
{
    class Command : ICommand
    {
        private readonly Action execute;
        private readonly Func<bool> canExecute;

        public Command(Action execute, Func<bool> canExecute)
        {
            this.execute = execute;
            this.canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return canExecute();
        }

        public void OnCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Execute(object parameter)
        {
            execute();
        }
    }
}
