using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Xml.Linq;
using System.Xml.XPath;

namespace editconfig
{
    internal class ViewModel : INotifyPropertyChanged
    {
        private const string ConfigPath = "pdf2rtf.exe.config";

        private bool _hasChanges;

        public ViewModel()
        {
            if (DesignerProperties.GetIsInDesignMode(new DependencyObject()))
            {
                return;
            }

            var document = XDocument.Load(ConfigPath);
            var settings = document.XPathSelectElement("//pdf2rtf.Properties.Settings");

            _inputFolder = settings.XPathSelectElement("//setting[@name='InputFolder']/value").Value;
            _outputFolder = settings.XPathSelectElement("//setting[@name='OutputFolder']/value").Value;
            _processedFolder = settings.XPathSelectElement("//setting[@name='ProcessedFolder']/value").Value;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private string _inputFolder;
        public string InputFolder
        {
            get => _inputFolder;
            set
            {
                if (_inputFolder != value)
                {
                    _inputFolder = value;
                    _hasChanges = true;
                    _saveCommand.OnCanExecuteChanged();
                    FirePropertyChanged("InputFolder");
                }
            }
        }

        private string _outputFolder;
        public string OutputFolder
        {
            get => _outputFolder;
            set
            {
                if (_outputFolder != value)
                {
                    _outputFolder = value;
                    _hasChanges = true;
                    _saveCommand.OnCanExecuteChanged();
                    FirePropertyChanged("OutputFolder");
                }
            }
        }

        private string _processedFolder;
        public string ProcessedFolder
        {
            get => _processedFolder;
            set
            {
                if (_processedFolder != value)
                {
                    _processedFolder = value;
                    _hasChanges = true;
                    _saveCommand.OnCanExecuteChanged();
                    FirePropertyChanged("ProcessedFolder");
                }
            }
        }

        private Command _saveCommand;
        public ICommand SaveCommand
        {
            get
            {
                if (_saveCommand == null)
                {
                    _saveCommand = new Command(SaveConfig, CanSaveConfig);
                }
                return _saveCommand;
            }
        }

        private bool CanSaveConfig()
        {
            return _hasChanges;
        }

        private void SaveConfig()
        {
            var document = XDocument.Load(ConfigPath);
            var settings = document.XPathSelectElement("//pdf2rtf.Properties.Settings");

            settings.XPathSelectElement("//setting[@name='InputFolder']/value").Value = InputFolder;
            settings.XPathSelectElement("//setting[@name='OutputFolder']/value").Value = OutputFolder;
            settings.XPathSelectElement("//setting[@name='ProcessedFolder']/value").Value = ProcessedFolder;

            document.Save(ConfigPath);

            _hasChanges = false;
            _saveCommand.OnCanExecuteChanged();
        }

        private void FirePropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
