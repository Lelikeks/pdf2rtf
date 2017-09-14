using System.Windows;
using System.Windows.Controls;

namespace editconfig
{
    /// <summary>
    /// Interaction logic for FolderChooser.xaml
    /// </summary>
    public partial class FolderChooser : UserControl
    {
        public FolderChooser()
        {
            InitializeComponent();
        }

        public string SelectedFolder
        {
            get { return (string)GetValue(SelectedFolderProperty); }
            set { SetValue(SelectedFolderProperty, value); }
        }

        public static readonly DependencyProperty SelectedFolderProperty = DependencyProperty.Register("SelectedFolder", typeof(string), typeof(FolderChooser), new PropertyMetadata(null));

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                SelectedPath = SelectedFolder
            };

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SelectedFolder = dlg.SelectedPath;
            }
        }
    }
}
