using System.Windows;

namespace Group4333
{
    public partial class MainWindow : Window
    {
        public MainWindow()
            => InitializeComponent();

        private void AuthorButton_Click(object sender, RoutedEventArgs e)
        {
            var infoWindow = new _4333_Minibaev();
            infoWindow.Owner = this;
            infoWindow.ShowDialog();
        }
    }
}