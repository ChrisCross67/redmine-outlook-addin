using Redmine.OutlookAddIn.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Redmine.OutlookAddIn
{
    /// <summary>
    /// Interaction logic for SelectProjectWindow.xaml
    /// </summary>
    public partial class SelectProjectWindow : Window
    {
        public SelectProjectWindow()
        {
            InitializeComponent();
        }

        private void buttonClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // connect to redmine
                Net.Api.RedmineManager manager = new Net.Api.RedmineManager(Settings.Default.RedmineServer, Settings.Default.RedmineApi, Net.Api.MimeFormat.xml);

                var user = manager.GetCurrentUser();
                if (user.Id > 0)
                {
                    Settings.Default.Save();

                    DialogResult = true;
                }
            }
            catch
            {
                MessageBox.Show("Cannot connect to the Redmine. Please check your configuration", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }


        }
    }
}
