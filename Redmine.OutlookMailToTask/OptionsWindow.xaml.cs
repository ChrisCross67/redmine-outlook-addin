//-----------------------------------------------------------------------
// 
//  Copyright (C) Microsoft Corporation.  All rights reserved.
// 
// THIS CODE AND INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY
// KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//-----------------------------------------------------------------------

using Redmine.OutlookMailToTask.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Redmine.OutlookMailToTask
{
    /// <summary>
    /// Interaction logic for OptionsWindow.xaml
    /// </summary>
    public partial class OptionsWindow : Window
    {
        public OptionsWindow()
        {
            InitializeComponent();

            Loaded += OptionsWindow_Loaded;
        }

        private void OptionsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (Settings.Default.RedmineApi != null)
            {
                apiKeyTextBox.Text = Settings.Default.RedmineApi;
            }

            if (Settings.Default.RedmineServer != null)
            {
                serverTextBox.Text = Settings.Default.RedmineServer;
            }

            openTaskInBrowser.IsChecked = Settings.Default.OpenTaskWhenCreated;
        }

        private void buttonClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // connect to redmine
                Net.Api.RedmineManager manager = new Net.Api.RedmineManager(serverTextBox.Text, apiKeyTextBox.Text, Net.Api.MimeFormat.xml);

                var user = manager.GetCurrentUser();
                if (user.Id > 0)
                {
                    Settings.Default.RedmineApi = apiKeyTextBox.Text;
                    Settings.Default.RedmineServer = serverTextBox.Text;
                    Settings.Default.OpenTaskWhenCreated = openTaskInBrowser.IsChecked.HasValue && openTaskInBrowser.IsChecked.Value == true;

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
