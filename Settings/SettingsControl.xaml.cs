using ComI.Core;
using ComI.Resources;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
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

namespace ComI
{
    /// <summary>
    /// Interaction logic for SettingsControl.xaml
    /// </summary>
    public partial class SettingsControl : UserControl
    {
        //private const string propertyName = "Test";

        private SettingsPage settingsPage;

        public SettingsControl(SettingsPage page)
        {
            this.settingsPage = page;
            InitializeComponent();

            textBox.Text = SettingsStorage.ReadProperty(SettingsProperties.documentMacroData);
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string message = "";
            if (textBox != null)
                message = textBox.Text;

            SettingsStorage.SaveProperty(SettingsProperties.documentMacroData, message);
        }
    }
}
