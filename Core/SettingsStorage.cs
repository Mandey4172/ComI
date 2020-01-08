using ComI.Resources;
using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComI.Core
{
    class SettingsStorage
    {
        //private const string settingsCollectionName = "ComI";

        public static void SaveProperty(string propertyName, string propertyValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            SettingsManager settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
            WritableSettingsStore userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            if (!userSettingsStore.CollectionExists(SettingsProperties.settingsCollectionName))
                userSettingsStore.CreateCollection(SettingsProperties.settingsCollectionName);

            userSettingsStore.SetString(SettingsProperties.settingsCollectionName, propertyName, propertyValue);                
        }

        public static string ReadProperty(string propertyName)
        {
            string output = "";
            ThreadHelper.ThrowIfNotOnUIThread();
            SettingsManager settingsManager = new ShellSettingsManager(ServiceProvider.GlobalProvider);
            WritableSettingsStore userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            if (userSettingsStore.CollectionExists(SettingsProperties.settingsCollectionName) && userSettingsStore.PropertyExists(SettingsProperties.settingsCollectionName, propertyName))
                output = userSettingsStore.GetString(SettingsProperties.settingsCollectionName, propertyName);

            return output;
        }
    }
}
