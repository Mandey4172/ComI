using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;

namespace ComI
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("1D9ECCF3-5D2F-4112-9B25-264596873DC9")]
    public class SettingsPage : UIElementDialogPage
    {
        /// <summary>
        /// </summary>
        protected override System.Windows.UIElement Child
        {
            get { return new SettingsControl(this); }
        }
    }
}
