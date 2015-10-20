using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Redmine.OutlookMailToTask
{
    [ComVisible(true)]
    public partial class ConfigurationOptions : UserControl, Outlook.PropertyPage
    {
        const int captionDispID = -518;
        bool isDirty = false;

        public ConfigurationOptions()
        {
            InitializeComponent();
        }

        void Outlook.PropertyPage.Apply()
        {

        }
        bool Outlook.PropertyPage.Dirty
        {
            get
            {
                return isDirty;
            }
        }
        void Outlook.PropertyPage.GetPageInfo(ref string helpFile, ref int helpContext)
        {

        }

        [DispId(captionDispID)]
        public string PageCaption
        {
            get
            {
                return "Test page";
            }
        }
    }
}
