using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace Redmine.OutlookMailToTask
{
    public partial class ThisAddIn
    {
        private Redmine _ribbon; 
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Outlook.Application oOutlook = Globals.ThisAddIn.Application;
            //oOutlook.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);
            if (oOutlook.ActiveExplorer() != null)
            {
                oOutlook.ActiveExplorer().SelectionChange += ThisAddIn_SelectionChange;
            }
        }

        private void ThisAddIn_SelectionChange()
        {
            bool enabled = Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count == 1;

            _ribbon.SetRibbonButtonStatus(enabled);

            //{
                System.Diagnostics.Debug.WriteLine(Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count);
            //}

        }

        private void OOutlook_ItemLoad(object Item)
        {
            System.Diagnostics.Debug.WriteLine(Item);
        }

        //void Application_OptionsPagesAdd(Outlook.PropertyPages Pages)
        //{
        //    Pages.Add(new ConfigurationOptions(), "");
        //}


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Redmine();
            return _ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
