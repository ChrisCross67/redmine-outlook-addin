using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace Redmine.OutlookAddIn
{
    public partial class RedmineAddIn
    {
        Redmine _ribbon;

        // keeping references on event subscibed objects (http://stackoverflow.com/questions/14369102/event-bindings-not-always-happening-during-outlook-add-in-startup)
        Outlook.Explorers _explorers;
        Outlook.Explorer _activeExplorer;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _explorers = Application.Explorers;
            _explorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(NewExplorerEventHandler);

            _activeExplorer = Application.ActiveExplorer();
            _activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ExplorerSelectionChange);
        }

        public void NewExplorerEventHandler(Outlook.Explorer explorer)
        {
            if (explorer != null)
            {
                _activeExplorer = explorer;

                //set up event handler so our add-in can respond when a new item is selected in the Outlook explorer window
                _activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(ExplorerSelectionChange);
            }
        }

        private void ExplorerSelectionChange()
        {
            _activeExplorer = Application.ActiveExplorer();
            if (_activeExplorer == null) { return; }
            Outlook.Selection selection = _activeExplorer.Selection;

            if (_ribbon != null)
            {
                if (selection != null && selection.Count == 1 && selection[1] is Outlook.MailItem)
                {
                    // one Mail object is selected
                    Outlook.MailItem selectedMail = selection[1] as Outlook.MailItem; // (collection is 1-indexed)
                                                                                      // tell the ribbon what our currently selected email is by setting this custom property, and run code against it
                    _ribbon.CurrentEmail = selectedMail;
                }
                else
                {
                    _ribbon.CurrentEmail = null;
                }

                _ribbon.SetRibbonButtonStatus(_ribbon.CurrentEmail != null);
            }
        }

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
