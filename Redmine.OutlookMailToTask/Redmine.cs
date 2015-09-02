﻿using Redmine.OutlookMailToTask.Properties;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Redmine();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Redmine.OutlookMailToTask
{
    [ComVisible(true)]
    public class Redmine : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private bool _isRibbonButtonEnabled = false;
        private string _userName = string.Empty;

        public Redmine()
        {
            UpdateRedmineUser();
        }

        public void OnMyButtonClick(Office.IRibbonControl control)
        {
            if (control.Context is Outlook.Selection)
            {
                Outlook.Selection sel = control.Context as Outlook.Selection;
                Outlook.MailItem mail = (Outlook.MailItem)sel[1];
                //MessageBox.Show(mail.Subject.ToString());

                // if no settings is saved prompt user to fill it
                if (string.IsNullOrEmpty(_userName))
                {
                    DoShowSettings();
                }

                // if still no password, skip it...
                if (string.IsNullOrEmpty(_userName))
                {
                    return;
                }

                // Ask for project
                SelectProjectWindow projectWindow = new SelectProjectWindow();

                // use WindowInteropHelper to set the Owner of our WPF window to the Outlook application window
                System.Windows.Interop.WindowInteropHelper hwndHelper = new System.Windows.Interop.WindowInteropHelper(projectWindow);

                hwndHelper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle; // new IntPtr(Globals.ThisAddIn.Application.ActiveWindow().WindowHandle32);

                // show our window
                bool? result = projectWindow.ShowDialog();
                if (result.HasValue && result.Value == false)
                {
                    // Cancel task
                    return;
                }

                //if (string.IsNullOrEmpty(_userName) == false)
                //{
                Net.Api.RedmineManager manager = new Net.Api.RedmineManager(Settings.Default.RedmineServer, Settings.Default.RedmineApi, Net.Api.MimeFormat.xml);

                Net.Api.Types.Issue issue = new Net.Api.Types.Issue();
                issue.Subject = mail.Subject;
                if (mail.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                    issue.Description = mail.HTMLBody;
                else
                    issue.Description = mail.Body;

                issue.Project = new Net.Api.Types.Project() { Id = 1 };

                //var users = manager.GetObjectList<Net.Api.Types.User>(new NameValueCollection { { "name", GetSenderSMTPAddress(mail) } });
                //if (users.Count == 1)
                //{
                //issue.Author = new Net.Api.Types.IdentifiableName() { Id = users.FirstOrDefault().Id };
                //issue.AssignedTo = new Net.Api.Types.IdentifiableName() { Id = users.FirstOrDefault().Id };
                //}

                List<Net.Api.Types.Upload> attachments = new List<Net.Api.Types.Upload>();

                if (mail.Attachments.Count > 0)
                {
                    foreach (Outlook.Attachment att in mail.Attachments)
                    {
                        //MessageBox.Show(att.FileName);

                        try
                        {
                            string tempFile = Path.GetTempFileName();
                            att.SaveAsFile(tempFile);
                            var upload = manager.UploadFile(File.ReadAllBytes(tempFile));
                            upload.FileName = att.FileName;
                            upload.Description = att.DisplayName;
                            upload.ContentType = System.Web.MimeMapping.GetMimeMapping(att.FileName);

                            attachments.Add(upload);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show(string.Format("Cannot upload attachment {0}. Error: {1}", att.FileName, e.Message));
                        }
                    }
                }

                try
                {
                    issue.Uploads = attachments;
                    var createdIssue = manager.CreateObject(issue);
                }
                catch
                {
                    MessageBox.Show("Creation of the task failed.");
                    return;
                }

                if (Settings.Default.OpenTaskWhenCreated)
                {
                    System.Diagnostics.Process.Start(string.Format("{0}/issues/{1}", Settings.Default.RedmineServer, createdIssue.Id));
                }
                else
                {
                    MessageBox.Show(string.Format("Task has been created in Redmine with ID #{0}.", createdIssue.Id));
                }
            }
        }

        private string GetSenderSMTPAddress(Outlook.MailItem mail)
        {
            string PR_SMTP_ADDRESS =
                @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            if (mail == null)
            {
                throw new ArgumentNullException();
            }
            if (mail.SenderEmailType == "EX")
            {
                Outlook.AddressEntry sender =
                    mail.Sender;
                if (sender != null)
                {
                    //Now we have an AddressEntry representing the Sender
                    if (sender.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.
                        olExchangeUserAddressEntry
                        || sender.AddressEntryUserType ==
                        Outlook.OlAddressEntryUserType.
                        olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        Outlook.ExchangeUser exchUser =
                            sender.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.PrimarySmtpAddress;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return sender.PropertyAccessor.GetProperty(
                            PR_SMTP_ADDRESS) as string;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return mail.SenderEmailAddress;
            }
        }


        public string labelUserNameValue(Office.IRibbonControl control)
        {
            if (!string.IsNullOrEmpty(_userName))
            {
                return string.Format("Logged in as {0}.", _userName);
            }

            return string.Empty;
        }

        public bool labelUserNameEnabled(Office.IRibbonControl control)
        {
            return !string.IsNullOrEmpty(_userName);
        }

        public bool RibbonRedmineButtonEnabled(Office.IRibbonControl control)
        {
            return _isRibbonButtonEnabled;
        }

        public void SetRibbonButtonStatus(bool enabled)
        {
            _isRibbonButtonEnabled = enabled;

            InvalidateRibbon();
        }

        private void InvalidateRibbon()
        {
            if (ribbon != null)
            {
                ribbon.Invalidate();
            }
        }

        public void OnShow(object contextObject)
        {
            //UpdateRedmineUser();
        }

        public Office.BackstageGroupStyle GetWorkStatusStyle(Office.IRibbonControl control)
        {
            return string.IsNullOrEmpty(_userName) ?
                Office.BackstageGroupStyle.BackstageGroupStyleWarning :
                Office.BackstageGroupStyle.BackstageGroupStyleNormal;
        }

        public void LogUserOut(Office.IRibbonControl control)
        {
            Settings.Default.RedmineApi = string.Empty;
            Settings.Default.RedmineServer = string.Empty;
            Settings.Default.Save();
            _userName = string.Empty;

            InvalidateRibbon();
        }

        private void DoShowSettings()
        {
            OptionsWindow window = new OptionsWindow();

            // use WindowInteropHelper to set the Owner of our WPF window to the Visio application window
            System.Windows.Interop.WindowInteropHelper hwndHelper = new System.Windows.Interop.WindowInteropHelper(window);

            hwndHelper.Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle; // new IntPtr(Globals.ThisAddIn.Application.ActiveWindow().WindowHandle32);

            // show our window
            window.ShowDialog();

            // if OK was selected then do work
            if (window.DialogResult.HasValue && window.DialogResult.Value)
            {
                // do any work based on the success of the DialogResult property
                UpdateRedmineUser();
            }
        }

        public void ShowSettings(Office.IRibbonControl control)
        {
            DoShowSettings();
        }

        private void UpdateRedmineUser()
        {
            try
            {
                if (string.IsNullOrEmpty(Settings.Default.RedmineServer) == false && string.IsNullOrEmpty(Settings.Default.RedmineApi) == false)
                {
                    Net.Api.RedmineManager manager = new Net.Api.RedmineManager(Settings.Default.RedmineServer, Settings.Default.RedmineApi, Net.Api.MimeFormat.xml);
                    var user = manager.GetCurrentUser();

                    _userName = string.Format("{0} {1}", user.FirstName, user.LastName);

                    InvalidateRibbon();
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Error: {0}", e.Message);

                _userName = string.Empty;

                InvalidateRibbon();
            }
        }

        public Image GetIcon(Office.IRibbonControl control)
        {
            return Resources.RedmineLogo;
        }

        public string GetConvertToRedmineLabel(Office.IRibbonControl control)
        {
            return "Convert to Redmine task";
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Redmine.OutlookMailToTask.Redmine.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}