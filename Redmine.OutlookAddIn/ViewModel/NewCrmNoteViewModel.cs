using Mach.Wpf.Mvvm;
using Redmine.Net.Api.Extensions;
using Redmine.OutlookAddIn.Properties;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Redmine.OutlookAddIn.ViewModel
{
    public class NewCrmNoteViewModel : NotifyPropertyBase
    {
        private ObservableCollection<ContactViewModel> _contacts;
        public ObservableCollection<ContactViewModel> Contacts
        {
            get { return _contacts; }
            set
            {
                _contacts = value;
                OnPropertyChanged();
            }
        }

        private ContactViewModel _selectedContact;
        public ContactViewModel SelectedContact
        {
            get { return _selectedContact; }
            set
            {
                _selectedContact = value;
                OnPropertyChanged();
            }
        }

        public NewCrmNoteViewModel()
        {
            ReloadContactsList();
        }

        public void ReloadContactsList()
        {
            Task.Factory.StartNew(() => LoadContactsFromRedmine()).ContinueWith((t) =>
            {
                if (t.Result != null)
                {
                    Contacts = t.Result;
                }
            });
        }

        private ObservableCollection<ContactViewModel> LoadContactsFromRedmine()
        {
            var contactsList = new ObservableCollection<ContactViewModel>();

            IList<Net.Api.Types.Contact> contacts = null;
            try
            {
                // connect to redmine
                Net.Api.RedmineManager manager = new Net.Api.RedmineManager(Settings.Default.RedmineServer, Settings.Default.RedmineApi, Net.Api.MimeFormat.xml);

                contacts = manager.GetAllObjectList<Net.Api.Types.Contact>(new System.Collections.Specialized.NameValueCollection { });
            }
            catch { }

            if (contacts == null)
            {
                return null;
            }

            foreach (var contact in contacts)
            {
                ContactViewModel contactViewModel = new ContactViewModel();
                contactViewModel.Id = contact.Id;
                contactViewModel.FirstName = contact.FirstName;
                contactViewModel.LastName = contact.LastName;
                if (contact.ContactType != null)
                {
                    ContactTypeViewModel contactType = new ContactTypeViewModel();
                    contactType.Id = contact.ContactType.Id;
                    contactType.Name = contact.ContactType.Name;

                    contactViewModel.ContactType = contactType;
                }

                //if (contact.CustomFields != null)
                //{
                //    foreach (var customField in contact.CustomFields)
                //    {
                //        CustomFieldViewModel customFieldViewModel = new CustomFieldViewModel();
                //        customFieldViewModel.Id = customField.Id;
                //        customFieldViewModel.Name = customField.Name;

                //        contactViewModel.CustomFields.Add(customFieldViewModel);
                //    }
                //}

                contactsList.Add(contactViewModel);
            }

            return contactsList;
        }
    }
}
