using Mach.Wpf.Mvvm;
using Redmine.Net.Api.Extensions;
using Redmine.OutlookAddIn.Properties;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace Redmine.OutlookAddIn.ViewModel
{
    public class NewCrmNoteViewModel : NotifyPropertyBase
    {
        private Dispatcher _uiDispatcher;

        private ListCollectionView _filteredContacts;
        public ListCollectionView FilteredContacts
        {
            get { return _filteredContacts; }
            set
            {
                _filteredContacts = value;
                OnPropertyChanged();
            }
        }

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

        private string _filter;
        public string Filter
        {
            get { return _filter; }
            set
            {
                _filter = value;
                OnPropertyChanged();
            }
        }

        private DelegateCommand _addNoteCommand;
        public ICommand AddNoteCommand
        {
            get { return _addNoteCommand; }
        }

        private DelegateCommand _filterContactsCommand;
        public ICommand FilterContactsCommand
        {
            get { return _filterContactsCommand; }
        }

        private bool _contactsLoaded;
        public bool ContactsLoaded
        {
            get { return _contactsLoaded; }
            private set
            {
                _contactsLoaded = value;
                OnPropertyChanged();
            }
        }

        public NewCrmNoteViewModel()
        {
            _addNoteCommand = new DelegateCommand(AddNote, CanAddNote);
            _filterContactsCommand = new DelegateCommand(FilterContacts, CanFilterContacts);

            ReloadContactsList();
        }

        private bool CanFilterContacts(object parameter)
        {
            return _contactsLoaded;
        }

        private void FilterContacts()
        {
            _filteredContacts.Filter = ContactFilter;

            _filteredContacts.Refresh();
        }

        private bool ContactFilter(object obj)
        {
            ContactViewModel contact = obj as ContactViewModel;

            if (contact == null)
                return false;

            if (string.IsNullOrEmpty(_filter))
                return true;

            return CultureInfo.CurrentCulture.CompareInfo.IndexOf(contact.Name, _filter, CompareOptions.IgnoreNonSpace | CompareOptions.IgnoreCase) > -1;
        }

        private bool CanAddNote(object parameter)
        {
            return _contactsLoaded;
        }

        private void AddNote()
        {
            throw new NotImplementedException();
        }

        public void ReloadContactsList()
        {
            _uiDispatcher = Dispatcher.CurrentDispatcher;
            _contacts = new ObservableCollection<ContactViewModel>();
            _filteredContacts = new ListCollectionView(_contacts);
            _filteredContacts.Filter = ContactFilter;

            Task.Factory.StartNew(() => LoadContactsFromRedmine()).ContinueWith((t) =>
            {
                if (t.Result != null)
                {
                    _uiDispatcher.Invoke(() =>
                    {
                        t.Result.ForEach(c => _contacts.Add(c));
                    });

                    ContactsLoaded = true;
                }
            });
        }

        private List<ContactViewModel> LoadContactsFromRedmine()
        {
            var contactsList = new List<ContactViewModel>();

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
