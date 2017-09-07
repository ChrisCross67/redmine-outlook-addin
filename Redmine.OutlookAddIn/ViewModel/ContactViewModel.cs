using Mach.Wpf.Mvvm;

namespace Redmine.OutlookAddIn.ViewModel
{
    public class ContactViewModel : NotifyPropertyBase
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public ContactTypeViewModel ContactType { get; set; }
        public string Name
        {
            get { return string.Format("{0} {1}", FirstName, LastName); }
        }
    }
}
