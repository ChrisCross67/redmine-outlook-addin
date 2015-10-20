using Mach.Wpf.Mvvm;

namespace Redmine.OutlookAddIn.ViewModel
{
    public class TrackerViewModel : NotifyPropertyBase
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
