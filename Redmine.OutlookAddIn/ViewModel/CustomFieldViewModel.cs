using Mach.Wpf.Mvvm;

namespace Redmine.OutlookMailToTask.ViewModel
{
    public class CustomFieldViewModel : NotifyPropertyBase
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
