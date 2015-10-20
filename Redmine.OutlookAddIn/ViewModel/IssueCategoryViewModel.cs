using Mach.Wpf.Mvvm;

namespace Redmine.OutlookAddIn.ViewModel
{
    public class IssueCategoryViewModel : NotifyPropertyBase
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
