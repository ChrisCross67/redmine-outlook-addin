using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Redmine.OutlookMailToTask.ViewModel
{
    public class ProjectViewModel
    {
        public int ParentId
        {
            get; set;
        }
        public int Id
        {
            get; set;
        }

        public int Level
        {
            get; set;
        }

        public string Name
        {
            get; set;
        }

        public string Path
        {
            get; set;
        }

        public TrackerViewModel Tracker { get; set; } 

        public List<TrackerViewModel> Trackers
        {
            get; set;
        }
        public List<CustomFieldViewModel> CustomFields
        {
            get; set;
        }

        public ObservableCollection<ProjectViewModel> Children
        {
            get; set;
        }

        public ProjectViewModel()
        {
            CustomFields = new List<CustomFieldViewModel>();
            Trackers = new List<TrackerViewModel>();
            Children = new ObservableCollection<ProjectViewModel>();
        }
    }
}
