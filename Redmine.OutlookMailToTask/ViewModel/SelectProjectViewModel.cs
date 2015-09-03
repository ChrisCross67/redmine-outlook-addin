using Mach.Wpf.Mvvm;
using System.Collections.ObjectModel;
using System;
using System.Collections.Generic;
using System.Linq;
using Redmine.OutlookMailToTask.Properties;
using System.Threading.Tasks;

namespace Redmine.OutlookMailToTask.ViewModel
{
    public class SelectProjectViewModel : NotifyPropertyBase
    {
        private ObservableCollection<ProjectViewModel> _projects;
        public ObservableCollection<ProjectViewModel> Projects
        {
            get { return _projects; }
            set
            {
                _projects = value;
                OnPropertyChanged();
            }
        }

        private ProjectViewModel _selectedProject;
        public ProjectViewModel SelectedProject
        {
            get { return _selectedProject; }
            set
            {
                _selectedProject = value;
                OnPropertyChanged();
            }
        }

        private ObservableCollection<ProjectViewModel> _flatProjects;
        public ObservableCollection<ProjectViewModel> FlatProjects
        {
            get { return _flatProjects; }
            set
            {
                _flatProjects = value;
                OnPropertyChanged();
            }
        }

        public SelectProjectViewModel() { }

        public void ReloadProjectsList()
        {
            Task.Factory.StartNew(() => LoadProjectsFromRedmine()).ContinueWith((t) =>
            {
                if (t.Result != null)
                {
                    Projects = SortProjects(t.Result);
                    FlatProjects = FlattenProjects(_projects);

                    SetSelectedProject(Settings.Default.LastUsedProjectId);
                }
            });
        }

        public void SetSelectedProject(int projectId)
        {
            SelectedProject = _flatProjects.Where(p => p.Id == projectId).FirstOrDefault();
        }

        private List<ProjectViewModel> LoadProjectsFromRedmine()
        {
            var projectsList = new List<ProjectViewModel>();

            try
            {
                // connect to redmine
                Net.Api.RedmineManager manager = new Net.Api.RedmineManager(Settings.Default.RedmineServer, Settings.Default.RedmineApi, Net.Api.MimeFormat.xml);

                var projects = manager.GetObjectList<Net.Api.Types.Project>(new System.Collections.Specialized.NameValueCollection { { "limit", "100" }, { "include", "trackers,issue_categories" } });

                if (projects != null)
                {
                    foreach (var project in projects)
                    {
                        ProjectViewModel projectViewModel = new ProjectViewModel();
                        projectViewModel.Id = project.Id;
                        projectViewModel.Name = project.Name;
                        if (project.Parent != null)
                        {
                            projectViewModel.ParentId = project.Parent.Id;
                        }

                        if (project.CustomFields != null)
                        {
                            foreach (var customField in project.CustomFields)
                            {
                                CustomFieldViewModel customFieldViewModel = new CustomFieldViewModel();
                                customFieldViewModel.Id = customField.Id;
                                customFieldViewModel.Name = customField.Name;

                                projectViewModel.CustomFields.Add(customFieldViewModel);
                            }
                        }

                        if (project.Trackers != null)
                        {
                            foreach (var tracker in project.Trackers)
                            {
                                TrackerViewModel trackerViewModel = new TrackerViewModel();
                                trackerViewModel.Id = tracker.Id;
                                trackerViewModel.Name = tracker.Name;

                                projectViewModel.Trackers.Add(trackerViewModel);

                                // set first one as default
                                if (projectViewModel.Tracker == null)
                                {
                                    projectViewModel.Tracker = trackerViewModel;
                                }
                            }
                        }

                        if (project.IssueCategories != null)
                        {
                            foreach (var issueCategory in project.IssueCategories)
                            {
                                IssueCategoryViewModel issueCategoryViewModel = new IssueCategoryViewModel();
                                issueCategoryViewModel.Id = issueCategory.Id;
                                issueCategoryViewModel.Name = issueCategory.Name;

                                projectViewModel.IssueCategories.Add(issueCategoryViewModel);
                            }
                        }

                        projectsList.Add(projectViewModel);
                    }

                    return projectsList;
                }

            }
            catch { }

            return null;
        }

        private ObservableCollection<ProjectViewModel> FlattenProjects(IList<ProjectViewModel> projects)
        {
            ObservableCollection<ProjectViewModel> flatProjects = new ObservableCollection<ProjectViewModel>();

            Action<ProjectViewModel> getChildren = null;

            getChildren = parent =>
            {
                flatProjects.Add(parent);
                parent.Children.ForEach(p =>
                {
                    getChildren(p);
                });
            };

            projects.ForEach(getChildren);

            return flatProjects;
        }

        private ObservableCollection<ProjectViewModel> SortProjects(List<ProjectViewModel> projects)
        {
            ObservableCollection<ProjectViewModel> sortedProjects = new ObservableCollection<ProjectViewModel>();

            Action<ProjectViewModel> setChildren = null;

            setChildren = parent =>
            {
                var childProjects = projects.Where(childItem => childItem.ParentId == parent.Id).ToList();
                childProjects.ForEach(p =>
                {
                    p.Level = parent.Level + 1;
                    p.Path = parent.Path + " » " + p.Name;
                });

                parent.Children = new ObservableCollection<ProjectViewModel>(childProjects);

                // Recursively call the SetChildren method for each child.
                parent.Children.ForEach(setChildren);
            };

            // Initialize the hierarchical list to root level items
            sortedProjects = new ObservableCollection<ProjectViewModel>(projects.Where(rootItem => rootItem.ParentId == 0).ToList());

            // Call the SetChildren method to set the children on each root level item.
            sortedProjects.ForEach(p => p.Path = p.Name);
            sortedProjects.ForEach(setChildren);

            return sortedProjects;
        }
    }
}
