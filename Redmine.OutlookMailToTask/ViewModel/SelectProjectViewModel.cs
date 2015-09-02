using Mach.Wpf.Mvvm;
using System.Collections.ObjectModel;
using System.Windows;
using Redmine;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Redmine.OutlookMailToTask.ViewModel
{
    public class SelectProjectViewModel : NotifyPropertyBase
    {
        private ObservableCollection<ProjectViewModel> _projects;
        public ObservableCollection<ProjectViewModel> Projects
        {
            get
            {
                return _projects;
            }
            set
            {
                _projects = value;
                OnPropertyChanged();
            }
        }

        private ProjectViewModel _selectedProject;
        public ProjectViewModel SelectedProject
        {
            get
            {
                return _selectedProject;
            }
            set
            {
                _selectedProject = value;
                OnPropertyChanged();
            }
        }

        private ObservableCollection<ProjectViewModel> _flatProjects;
        public ObservableCollection<ProjectViewModel> FlatProjects
        {
            get
            {
                return _flatProjects;
            }
            set
            {
                _flatProjects = value;
                OnPropertyChanged();
            }
        }

        public SelectProjectViewModel()
        {
            var projectsList = new List<ProjectViewModel>();

            try
            {
                // connect to redmine
                 Net.Api.RedmineManager manager = new Net.Api.RedmineManager("http://redmine.aps-holding.com/", "170139117f26a96e5952339bac48a3bffdd58f37", Net.Api.MimeFormat.xml);

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
                            }
                        }

                        projectsList.Add(projectViewModel);
                    }

                    _projects = SortProjects(projectsList);
                    _flatProjects = FlattenProjects(_projects);
                }

            }
            catch
            {
                //MessageBox.Show("Cannot connect to the Redmine and load projects. Please check your configuration", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private ObservableCollection<ProjectViewModel> FlattenProjects(IList<ProjectViewModel> projects)
        {
            ObservableCollection<ProjectViewModel> flatProjects = new ObservableCollection<ProjectViewModel>();

            Action<ProjectViewModel> getChildren = null;

            getChildren = parent =>
            {
                flatProjects.Add(parent);
                parent.Children.ForEach(p => {
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
                childProjects.ForEach(p => {
                    p.Level = parent.Level + 1;
                    p.Path = parent.Path + " » " + p.Name;
                });

                parent.Children = new ObservableCollection<ProjectViewModel>(childProjects);

                // Recursively call the SetChildren method for each child.
                parent.Children.ForEach(setChildren);
            };

            //Initialize the hierarchical list to root level items
            sortedProjects = new ObservableCollection<ProjectViewModel>(projects.Where(rootItem => rootItem.ParentId == 0).ToList());

            //Call the SetChildren method to set the children on each root level item.
            sortedProjects.ForEach(p => p.Path = p.Name);
            sortedProjects.ForEach(setChildren);

            return sortedProjects;
        }
    }
}
