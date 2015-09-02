using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Mach.Wpf.Mvvm
{
    /// <summary>
    /// Base class that implements <see cref="INotifyPropertyChanged"/> interface.
    /// </summary>
    public class NotifyPropertyBase : INotifyPropertyChanged
    {
        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Fires <see cref="PropertyChanged"/> event.
        /// </summary>
        /// <param name="propertyName">Name of the property that has changed.</param>
        protected void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            var handler = PropertyChanged;
            if (handler != null)
                handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}