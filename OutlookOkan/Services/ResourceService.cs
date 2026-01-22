using OutlookOkan.Properties;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace OutlookOkan.Services
{
    internal sealed class ResourceService : INotifyPropertyChanged
    {
        // Thể hiện Singleton.
        public static ResourceService Instance { get; } = new ResourceService();
        private ResourceService() { }

        public Resources Resources { get; } = new Resources();

        public event PropertyChangedEventHandler? PropertyChanged;
        private void RaisePropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void ChangeCulture(string name)
        {
            if (string.IsNullOrEmpty(name)) return;
            if (Resources.Culture != null && Resources.Culture.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) return;

            Resources.Culture = CultureInfo.GetCultureInfo(name);
            RaisePropertyChanged(nameof(Resources));
        }
    }
}