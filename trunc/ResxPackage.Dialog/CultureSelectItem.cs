using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;

namespace ResxPackage.Dialog
{
    public class CultureSelectItem : INotifyPropertyChanged
    {
        private static readonly int InvariantCultureId = CultureInfo.InvariantCulture.LCID;
        private bool _isVisible;
        private bool _isSelected;

        public CultureSelectItem()
        {
        }

        public CultureSelectItem(CultureInfo culture, bool isSelected)
        {
            CultureId = culture.LCID;
            CultureName = String.Format("{0} ({1})", culture.Name, culture.DisplayName);
            IsSelected = isSelected;
            InitiallySelected = isSelected;
            IsDefault = CultureId == InvariantCultureId;
            IsVisible = true;
        }

        public CultureSelectItem(CultureInfo culture)
            : this(culture, false)
        {
        }

        public int CultureId { get; set; }

        public string CultureName { get; set; }

        public bool InitiallySelected { get; set; }

        public bool IsDefault { get; private set; }

        public bool IsEnabled{get { return !IsDefault; }}

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public bool IsVisible
        {
            get { return _isVisible; }
            set
            {
                _isVisible = value;
                OnPropertyChanged();
                OnPropertyChanged("Visibility");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
