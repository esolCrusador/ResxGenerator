using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ResxPackage.Dialog.Annotations;

namespace ResxPackage.Dialog.Models
{
    public class ExternalSourceVm: INotifyPropertyChanged
    {
        private ExternalSource _externalSource;

        public ExternalSourceVm(ExternalSource externalSource)
        {
            _externalSource = externalSource;
        }

        public ExternalSource ExternalSource
        {
            get { return _externalSource; }
            set
            {
                _externalSource = value;
                OnPropertyChanged();

                OnPropertyChanged("IsExcelSelected");
                OnPropertyChanged("IsGoogleDriveSelected");
            }
        }

        public bool IsExcelSelected
        {
            get { return ExternalSource == ExternalSource.Excel; }
            set
            {
                if (value)
                {
                    ExternalSource = ExternalSource.Excel;
                }
            }
        }

        public bool IsGoogleDriveSelected
        {
            get { return ExternalSource == ExternalSource.GDrive; }
            set
            {
                if (value)
                {
                    ExternalSource = ExternalSource.GDrive;
                }
            }
        }
        
        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
