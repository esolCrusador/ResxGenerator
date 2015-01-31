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

                OnPropertyChanged("IsResxSyncSelected");
                OnPropertyChanged("IsExcelSelected");
                OnPropertyChanged("IsGoogleSheetsSelected");
            }
        }

        public bool IsResxSyncSelected
        {
            get { return ExternalSource == ExternalSource.Sync; }
            set
            {
                if (value)
                {
                    ExternalSource=ExternalSource.Sync;
                }
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

        public bool IsGoogleSheetsSelected
        {
            get { return ExternalSource == ExternalSource.GSheets; }
            set
            {
                if (value)
                {
                    ExternalSource = ExternalSource.GSheets;
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
