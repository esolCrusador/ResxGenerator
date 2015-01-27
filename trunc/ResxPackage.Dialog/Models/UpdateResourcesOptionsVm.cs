using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using ResourcesAutogenerate.DomainModels;
using ResxPackage.Dialog.Annotations;

namespace ResxPackage.Dialog.Models
{
    public class UpdateResourcesOptionsVm: INotifyPropertyChanged
    {
        private bool? _embeedSubCultures;
        private bool? _useDefaultContentType;
        private bool? _useDefaultCustomTool;

        public UpdateResourcesOptionsVm()
        {
            RemoveNotSelectedCultures = true;
        }

        public bool RemoveNotSelectedCultures { get; set; }

        public bool? EmbeedSubCultures
        {
            get { return _embeedSubCultures; }
            set
            {
                if (value == false)
                {
                    if (_embeedSubCultures == null)
                        _embeedSubCultures = true;
                    else if (_embeedSubCultures == true)
                        _embeedSubCultures = null;
                    OnPropertyChanged();
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
        }

        public bool? UseDefaultContentType
        {
            get { return _useDefaultContentType; }
            set
            {
                if (_useDefaultContentType == true && value == false)
                {
                    _useDefaultContentType = null;
                    OnPropertyChanged();
                }
                else
                {
                    _useDefaultContentType = value;
                }
            }
        }

        public bool? UseDefaultCustomTool
        {
            get { return _useDefaultCustomTool; }
            set
            {
                if (_useDefaultCustomTool == true && value == false)
                {
                    _useDefaultCustomTool = null;
                    OnPropertyChanged();
                }
                else
                {
                    _useDefaultCustomTool = value;
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

        public UpdateResourcesOptions GetOptions()
        {
            return new UpdateResourcesOptions
            {
                RemoveNotSelectedCultures = RemoveNotSelectedCultures,
                EmbeedSubCultures = EmbeedSubCultures,
                UseDefaultContentType = UseDefaultContentType,
                UseDefaultCustomTool = UseDefaultCustomTool
            };
        }
    }
}
