using System;
using System.Globalization;

namespace ResxPackage.Dialog
{
    public class CultureSelectItem
    {
        public CultureSelectItem()
        {
        }

        public CultureSelectItem(CultureInfo culture, bool isSelected)
        {
            CultureId = culture.LCID;
            CultureName = String.Format("{0} ({1})", culture.Name, culture.DisplayName);
            IsSelected = isSelected;
        }

        public CultureSelectItem(CultureInfo culture)
            : this(culture, false)
        {
        }

        public int CultureId { get; set; }

        public string CultureName { get; set; }

        public bool IsSelected { get; set; }
    }
}
