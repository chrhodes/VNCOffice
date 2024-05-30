using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;
using System.Xml;

using SupportTools_Excel.AzureDevOpsExplorer.Presentation.ModelWrappers;

namespace SupportTools_Excel.Presentation.Converters
{
    public class SelectedItemsConverter2 : MarkupExtension, IValueConverter
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        object IValueConverter.Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
                return new List<object>((IEnumerable<string>)value);

            return null;
        }

        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var valueType = value.GetType();

            var returnValue = ((XmlElement)value).Attributes["Name"].Value;

            return value;
        }
    }
}
