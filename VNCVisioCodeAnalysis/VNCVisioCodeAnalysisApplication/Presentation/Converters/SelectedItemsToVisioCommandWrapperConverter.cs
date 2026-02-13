using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace VNCVisioToolsApplication.Presentation.Converters
{
    public class SelectedItemsToVisioCommandWrapperConverter: MarkupExtension, IValueConverter
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        object IValueConverter.Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                //return ((VisioCommandWrapper)value).Name;
            }

            return null;
        }

        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {

            var valueType = value.GetType();

            //var returnValue = (VisioCommandWrapper)value;

            //return returnValue;

            return null;
        }
    }
}
