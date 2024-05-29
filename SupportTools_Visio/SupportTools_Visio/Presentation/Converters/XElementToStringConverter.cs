using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;
using System.Xml.Linq;

namespace SupportTools_Visio.Presentation.Converters
{
    public class XElementToStringConverter : MarkupExtension, IValueConverter
    {
        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                XElement e = (XElement)value;

                if (parameter != null)
                {
                    return $"{e.Attribute((string)parameter).Value}";
                }
                //return $"{e.Attribute("Name").Value}";

                return e.ToString();
            }
            else
            {
                return "<no export>";
                //return "<none selected>";
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                var foo = XElement.Parse(value.ToString());
                return foo;
            }
            else
            {
                return "ConvertBack value is null";
            }
        }
    }
}
