using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace Audit
{
    [ValueConversion(typeof(ApplicationViewModel.CheckingStatus), typeof(string))]
    public sealed class CheckingStatusToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ApplicationViewModel.CheckingStatus status = (ApplicationViewModel.CheckingStatus)value;
            if (status == ApplicationViewModel.CheckingStatus.CheckingSuccessful)
            {
                return "Проверка пройдена";
            }
            else if (status == ApplicationViewModel.CheckingStatus.CheckingFailed)
            {
                return "Проверка не пройдена";
            }
            else if (status == ApplicationViewModel.CheckingStatus.NotUpdated)
            {
                return "Результаты проверки устарели";
            }
            else if (status == ApplicationViewModel.CheckingStatus.NotLaunched)
            {
                return "Файл добавлен впервые, проверки не проводились";
            }
            else
            {
                return "Проверка не пройдена";
            }

        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
