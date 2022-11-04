using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Audit
{
    [ValueConversion(typeof(ApplicationViewModel.CheckingStatus), typeof(bool))]
    public sealed class CheckingStatusToImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            ApplicationViewModel.CheckingStatus status = (ApplicationViewModel.CheckingStatus)value;
            if (status == ApplicationViewModel.CheckingStatus.CheckingSuccessful)
            {
                Bitmap rareImage = Properties.Resource1.CheckingSuccessful;
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                       rareImage.GetHbitmap(),
                       IntPtr.Zero,
                       Int32Rect.Empty,
                       BitmapSizeOptions.FromEmptyOptions());
                ImageSource source = mediumImage as ImageSource;
                return source;
            }
            else if (status == ApplicationViewModel.CheckingStatus.CheckingFailed)
            {
                Bitmap rareImage = Properties.Resource1.CheckingFailed;
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                       rareImage.GetHbitmap(),
                       IntPtr.Zero,
                       Int32Rect.Empty,
                       BitmapSizeOptions.FromEmptyOptions());
                ImageSource source = mediumImage as ImageSource;
                return source;
            }
            else if (status == ApplicationViewModel.CheckingStatus.NotUpdated)
            {
                Bitmap rareImage = Properties.Resource1.NotUpdated;
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                       rareImage.GetHbitmap(),
                       IntPtr.Zero,
                       Int32Rect.Empty,
                       BitmapSizeOptions.FromEmptyOptions());
                ImageSource source = mediumImage as ImageSource;
                return source;
            }
            else if (status == ApplicationViewModel.CheckingStatus.NotLaunched)
            {
                Bitmap rareImage = Properties.Resource1.NotLaunched;
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                       rareImage.GetHbitmap(),
                       IntPtr.Zero,
                       Int32Rect.Empty,
                       BitmapSizeOptions.FromEmptyOptions());
                ImageSource source = mediumImage as ImageSource;
                return source;
            }
            else
            {
                Bitmap rareImage = Properties.Resource1.CheckingFailed;
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                       rareImage.GetHbitmap(),
                       IntPtr.Zero,
                       Int32Rect.Empty,
                       BitmapSizeOptions.FromEmptyOptions());
                ImageSource source = mediumImage as ImageSource;
                return source;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
