using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Image = System.Drawing.Image;

namespace Audit
{
    [ValueConversion(typeof(string), typeof(bool))]
    internal class HeaderToImageConverter : IValueConverter
    {
        public static HeaderToImageConverter Instance = new HeaderToImageConverter();
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string filepath = value as string;
            ImageSourceConverter converter = new ImageSourceConverter();
            try
            {
                if (filepath.Length < 4)
                {
                    Bitmap rareImage = Properties.Resource1.driveImage;
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
                    try
                    {
                        System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(filepath);
                        ImageSourceConverter imageConverter = new ImageSourceConverter();
                        ImageSource filesource = Imaging.CreateBitmapSourceFromHIcon(icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                        return filesource;
                    }
                    catch (Exception)
                    {
                        Bitmap rareImage = Properties.Resource1.folderImage;
                        BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                               rareImage.GetHbitmap(),
                               IntPtr.Zero,
                               Int32Rect.Empty,
                               BitmapSizeOptions.FromEmptyOptions());
                        ImageSource source = mediumImage as ImageSource;
                        return source;
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException("Cannot convert back");
        }
    }
}
