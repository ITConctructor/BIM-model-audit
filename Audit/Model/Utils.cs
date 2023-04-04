using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Xml;
using System.IO;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Audit.Model
{
    public static class Utils
    {
        /// <summary>
        /// Десериализует класс из xml-строки
        /// </summary>
        public static T Deserialize<T>(string content)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
            try
            {
                using (StringReader reader = new StringReader(content))
                {
                    T results = (T)xmlSerializer.Deserialize(reader);
                    return results;
                }
            }
            catch (Exception)
            {
                return default(T);
            }
            return default(T);
        }

        /// <summary>
        /// Создает xml файл
        /// </summary>
        public static string Serialize<T>(T result)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
            string xml;
            using (StringWriter writer = new StringWriter())
            {
                xmlSerializer.Serialize(writer, result);
                xml = writer.ToString();
            }
            return xml;
        }

        /// <summary>
        /// Создает xml файл по указанному пути
        /// </summary>
        public static void Serialize<T>(T result, string filePath)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
            string xml;
            using (StringWriter writer = new StringWriter())
            {
                xmlSerializer.Serialize(writer, result);
                xml = writer.ToString();
            }
            File.AppendAllText(filePath, xml, Encoding.UTF8);
        }

        /// <summary>
        /// Общий для всех классов проверок метод добавления результата проверки элемента в список
        /// </summary>
        public static void AddElementCheckingResult(ElementCheckingResult newResult, BindingList<ElementCheckingResult> results)
        {
            if (results.Remove(newResult))
            {
                newResult.Status = "Активная";
                results.Add(newResult);
            }
            else
            {
                newResult.Status = "Созданная";
                results.Add(newResult);
            }
        }

        /// <summary>
        /// Загружает файловую структуру с ПК для просмотрщика TreeView
        /// </summary>
        private static List<object> GetFileStructure(Bitmap DriveIcon, Bitmap FolderIcon, string type = "")
        {
            List<object> FileStructure = new List<object>();
            foreach (string drive in Directory.GetLogicalDrives())
            {
                List<object> Elements = new List<object>();
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                            DriveIcon.GetHbitmap(),
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                ImageSource DriveImage = mediumImage as ImageSource;
                BitmapSource mediumFolderImage = Imaging.CreateBitmapSourceFromHBitmap(
                            FolderIcon.GetHbitmap(),
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                ImageSource FolderImage = mediumFolderImage as ImageSource;
                try
                {
                    foreach (string Folder in Directory.GetDirectories(drive))
                    {
                        Elements.Add((FolderImage, Folder, ExploreFolders(Folder, FolderIcon, "rvt")));
                    }
                    foreach (string File in Directory.GetFiles(drive))
                    {
                        if (File.EndsWith(type))
                        {
                            System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(File);
                            ImageSourceConverter imageConverter = new ImageSourceConverter();
                            ImageSource FileImage = Imaging.CreateBitmapSourceFromHIcon(icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                            Elements.Add((FileImage, File));
                        }
                    }
                }
                catch (Exception)
                {
                    Elements.Add((DriveImage, drive));
                }
                FileStructure.Add((DriveImage, drive, Elements));
            }
            return FileStructure;
        }

        /// <summary>
        /// Возвращает все папки и файлы, содержащиеся внутри папки
        /// </summary>
        private static List<object> ExploreFolders(string Folder, Bitmap FolderIcon, string type = "")
        {
            List<object> Elements = new List<object>();
            BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                        FolderIcon.GetHbitmap(),
                        IntPtr.Zero,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
            ImageSource FolderImage = mediumImage as ImageSource;
            foreach (string folder in Directory.GetDirectories(Folder))
            {
                try
                {
                    if (Directory.GetDirectories(folder).Length > 0 || Directory.GetFiles(Folder).Length > 0)
                    {
                        List<object> subFolders = ExploreFolders((string)folder, FolderIcon);
                        Elements.Add((FolderImage, folder, subFolders));
                    }
                    else
                    {
                        Elements.Add((FolderImage, folder));
                    }
                }
                catch (Exception)
                {
                    Elements.Add((FolderImage, folder));
                }
            }
            foreach (string File in Directory.GetFiles(Folder))
            {
                if (File.EndsWith(type))
                {
                    System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(File);
                    ImageSourceConverter imageConverter = new ImageSourceConverter();
                    ImageSource FileImage = Imaging.CreateBitmapSourceFromHIcon(icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    Elements.Add((FileImage, File));
                }
            }
            return Elements;
        }

        /// <summary>
        /// Открывает файл rvt и возвращает его в виде объекта Document
        /// </summary>
        /// <param name="filePath"> Путь к файлу</param>
        /// <returns></returns>
        public static Document OpenRvtFile(string filePath)
        {
            UIApplication uiapp = CommandLauncher.uiapp;
            ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            OpenOptions openOptions = new OpenOptions();
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            WorksetConfiguration worksets = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksets);
            Autodesk.Revit.ApplicationServices.Application app = CommandLauncher.app;
            Document doc = app.OpenDocumentFile(path, openOptions);
            return doc;
        }

        /// <summary>
        /// Закрывает файл rvt
        /// </summary>
        /// <param name="doc"></param>
        public static void CloseRvtFile(Document doc)
        {
            if (CommandLauncher.uiapp.ActiveUIDocument != null)
            {
                if (doc.Title != CommandLauncher.uiapp.ActiveUIDocument.Document.Title)
                {
                    doc.Close(false);
                }
            }
            else
            {
                doc.Close(false);
            }
        }
    }
}
