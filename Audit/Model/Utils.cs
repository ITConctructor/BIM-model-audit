using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Xml;
using System.IO;

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
    }
}
