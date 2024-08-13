using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;

namespace Code_Report
{
    class IOData
    {
        public static void WriteToBinaryFile<T>(string filePath, T objectToWrite, bool append = false)
        {
            using (Stream stream = File.Open(filePath, append ? FileMode.Append : FileMode.Create))
            {
                var binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                binaryFormatter.Serialize(stream, objectToWrite);
            }
        }
        public static void WriteToJsonFile<T>(string filePath, T objectToWrite, bool append = false)
        {
            File.WriteAllText(filePath, JsonConvert.SerializeObject(objectToWrite));
            using (StreamWriter file = File.CreateText(filePath))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, objectToWrite);
            }
        }

        public static void WriteToXmlFile<T>(String path, T file)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));

            using (StreamWriter writer = new StreamWriter(path))
            {
                serializer.Serialize(writer, file);
            }
        }
        public static T ReadFromBinaryFile<T>(string filePath)
        {
            using (Stream stream = File.Open(filePath, FileMode.Open))
            {
                var binaryFormatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
                return (T)binaryFormatter.Deserialize(stream);
            }
        }
        public static DateTime convertDateTime(string input, string dateFormat)
        {
            var culture = System.Globalization.CultureInfo.CurrentCulture;
            DateTime dateTime = DateTime.ParseExact(input, dateFormat, culture);
            return dateTime;
        }
        public static List<int> ColorToList(Color color)
        {
            return new List<int> { color.A, color.R, color.G, color.B  };
        }
        public static Color ListToColor(List<int> list)
        {
            return Color.FromArgb(list[0], list[1], list[2], list[3]);
        }
    }
}
