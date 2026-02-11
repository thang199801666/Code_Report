using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using CodeReportTracker.Core.Models;

namespace CodeReportTracker.Components.Persistence
{
    public static class TabPersistence
    {
        private static readonly JsonSerializerOptions DefaultOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public static void SaveTabsToFile(string filePath, IEnumerable<TabModel> tabs)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(tabs, DefaultOptions);
            var tmp = filePath + ".tmp";
            File.WriteAllText(tmp, json);
            File.Copy(tmp, filePath, overwrite: true);
            try { File.Delete(tmp); } catch { /* ignore */ }
        }

        public static List<TabModel>? LoadTabsFromFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            try
            {
                var json = File.ReadAllText(filePath);
                var tabs = JsonSerializer.Deserialize<List<TabModel>>(json, DefaultOptions);
                return tabs ?? new List<TabModel>();
            }
            catch
            {
                return null;
            }
        }
    }
}