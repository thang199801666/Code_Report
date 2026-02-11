using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using CodeReportTracker.Core.Models;

namespace CodeReportTracker.Core.Persistence
{
    /// <summary>
    /// Versioned binary serializer for application data (.crp).
    /// Format v1:
    /// - 4 bytes ASCII magic "CRPB"
    /// - 1 byte  version (1)
    /// - Int32  tab count
    /// For each tab:
    ///   - string header (Int32 length, UTF8 bytes, -1 = null)
    ///   - Int32 itemCount
    ///   For each item:
    ///     - string Number
    ///     - string Link
    ///     - string WebType
    ///     - string ProductCategory
    ///     - string Description
    ///     - string ProductsListed
    ///     - string LatestCode
    ///     - string LatestCode_Old
    ///     - string IssueDate
    ///     - string IssueDate_Old
    ///     - string ExpirationDate
    ///     - string ExpirationDate_Old
    ///     - Int32 DownloadProcess
    ///     - string LastCheck
    ///     - byte HasCheck (0/1)
    ///     - byte HasUpdate (0/1)
    ///
    /// Format v2: same as v1 but appends:
    ///     - byte CodeExists (0/1)
    /// 
    /// Note: Save() intentionally no longer persists HasCheck/HasUpdate or DownloadProcess values.
    ///       Those flags/values are always written as 0 when saving so files do not retain
    ///       previous "checked/updated" or in-progress download state across sessions.
    /// </summary>
    public static class BinaryCrpSerializer
    {
        private const string Magic = "CRPB";
        private const byte CurrentVersion = 2;

        public static void Save(string filePath, IEnumerable<TabModel> tabs)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
            using var bw = new BinaryWriter(fs, Encoding.UTF8, leaveOpen: false);

            // header
            bw.Write(Encoding.ASCII.GetBytes(Magic));
            bw.Write(CurrentVersion);

            var list = tabs as IReadOnlyCollection<TabModel> ?? new List<TabModel>(tabs);
            bw.Write(list.Count);

            foreach (var tab in list)
            {
                WriteString(bw, tab?.Header);

                var items = tab?.Items ?? new List<CodeItem>();
                bw.Write(items.Count);
                foreach (var it in items)
                {
                    WriteString(bw, it?.Number);
                    WriteString(bw, it?.Link);
                    WriteString(bw, it?.WebType);
                    WriteString(bw, it?.ProductCategory);
                    WriteString(bw, it?.Description);
                    WriteString(bw, it?.ProductsListed);
                    WriteString(bw, it?.LatestCode);
                    WriteString(bw, it?.LatestCode_Old);
                    WriteString(bw, it?.IssueDate);
                    WriteString(bw, it?.IssueDate_Old);
                    WriteString(bw, it?.ExpirationDate);
                    WriteString(bw, it?.ExpirationDate_Old);

                    // Do NOT persist DownloadProcess - always write zero so files do not retain in-progress download state.
                    bw.Write(0);

                    WriteString(bw, it?.LastCheck);

                    // Do NOT persist HasCheck / HasUpdate state.
                    // Always write 0 so files do not remember checked/updated flags.
                    bw.Write((byte)0); // HasCheck
                    bw.Write((byte)0); // HasUpdate

                    // version 2 field
                    bw.Write(it?.CodeExists == true ? (byte)1 : (byte)0);
                }
            }

            bw.Flush();
        }

        public static List<TabModel>? Load(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return null;
            if (!File.Exists(filePath)) return null;

            try
            {
                using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var br = new BinaryReader(fs, Encoding.UTF8, leaveOpen: false);

                var magicBytes = br.ReadBytes(4);
                if (magicBytes.Length != 4) return null;
                var magic = Encoding.ASCII.GetString(magicBytes);
                if (!string.Equals(magic, Magic, StringComparison.Ordinal)) return null;

                var version = br.ReadByte();

                if (version == 1)
                {
                    return LoadV1(br);
                }

                if (version == 2)
                {
                    return LoadV2(br);
                }

                // future: add migration handlers for newer versions
                throw new NotSupportedException($"Unsupported CRP binary version {version}");
            }
            catch
            {
                return null;
            }
        }

        private static List<TabModel>? LoadV1(BinaryReader br)
        {
            var tabs = new List<TabModel>();
            var tabCount = br.ReadInt32();
            for (int ti = 0; ti < tabCount; ti++)
            {
                var header = ReadString(br) ?? string.Empty;
                var itemCount = br.ReadInt32();
                var items = new List<CodeItem>(itemCount);

                for (int ii = 0; ii < itemCount; ii++)
                {
                    var ci = new CodeItem
                    {
                        Number = ReadString(br) ?? string.Empty,
                        Link = ReadString(br) ?? string.Empty,
                        WebType = ReadString(br) ?? string.Empty,
                        ProductCategory = ReadString(br) ?? string.Empty,
                        Description = ReadString(br) ?? string.Empty,
                        ProductsListed = ReadString(br) ?? string.Empty,
                        LatestCode = ReadString(br) ?? string.Empty,
                        LatestCode_Old = ReadString(br) ?? string.Empty,
                        IssueDate = ReadString(br) ?? string.Empty,
                        IssueDate_Old = ReadString(br) ?? string.Empty,
                        ExpirationDate = ReadString(br) ?? string.Empty,
                        ExpirationDate_Old = ReadString(br) ?? string.Empty,
                        DownloadProcess = br.ReadInt32(),
                        LastCheck = ReadString(br) ?? string.Empty,
                        HasCheck = br.ReadByte() != 0,
                        HasUpdate = br.ReadByte() != 0,
                        CodeExists = false
                    };
                    items.Add(ci);
                }

                tabs.Add(new TabModel { Header = header, Items = items });
            }

            return tabs;
        }

        private static List<TabModel>? LoadV2(BinaryReader br)
        {
            var tabs = new List<TabModel>();
            var tabCount = br.ReadInt32();
            for (int ti = 0; ti < tabCount; ti++)
            {
                var header = ReadString(br) ?? string.Empty;
                var itemCount = br.ReadInt32();
                var items = new List<CodeItem>(itemCount);

                for (int ii = 0; ii < itemCount; ii++)
                {
                    var ci = new CodeItem
                    {
                        Number = ReadString(br) ?? string.Empty,
                        Link = ReadString(br) ?? string.Empty,
                        WebType = ReadString(br) ?? string.Empty,
                        ProductCategory = ReadString(br) ?? string.Empty,
                        Description = ReadString(br) ?? string.Empty,
                        ProductsListed = ReadString(br) ?? string.Empty,
                        LatestCode = ReadString(br) ?? string.Empty,
                        LatestCode_Old = ReadString(br) ?? string.Empty,
                        IssueDate = ReadString(br) ?? string.Empty,
                        IssueDate_Old = ReadString(br) ?? string.Empty,
                        ExpirationDate = ReadString(br) ?? string.Empty,
                        ExpirationDate_Old = ReadString(br) ?? string.Empty,
                        DownloadProcess = br.ReadInt32(),
                        LastCheck = ReadString(br) ?? string.Empty,
                        HasCheck = br.ReadByte() != 0,
                        HasUpdate = br.ReadByte() != 0,
                        CodeExists = br.ReadByte() != 0
                    };
                    items.Add(ci);
                }

                tabs.Add(new TabModel { Header = header, Items = items });
            }

            return tabs;
        }

        private static void WriteString(BinaryWriter bw, string? value)
        {
            if (value == null)
            {
                bw.Write(-1);
                return;
            }

            var bytes = Encoding.UTF8.GetBytes(value);
            bw.Write(bytes.Length);
            bw.Write(bytes);
        }

        private static string? ReadString(BinaryReader br)
        {
            var len = br.ReadInt32();
            if (len < 0) return null;
            var bytes = br.ReadBytes(len);
            return Encoding.UTF8.GetString(bytes);
        }
    }
}