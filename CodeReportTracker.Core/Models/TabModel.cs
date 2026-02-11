using System.Collections.Generic;

namespace CodeReportTracker.Core.Models
{
    public class TabModel
    {
        public string Header { get; set; } = string.Empty;
        public List<CodeItem> Items { get; set; } = new List<CodeItem>();
    }
}