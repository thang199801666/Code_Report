using System;

namespace CodeReportTracker.Core.Models
{
    public class CodeItemDto
    {
        public string Number { get; set; } = string.Empty;
        public string Link { get; set; } = string.Empty;
        public string WebType { get; set; } = string.Empty;
        public string ProductCategory { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string ProductsListed { get; set; } = string.Empty;
        public string LatestCode { get; set; } = string.Empty;
        public string LatestCode_Old { get; set; } = string.Empty;
        public string IssueDate { get; set; } = string.Empty;
        public string IssueDate_Old { get; set; } = string.Empty;
        public string ExpirationDate { get; set; } = string.Empty;
        public string ExpirationDate_Old { get; set; } = string.Empty;
        public int DownloadProcess { get; set; }
        public string LastCheck { get; set; } = string.Empty;
        public bool HasCheck { get; set; }
        public bool HasUpdate { get; set; }

        public static CodeItemDto FromCodeItem(CodeItem ci)
        {
            if (ci == null) return new CodeItemDto();
            return new CodeItemDto
            {
                Number = ci.Number ?? string.Empty,
                Link = ci.Link ?? string.Empty,
                WebType = ci.WebType ?? string.Empty,
                ProductCategory = ci.ProductCategory ?? string.Empty,
                Description = ci.Description ?? string.Empty,
                ProductsListed = ci.ProductsListed ?? string.Empty,
                LatestCode = ci.LatestCode ?? string.Empty,
                LatestCode_Old = ci.LatestCode_Old ?? string.Empty,
                IssueDate = ci.IssueDate ?? string.Empty,
                IssueDate_Old = ci.IssueDate_Old ?? string.Empty,
                ExpirationDate = ci.ExpirationDate ?? string.Empty,
                ExpirationDate_Old = ci.ExpirationDate_Old ?? string.Empty,
                DownloadProcess = ci.DownloadProcess,
                LastCheck = ci.LastCheck ?? string.Empty,
                HasCheck = ci.HasCheck,
                HasUpdate = ci.HasUpdate
            };
        }

        public CodeItem ToCodeItem()
        {
            var ci = new CodeItem();
            ci.Number = this.Number ?? string.Empty;
            ci.Link = this.Link ?? string.Empty;
            ci.WebType = this.WebType ?? string.Empty;
            ci.ProductCategory = this.ProductCategory ?? string.Empty;
            ci.Description = this.Description ?? string.Empty;
            ci.ProductsListed = this.ProductsListed ?? string.Empty;
            ci.LatestCode = this.LatestCode ?? string.Empty;
            ci.LatestCode_Old = this.LatestCode_Old ?? string.Empty;
            ci.IssueDate = this.IssueDate ?? string.Empty;
            ci.IssueDate_Old = this.IssueDate_Old ?? string.Empty;
            ci.ExpirationDate = this.ExpirationDate ?? string.Empty;
            ci.ExpirationDate_Old = this.ExpirationDate_Old ?? string.Empty;
            ci.DownloadProcess = this.DownloadProcess;
            ci.LastCheck = this.LastCheck ?? string.Empty;
            ci.HasCheck = this.HasCheck;
            ci.HasUpdate = this.HasUpdate;
            return ci;
        }
    }
}