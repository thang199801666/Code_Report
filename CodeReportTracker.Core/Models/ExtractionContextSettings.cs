using CommunityToolkit.Mvvm.ComponentModel;

namespace CodeReportTracker.Models
{
    public sealed class ExtractionContextSettings : ObservableObject
    {
        private string _latestCode = string.Empty;
        private string _issueDate = string.Empty;
        private string _expirationDate = string.Empty;

        public string LatestCode
        {
            get => _latestCode;
            set => SetProperty(ref _latestCode, value);
        }

        public string IssueDate
        {
            get => _issueDate;
            set => SetProperty(ref _issueDate, value);
        }

        public string ExpirationDate
        {
            get => _expirationDate;
            set => SetProperty(ref _expirationDate, value);
        }
    }
}
