using CommunityToolkit.Mvvm.ComponentModel;

namespace CodeReportTracker.Models
{
    public class SettingEntry : ObservableObject
    {
        private string _name = string.Empty;
        private string _type = string.Empty;
        private string _link = string.Empty;
        private string _pdfFolder = string.Empty;

        // Friendly name for the web/link (e.g. "IAPMO", "ICC-ES")
        public string Name
        {
            get => _name;
            set => SetProperty(ref _name, value);
        }

        // Type/category for the entry (e.g. "ER", "ESR", etc.)
        public string Type
        {
            get => _type;
            set => SetProperty(ref _type, value);
        }

        // The web link for the entry
        public string Link
        {
            get => _link;
            set => SetProperty(ref _link, value);
        }

        // Optional local PDF folder path associated with this web entry
        public string PdfFolder
        {
            get => _pdfFolder;
            set => SetProperty(ref _pdfFolder, value);
        }
    }
}