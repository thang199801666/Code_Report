using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace WinUx.Controls
{
    public partial class MessageBoxViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _messageText = string.Empty;

        [ObservableProperty]
        private string _okButtonText = "OK";
        [ObservableProperty]
        private string _cancelButtonText = "Cancel";
        [ObservableProperty]
        private string _yesButtonText = "Yes";
        [ObservableProperty]
        private string _noButtonText = "No";
        [ObservableProperty]
        private string _yesToAllButtonText = "Yes to All";

        [ObservableProperty]
        private bool _showOkButton = false;
        [ObservableProperty]
        private bool _showCancelButton = false;
        [ObservableProperty]
        private bool _showYesButton = false;
        [ObservableProperty]
        private bool _showNoButton = false;
        [ObservableProperty]
        private bool _showYesToAllButton = false;

        [ObservableProperty]
        private string? _iconUri;

        public MessageBoxResults DialogResult { get; private set; } = MessageBoxResults.None;

        public event Action? RequestClose;

        public void Setup(string message, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            MessageText = message ?? string.Empty;

            // Use centralized config from MessageBoxButtons to decide visibility.
            var cfg = buttons.GetConfig();

            ShowOkButton = cfg.ShowOk;
            ShowCancelButton = cfg.ShowCancel;
            ShowYesButton = cfg.ShowYes;
            ShowNoButton = cfg.ShowNo;
            ShowYesToAllButton = cfg.ShowYesToAll;

            // Apply sensible default labels only when the corresponding button is shown.
            if (ShowOkButton) OkButtonText = string.IsNullOrWhiteSpace(OkButtonText) ? "OK" : OkButtonText;
            if (ShowCancelButton) CancelButtonText = string.IsNullOrWhiteSpace(CancelButtonText) ? "Cancel" : CancelButtonText;
            if (ShowYesButton) YesButtonText = string.IsNullOrWhiteSpace(YesButtonText) ? "Yes" : YesButtonText;
            if (ShowNoButton) NoButtonText = string.IsNullOrWhiteSpace(NoButtonText) ? "No" : NoButtonText;
            if (ShowYesToAllButton) YesToAllButtonText = string.IsNullOrWhiteSpace(YesToAllButtonText) ? "Yes to All" : YesToAllButtonText;

            IconUri = MessageBoxHelper.GetIconUri(icon);
        }

        [RelayCommand]
        public void Ok()
        {
            DialogResult = MessageBoxResults.OK;
            RequestClose?.Invoke();
        }

        [RelayCommand]
        public void Cancel()
        {
            DialogResult = MessageBoxResults.Cancel;
            RequestClose?.Invoke();
        }

        [RelayCommand]
        public void Yes()
        {
            DialogResult = MessageBoxResults.Yes;
            RequestClose?.Invoke();
        }

        [RelayCommand]
        public void No()
        {
            DialogResult = MessageBoxResults.No;
            RequestClose?.Invoke();
        }

        [RelayCommand]
        public void YesToAll()
        {
            DialogResult = MessageBoxResults.YesToAll;
            RequestClose?.Invoke();
        }
    }
}