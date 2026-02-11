using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace CodeReportTracker.Core.Models
{
    /// <summary>
    /// MVVM-friendly view-model for a grid row item.
    /// - Implements INotifyPropertyChanged.
    /// - Exposes row-level commands (OpenLink / Edit) that can be wired by consumers.
    /// - Keeps no dependency on IO.Codes so it can live in the Core project.
    /// </summary>
    public sealed class CodeItem : INotifyPropertyChanged
    {
        private string _number = string.Empty;
        private string _link = string.Empty;
        private string _webType = string.Empty;
        private string _productCategory = string.Empty;
        private string _description = string.Empty;
        private string _productsListed = string.Empty;
        private string _latestCode = string.Empty;
        private string _latestCode_Old = string.Empty;
        private string _issueDate = string.Empty;
        private string _issueDate_Old = string.Empty;
        private string _expirationDate = string.Empty;
        private string _expirationDate_Old = string.Empty;
        private int _downloadProcess;
        private string _lastCheck = string.Empty;
        private bool _hasCheck;
        private bool _hasUpdate;
        private bool _codeExists = true;

        public event PropertyChangedEventHandler? PropertyChanged;

        public Action<string>? OpenLinkAction { get; set; }
        public Action<CodeItem>? EditAction { get; set; }

        // Commands suitable for binding directly from DataTemplates / DataGrid rows.
        private ICommand? _openLinkCommand;
        public ICommand OpenLinkCommand => _openLinkCommand ??= new RelayCommand(_ => ExecuteOpenLink());

        private ICommand? _editCommand;
        public ICommand EditCommand => _editCommand ??= new RelayCommand(_ => ExecuteEdit());

        public string Number
        {
            get => _number;
            set => SetProperty(ref _number, value);
        }

        public string Link
        {
            get => _link;
            set => SetProperty(ref _link, value);
        }

        public string WebType
        {
            get => _webType;
            set => SetProperty(ref _webType, value);
        }

        public string ProductCategory
        {
            get => _productCategory;
            set => SetProperty(ref _productCategory, value);
        }

        public string Description
        {
            get => _description;
            set => SetProperty(ref _description, value);
        }

        public string ProductsListed
        {
            get => _productsListed;
            set => SetProperty(ref _productsListed, value);
        }

        public string LatestCode
        {
            get => _latestCode;
            set => SetProperty(ref _latestCode, value);
        }

        public string LatestCode_Old
        {
            get => _latestCode_Old;
            set => SetProperty(ref _latestCode_Old, value);
        }

        public string IssueDate
        {
            get => _issueDate;
            set => SetProperty(ref _issueDate, value);
        }

        public string IssueDate_Old
        {
            get => _issueDate_Old;
            set => SetProperty(ref _issueDate_Old, value);
        }

        public string ExpirationDate
        {
            get => _expirationDate;
            set => SetProperty(ref _expirationDate, value);
        }

        public string ExpirationDate_Old
        {
            get => _expirationDate_Old;
            set => SetProperty(ref _expirationDate_Old, value);
        }

        public int DownloadProcess
        {
            get => _downloadProcess;
            set => SetProperty(ref _downloadProcess, value);
        }

        public string LastCheck
        {
            get => _lastCheck;
            set => SetProperty(ref _lastCheck, value);
        }

        public bool HasCheck
        {
            get => _hasCheck;
            set => SetProperty(ref _hasCheck, value);
        }

        public bool HasUpdate
        {
            get => _hasUpdate;
            set => SetProperty(ref _hasUpdate, value);
        }

        /// <summary>
        /// New property: true when the code still exists; false otherwise.
        /// Update this from your web-reader / verification logic.
        /// UI can bind to this to switch Latest Code cell background (green/red).
        /// </summary>
        public bool CodeExists
        {
            get => _codeExists;
            set => SetProperty(ref _codeExists, value);
        }

        public CodeItem() { }

        /// <summary>
        /// Convenience ctor to attach actions up-front.
        /// </summary>
        public CodeItem(Action<string>? openLinkAction = null, Action<CodeItem>? editAction = null)
        {
            OpenLinkAction = openLinkAction;
            EditAction = editAction;
        }

        private void ExecuteOpenLink()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(Link))
                {
                    if (OpenLinkAction != null)
                    {
                        OpenLinkAction(Link);
                        return;
                    }

                    // fallback: try to open using shell
                    var psi = new ProcessStartInfo
                    {
                        FileName = Link,
                        UseShellExecute = true,
                        Verb = "open"
                    };
                    Process.Start(psi);
                }
            }
            catch
            {
                // swallow exceptions here to keep command safe; consumer can provide OpenLinkAction to handle errors.
            }
        }

        private void ExecuteEdit()
        {
            try
            {
                EditAction?.Invoke(this);
            }
            catch
            {
                // swallow - consumer handles edit workflow / errors
            }
        }

        private void SetProperty<T>(ref T backingField, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(backingField, value)) return;
            backingField = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #region Simple RelayCommand (internal - small, dependency-free)

        private sealed class RelayCommand : ICommand
        {
            private readonly Action<object?> _execute;
            private readonly Func<object?, bool>? _canExecute;

            public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

            public void Execute(object? parameter) => _execute(parameter);

#pragma warning disable CS0067
            public event EventHandler? CanExecuteChanged;
#pragma warning restore CS0067
        }

        #endregion
    }
}