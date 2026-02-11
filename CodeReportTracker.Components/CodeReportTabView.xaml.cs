using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using CodeReportTracker.Components.Persistence;
using CodeReportTracker.Components.ViewModels;
using CodeReportTracker.Core.Models;
using WinUx.Controls;
using System.Diagnostics;

namespace CodeReportTracker.Components
{
    public partial class CodeReportTabView : UserControl
    {
        private readonly ObservableCollection<object> _fallbackItems = new();
        private INotifyCollectionChanged? _itemsSourceNotifier;
        private int _newTabCounter = 1;
        private readonly Dictionary<TabItem, string?> _originalHeaders = new();

        // base directory for per-tab "Pdf Files" folders (use exe path)
        private readonly string _exeBaseDir;

        // drag/drop state
        private Point _dragStartPoint;
        private bool _isDragging;
        private object? _draggedData;

        // subscriptions per TabViewModel (collection, header, item property changes)
        private readonly Dictionary<TabViewModel, (NotifyCollectionChangedEventHandler collectionHandler, PropertyChangedEventHandler headerHandler, PropertyChangedEventHandler itemHandler)> _tabSubscriptions
            = new();

        // guard to prevent re-entrant property-change loops when synchronizing SelectedItem
        private bool _suppressSelectedItemChange;

        public CodeReportTabView()
        {
            InitializeComponent();

            _exeBaseDir = AppContext.BaseDirectory ?? Directory.GetCurrentDirectory();

            // Hook selection and generator events after load so ItemContainerGenerator is available.
            Dispatcher.BeginInvoke(() =>
            {
                if (PART_TabControl != null)
                {
                    PART_TabControl.SelectionChanged += TabControl_SelectionChanged;
                    PART_TabControl.ItemContainerGenerator.StatusChanged += (_, _) => OnItemContainerGeneratorStatusChanged();
                    AttachAddButtonHandler();
                }
            }, System.Windows.Threading.DispatcherPriority.Loaded);

            // default startup state
            var initialVm = new TabViewModel("New Tab", null);
            // ensure pdf folder for the default tab using exe base directory
            try { initialVm.InitializePdfFolder(_exeBaseDir); } catch { }
            _fallbackItems.Add(initialVm);
            UpdateEffectiveItems();

            SelectedItem = initialVm;
            Dispatcher.BeginInvoke(() => { if (PART_TabControl != null) PART_TabControl.SelectedItem = initialVm; }, System.Windows.Threading.DispatcherPriority.Loaded);

            Loaded += (_, _) => OnItemContainerGeneratorStatusChanged();

            _fallbackItems.CollectionChanged += (_, __) =>
            {
                UpdateEffectiveItems();
                EnsureUniqueHeaders(EffectiveItems);
                UpdateTabSubscriptions();
                InvokeSaveCommand(null);
            };

            AddCommand ??= new DelegateCommand(_ => DefaultAdd());
            CloseCommand ??= new DelegateCommand(item => DefaultClose(item));
            RenameCommand ??= new DelegateCommand(item => DefaultBeginRename(item));
        }

        public List<TabModel> GetTabModels() => CollectTabModels();

        private void OnItemContainerGeneratorStatusChanged()
        {
            AttachContextMenus();
            EnsurePerTabTables();
            AttachAddButtonHandler();
        }

        private void TabControl_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (PART_TabControl == null) return;

            var current = PART_TabControl.SelectedItem;
            if (current == null)
            {
                if (!_suppressSelectedItemChange)
                {
                    try
                    {
                        _suppressSelectedItemChange = true;
                        SelectedItem = null;
                    }
                    finally { _suppressSelectedItemChange = false; }
                }
                return;
            }

            if (current is TabViewModel || current is TabModel)
            {
                if (!_suppressSelectedItemChange)
                {
                    try
                    {
                        _suppressSelectedItemChange = true;
                        SelectedItem = current;
                    }
                    finally { _suppressSelectedItemChange = false; }
                }
            }
        }

        private void AttachContextMenus()
        {
            if (PART_TabControl == null) return;

            var gen = PART_TabControl.ItemContainerGenerator;
            for (int i = 0; i < PART_TabControl.Items.Count; i++)
            {
                var item = PART_TabControl.Items[i];
                if (gen.ContainerFromIndex(i) is not TabItem container) continue;
                if (container.ContextMenu != null) continue;

                var cm = new ContextMenu();

                // Rename menu item (existing)
                var mi = new MenuItem { Header = "Rename", Tag = item };
                var icon = LoadIcon("pack://application:,,,/WinUx.Styles;component/Resources/Rename.png");
                if (icon != null) mi.Icon = icon;
                mi.Click += MenuItem_Rename_Click;
                cm.Items.Add(mi);

                // Place "Open Pdf Folder" immediately below Rename (with a separator)
                cm.Items.Add(new Separator());
                var miOpen = new MenuItem { Header = "Open Pdf Folder", Tag = item };
                var openIcon = LoadIcon("pack://application:,,,/WinUx.Styles;component/Resources/Folder.png");
                if (openIcon != null) miOpen.Icon = openIcon;
                miOpen.Click += MenuItem_OpenPdfFolder_Click;
                cm.Items.Add(miOpen);

                container.ContextMenu = cm;
            }
        }

        /// <summary>
        /// Ensure a CodeReportTable instance is assigned to each TabItem's Content when appropriate.
        /// Prefer TabViewModel.Content if present; otherwise create and bind a new table instance.
        /// </summary>
        private void EnsurePerTabTables()
        {
            if (PART_TabControl == null) return;

            var gen = PART_TabControl.ItemContainerGenerator;
            for (int i = 0; i < PART_TabControl.Items.Count; i++)
            {
                var item = PART_TabControl.Items[i];
                if (gen.ContainerFromIndex(i) is not TabItem container) continue;

                if (container.Content is CodeReportTable) continue;

                try
                {
                    if (item is TabViewModel tvm)
                    {
                        if (tvm.Content is CodeReportTable hostedExisting)
                        {
                            container.Content = hostedExisting;
                        }
                        else
                        {
                            var table = new CodeReportTable { DataContext = tvm };
                            Bind(table, CodeReportTable.ItemsSourceProperty, nameof(TabViewModel.Items), tvm, BindingMode.OneWay);
                            Bind(table, CodeReportTable.SelectedItemProperty, nameof(TabViewModel.SelectedItem), tvm, BindingMode.TwoWay);
                            container.Content = table;
                            SafeAction(() => tvm.Content = table);
                        }
                    }
                    else
                    {
                        var type = item?.GetType();
                        if (type != null && type.GetProperty("Items") != null)
                        {
                            var table = new CodeReportTable { DataContext = item as DependencyObject };
                            Bind(table, CodeReportTable.ItemsSourceProperty, "Items", item, BindingMode.OneWay);
                            container.Content = table;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"EnsurePerTabTables: failed to instantiate table for item index {i}: {ex.Message}");
                }
            }
        }

        private void AttachAddButtonHandler()
        {
            if (PART_TabControl == null) return;

            SafeAction(() =>
            {
                PART_TabControl.ApplyTemplate();
                var addBtn = PART_TabControl.Template.FindName("PART_AddButton", PART_TabControl) as Button;
                if (addBtn != null)
                {
                    addBtn.Click -= AddButton_Click;
                    addBtn.Click += AddButton_Click;
                }
            });
        }

        private void AddButton_Click(object? sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Command != null) return;
            OnAddRequested();
        }

        public IEnumerable? ItemsSource
        {
            get => (IEnumerable?)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }

        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(nameof(ItemsSource), typeof(IEnumerable), typeof(CodeReportTabView),
                new PropertyMetadata(null, OnItemsSourceChanged));

        private static void OnItemsSourceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not CodeReportTabView ctl) return;

            ctl.UnsubscribeFromOldNotifier(e.OldValue as IEnumerable);
            ctl.SubscribeToNotifier(e.NewValue as IEnumerable);

            ctl.UpdateEffectiveItems();
            ctl.EnsureUniqueHeaders(ctl.EffectiveItems);
            ctl.UpdateTabSubscriptions();
        }

        public object? SelectedItem
        {
            get
            {
                if (Dispatcher?.CheckAccess() ?? false) return GetValue(SelectedItemProperty);
                return Dispatcher.Invoke(() => GetValue(SelectedItemProperty));
            }
            set
            {
                if (value != null && !(value is TabViewModel) && !(value is TabModel)) return;

                if (Dispatcher?.CheckAccess() ?? false)
                {
                    SetValue(SelectedItemProperty, value);
                    return;
                }

                Dispatcher.Invoke(() => SetValue(SelectedItemProperty, value));
            }
        }

        public static readonly DependencyProperty SelectedItemProperty = DependencyProperty.Register(
                                                                        nameof(SelectedItem),
                                                                        typeof(object),
                                                                        typeof(CodeReportTabView),
                                                                        new FrameworkPropertyMetadata(
                                                                            null,
                                                                            FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
                                                                            OnSelectedItemPropertyChanged,
                                                                            CoerceSelectedItem));

        private static object? CoerceSelectedItem(DependencyObject d, object? baseValue)
        {
            // Reject non-tab values (CodeItem etc.) so bindings cannot push a CodeItem into the control SelectedItem DP.
            if (baseValue == null) return null;
            if (baseValue is TabViewModel || baseValue is TabModel) return baseValue;
            return null;
        }

        private static void OnSelectedItemPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not CodeReportTabView ctl) return;
            if (ctl._suppressSelectedItemChange) return;

            var newVal = e.NewValue;
            if (newVal == null || newVal is TabViewModel || newVal is TabModel)
            {
                if (ctl.PART_TabControl != null && !ReferenceEquals(ctl.PART_TabControl.SelectedItem, newVal))
                {
                    try
                    {
                        ctl._suppressSelectedItemChange = true;
                        ctl.PART_TabControl.SelectedItem = newVal;
                    }
                    catch (Exception ex) { Debug.WriteLine($"OnSelectedItemPropertyChanged: sync failed: {ex.Message}"); }
                    finally { ctl._suppressSelectedItemChange = false; }
                }
            }
            else
            {
                object? restore = e.OldValue;
                if (!(restore is TabViewModel) && !(restore is TabModel)) restore = null;

                try
                {
                    ctl._suppressSelectedItemChange = true;
                    ctl.SetCurrentValue(SelectedItemProperty, restore);
                }
                catch (Exception ex) { Debug.WriteLine($"OnSelectedItemPropertyChanged: restore failed: {ex.Message}"); }
                finally { ctl._suppressSelectedItemChange = false; }
            }
        }

        public DataTemplate? HeaderTemplate
        {
            get => (DataTemplate?)GetValue(HeaderTemplateProperty);
            set => SetValue(HeaderTemplateProperty, value);
        }

        public static readonly DependencyProperty HeaderTemplateProperty =
            DependencyProperty.Register(nameof(HeaderTemplate), typeof(DataTemplate), typeof(CodeReportTabView), new PropertyMetadata(null));

        public DataTemplate? ContentTemplate
        {
            get => (DataTemplate?)GetValue(ContentTemplateProperty);
            set => SetValue(ContentTemplateProperty, value);
        }

        public static readonly DependencyProperty ContentTemplateProperty =
            DependencyProperty.Register(nameof(ContentTemplate), typeof(DataTemplate), typeof(CodeReportTabView), new PropertyMetadata(null));

        public ICommand? CloseCommand
        {
            get => (ICommand?)GetValue(CloseCommandProperty);
            set => SetValue(CloseCommandProperty, value);
        }

        public static readonly DependencyProperty CloseCommandProperty =
            DependencyProperty.Register(nameof(CloseCommand), typeof(ICommand), typeof(CodeReportTabView), new PropertyMetadata(null));

        public ICommand? AddCommand
        {
            get => (ICommand?)GetValue(AddCommandProperty);
            set => SetValue(AddCommandProperty, value);
        }

        public static readonly DependencyProperty AddCommandProperty =
            DependencyProperty.Register(nameof(AddCommand), typeof(ICommand), typeof(CodeReportTabView), new PropertyMetadata(null));

        public ICommand? RenameCommand
        {
            get => (ICommand?)GetValue(RenameCommandProperty);
            set => SetValue(RenameCommandProperty, value);
        }

        public static readonly DependencyProperty RenameCommandProperty =
            DependencyProperty.Register(nameof(RenameCommand), typeof(ICommand), typeof(CodeReportTabView), new PropertyMetadata(null));

        public ICommand? SaveDataCommand
        {
            get => (ICommand?)GetValue(SaveDataCommandProperty);
            set => SetValue(SaveDataCommandProperty, value);
        }

        public static readonly DependencyProperty SaveDataCommandProperty =
            DependencyProperty.Register(nameof(SaveDataCommand), typeof(ICommand), typeof(CodeReportTabView), new PropertyMetadata(null));

        public IEnumerable EffectiveItems
        {
            get => (IEnumerable)GetValue(EffectiveItemsProperty);
            private set => SetValue(EffectiveItemsPropertyKey, value);
        }

        private static readonly DependencyPropertyKey EffectiveItemsPropertyKey =
            DependencyProperty.RegisterReadOnly(nameof(EffectiveItems), typeof(IEnumerable), typeof(CodeReportTabView), new PropertyMetadata(null));

        public static readonly DependencyProperty EffectiveItemsProperty = EffectiveItemsPropertyKey.DependencyProperty;

        public string? AutoSaveFilePath
        {
            get => (string?)GetValue(AutoSaveFilePathProperty);
            set => SetValue(AutoSaveFilePathProperty, value);
        }

        public static readonly DependencyProperty AutoSaveFilePathProperty =
            DependencyProperty.Register(nameof(AutoSaveFilePath), typeof(string), typeof(CodeReportTabView), new PropertyMetadata(null));

        public event EventHandler<object?>? CloseRequested;
        public event EventHandler<TabViewModel?>? SaveDataRequested;
        public event EventHandler? AddRequested;
        public event EventHandler<object?>? RenameRequested;

        internal void OnCloseRequested(object? item)
        {
            if (CloseCommand?.CanExecute(item) == true) { CloseCommand.Execute(item); return; }
            CloseRequested?.Invoke(this, item);
        }

        internal void OnSaveDataRequested(TabViewModel? tvm) => SafeAction(() => SaveDataRequested?.Invoke(this, tvm));

        private void InvokeSaveCommand(TabViewModel? tvm)
        {
            SafeAction(() =>
            {
                if (SaveDataCommand?.CanExecute(tvm) == true) SaveDataCommand.Execute(tvm);
            });

            OnSaveDataRequested(tvm);
        }

        internal void OnAddRequested()
        {
            if (AddCommand?.CanExecute(null) == true) { SafeAction(() => AddCommand.Execute(null)); return; }
            if (AddRequested != null) { SafeAction(() => AddRequested.Invoke(this, EventArgs.Empty)); return; }
            SafeAction(DefaultAdd);
        }

        internal void OnRenameRequested(object? item)
        {
            if (RenameCommand?.CanExecute(item) == true) { RenameCommand.Execute(item); return; }
            RenameRequested?.Invoke(this, item);
        }

        private void DefaultAdd()
        {
            const string baseName = "New Tab";
            string candidate;

            while (true)
            {
                candidate = _newTabCounter == 1 ? baseName : $"{baseName} {_newTabCounter}";
                if (!DoesHeaderExist(candidate)) { _newTabCounter++; break; }
                _newTabCounter++;
            }

            var vm = new TabViewModel(candidate, null);
            try { vm.InitializePdfFolder(_exeBaseDir); } catch { }

            if (ItemsSource is IList list && !list.IsReadOnly)
            {
                list.Add(vm);
                SelectedItem = vm;

                UpdateEffectiveItems();
                EnsureUniqueHeaders(EffectiveItems);
                UpdateTabSubscriptions();
                InvokeSaveCommand(null);

                SafeAction(() => CollectionViewSource.GetDefaultView(EffectiveItems)?.Refresh());
                return;
            }

            _fallbackItems.Add(vm);
            SelectedItem = vm;
        }

        private bool DoesHeaderExist(string header)
        {
            if (string.IsNullOrEmpty(header)) return false;
            var items = EffectiveItems;
            if (items == null) return false;

            foreach (var it in items)
            {
                if (it is TabViewModel tvm && string.Equals(tvm.Header, header, StringComparison.OrdinalIgnoreCase)) return true;
                if (it is TabModel tm && string.Equals(tm.Header, header, StringComparison.OrdinalIgnoreCase)) return true;
            }

            return false;
        }

        private bool IsHeaderDuplicate(string header, TabViewModel self)
        {
            if (string.IsNullOrEmpty(header)) return false;
            var items = EffectiveItems;
            if (items == null) return false;

            foreach (var it in items)
            {
                if (it is TabViewModel tvm && !ReferenceEquals(tvm, self) && string.Equals(tvm.Header, header, StringComparison.OrdinalIgnoreCase)) return true;
                if (it is TabModel tm && string.Equals(tm.Header, header, StringComparison.OrdinalIgnoreCase)) return true;
            }

            return false;
        }

        private string GenerateUniqueHeader(string baseName)
        {
            if (string.IsNullOrEmpty(baseName)) baseName = "Tab";
            if (!DoesHeaderExist(baseName)) return baseName;

            int i = 1;
            string candidate;
            do
            {
                candidate = $"{baseName} {i}";
                i++;
            } while (DoesHeaderExist(candidate));

            return candidate;
        }

        private string MakeUniqueAgainstSet(string baseName, HashSet<string> existing)
        {
            if (string.IsNullOrEmpty(baseName)) baseName = "Tab";
            if (!existing.Contains(baseName)) return baseName;

            int i = 1;
            string candidate;
            do
            {
                candidate = $"{baseName} {i}";
                i++;
            } while (existing.Contains(candidate));

            return candidate;
        }

        private void EnsureUniqueHeaders(IEnumerable? items)
        {
            if (items == null) return;

            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var eff = EffectiveItems;

            if (eff != null)
            {
                foreach (var cur in eff)
                {
                    if (cur is TabViewModel curTvm)
                    {
                        var inIncoming = false;
                        if (items is IEnumerable en)
                        {
                            foreach (var x in en)
                            {
                                if (ReferenceEquals(x, curTvm)) { inIncoming = true; break; }
                            }
                        }
                        if (!inIncoming)
                        {
                            var h = (curTvm.Header ?? string.Empty).Trim();
                            if (string.IsNullOrEmpty(h)) h = "Tab";
                            existing.Add(h);
                        }
                    }
                    else if (cur is TabModel curTm)
                    {
                        var h = (curTm.Header ?? string.Empty).Trim();
                        if (string.IsNullOrEmpty(h)) h = "Tab";
                        existing.Add(h);
                    }
                }
            }

            var seen = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);

            foreach (var it in items)
            {
                if (it is TabViewModel tvm)
                {
                    var header = (tvm.Header ?? string.Empty).Trim();
                    if (string.IsNullOrEmpty(header)) header = "Tab";

                    if (seen.Contains(header))
                    {
                        var unique = MakeUniqueAgainstSet(header, seen);
                        tvm.Header = unique;
                        seen.Add(unique);
                    }
                    else
                    {
                        tvm.Header = header;
                        seen.Add(header);
                    }
                }
                else if (it is TabModel tm)
                {
                    var header = (tm.Header ?? string.Empty).Trim();
                    if (string.IsNullOrEmpty(header)) header = "Tab";

                    if (seen.Contains(header))
                    {
                        var unique = MakeUniqueAgainstSet(header, seen);
                        tm.Header = unique;
                        seen.Add(unique);
                    }
                    else
                    {
                        tm.Header = header;
                        seen.Add(header);
                    }
                }
            }

            Dispatcher.BeginInvoke(() => SafeAction(() => CollectionViewSource.GetDefaultView(EffectiveItems)?.Refresh()), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void DefaultClose(object? item)
        {
            if (item == null) return;

            if (ItemsSource is IList list && !list.IsReadOnly)
            {
                if (list.Contains(item))
                {
                    list.Remove(item);
                    RunPostRemovalHousekeeping();
                    return;
                }

                for (int i = 0; i < list.Count; i++)
                {
                    var candidate = list[i];

                    if (ReferenceEquals(candidate, item))
                    {
                        list.RemoveAt(i);
                        RunPostRemovalHousekeeping();
                        return;
                    }

                    try
                    {
                        string? candHeader = candidate is TabViewModel cTvm ? cTvm.Header : candidate is TabModel cTm ? cTm.Header : null;
                        string? itemHeader = item is TabViewModel iTvm ? iTvm.Header : item is TabModel iTm ? iTm.Header : null;

                        if (!string.IsNullOrEmpty(candHeader) && !string.IsNullOrEmpty(itemHeader) &&
                            string.Equals(candHeader.Trim(), itemHeader.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            list.RemoveAt(i);
                            RunPostRemovalHousekeeping();
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"DefaultClose: comparison failed: {ex.Message}");
                    }
                }

                RunPostRemovalHousekeeping();
                return;
            }

            if (_fallbackItems.Contains(item))
            {
                _fallbackItems.Remove(item);
                RunPostRemovalHousekeeping();
                return;
            }
        }

        private void MenuItem_Rename_Click(object? sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem mi) return;

            var item = mi.Tag;
            if (item == null)
            {
                var cm = mi.Parent as ContextMenu ?? (mi.CommandParameter as ContextMenu);
                var placementTarget = cm?.PlacementTarget as FrameworkElement;
                item = placementTarget?.DataContext;
            }

            if (item == null) return;

            if (RenameCommand?.CanExecute(item) == true) { RenameCommand.Execute(item); return; }
            DefaultBeginRename(item);
        }

        private void DefaultBeginRename(object? item)
        {
            if (item == null) return;

            SelectedItem = item;

            if (item is TabViewModel tvm)
            {
                var tabItem = PART_TabControl?.ItemContainerGenerator.ContainerFromItem(item) as TabItem;
                if (tabItem != null && !_originalHeaders.ContainsKey(tabItem)) _originalHeaders[tabItem] = tvm.Header;

                tvm.IsEditing = true;

                Dispatcher.BeginInvoke(() =>
                {
                    var searchRoot = (DependencyObject?)tabItem ?? PART_TabControl as DependencyObject;
                    if (searchRoot == null) return;

                    var editBox = FindDescendantByName<TextBox>(searchRoot, "tbHeaderEdit");
                    if (editBox == null) return;

                    if (!editBox.Focusable) editBox.Focusable = true;
                    editBox.Focus();
                    editBox.SelectAll();
                }, System.Windows.Threading.DispatcherPriority.Input);

                return;
            }

            if (RenameCommand?.CanExecute(item) == true) { RenameCommand.Execute(item); return; }
            RenameRequested?.Invoke(this, item);
        }

        private void HeaderEdit_KeyDown(object? sender, KeyEventArgs e)
        {
            if (sender is not TextBox tb) return;
            if (tb.DataContext is not TabViewModel tvm) return;

            if (e.Key == Key.Enter)
            {
                tb.GetBindingExpression(TextBox.TextProperty)?.UpdateSource();
                var newHeader = tvm.Header?.Trim() ?? string.Empty;
                var tabItem = PART_TabControl?.ItemContainerGenerator.ContainerFromItem(tvm) as TabItem;

                if (string.IsNullOrEmpty(newHeader))
                {
                    if (tabItem != null && _originalHeaders.TryGetValue(tabItem, out var orig) && !string.IsNullOrEmpty(orig))
                    {
                        tvm.Header = orig;
                        _originalHeaders.Remove(tabItem);
                    }
                    else
                    {
                        tvm.Header = GenerateUniqueHeader("Tab");
                    }

                    tvm.IsEditing = false;
                    e.Handled = true;
                    return;
                }

                if (IsHeaderDuplicate(newHeader, tvm))
                {
                    var unique = GenerateUniqueHeader(newHeader);
                    if (tabItem != null) _originalHeaders.Remove(tabItem);

                    tvm.Header = unique;

                    WinUxMessageBox.Show(
                        $"A tab named \"{newHeader}\" already exists. The tab was renamed to \"{unique}\".",
                        "Duplicate Tab Name",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);

                    tvm.IsEditing = false;
                    EnsureUniqueHeaders(EffectiveItems);
                    Dispatcher.BeginInvoke(() => InvokeSaveCommand(null), System.Windows.Threading.DispatcherPriority.Background);

                    e.Handled = true;
                    return;
                }

                if (tabItem != null) _originalHeaders.Remove(tabItem);

                tvm.IsEditing = false;
                EnsureUniqueHeaders(EffectiveItems);
                Dispatcher.BeginInvoke(() => InvokeSaveCommand(null), System.Windows.Threading.DispatcherPriority.Background);

                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                if (tb.DataContext is not TabViewModel tvm2) return;
                var tabItem = PART_TabControl?.ItemContainerGenerator.ContainerFromItem(tvm2) as TabItem;
                if (tabItem != null && _originalHeaders.TryGetValue(tabItem, out var orig))
                {
                    tvm2.Header = orig ?? string.Empty;
                    _originalHeaders.Remove(tabItem);
                }

                tvm2.IsEditing = false;
                e.Handled = true;
            }
        }

        // Drag & drop handlers for reordering tabs
        private void TabItem_PreviewMouseLeftButtonDown(object? sender, MouseButtonEventArgs e)
        {
            var original = e.OriginalSource as DependencyObject;
            if (sender is not TabItem tabItem) { _draggedData = null; return; }

            if (tabItem.DataContext is TabViewModel tvm && tvm.IsEditing) { _draggedData = null; return; }
            if (!IsValidDragSource(original)) { _draggedData = null; return; }

            _dragStartPoint = e.GetPosition(null);
            _isDragging = false;
            _draggedData = tabItem.DataContext;
        }

        private void TabItem_MouseMove(object? sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed) return;
            if (_draggedData == null) return;

            var currentPos = e.GetPosition(null);
            var dx = Math.Abs(currentPos.X - _dragStartPoint.X);
            var dy = Math.Abs(currentPos.Y - _dragStartPoint.Y);

            if (_isDragging) return;

            if (dx > SystemParameters.MinimumHorizontalDragDistance || dy > SystemParameters.MinimumVerticalDragDistance)
            {
                _isDragging = true;
                var tabItem = sender as TabItem ?? (sender as FrameworkElement)?.TemplatedParent as TabItem;
                if (tabItem == null) { _isDragging = false; _draggedData = null; return; }

                var data = new DataObject("CodeReportTabView.Tab", _draggedData);
                try
                {
                    DragDrop.DoDragDrop((DependencyObject?)tabItem ?? this, data, DragDropEffects.Move);
                }
                finally
                {
                    _isDragging = false;
                    _draggedData = null;
                }
            }
        }

        private void TabItem_DragOver(object? sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent("CodeReportTabView.Tab") ? DragDropEffects.Move : DragDropEffects.None;
            e.Handled = true;
        }

        private void TabItem_Drop(object? sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent("CodeReportTabView.Tab")) return;

            var sourceItem = e.Data.GetData("CodeReportTabView.Tab");
            var targetItem = (sender as FrameworkElement)?.DataContext;
            if (sourceItem == null || targetItem == null) return;
            if (ReferenceEquals(sourceItem, targetItem)) return;

            IList? list = ItemsSource as IList;
            if (list == null || list.IsReadOnly) list = _fallbackItems;

            var srcIndex = list.IndexOf(sourceItem);
            var tgtIndex = list.IndexOf(targetItem);
            if (srcIndex < 0 || tgtIndex < 0 || srcIndex == tgtIndex) return;

            var item = list[srcIndex];
            list.RemoveAt(srcIndex);

            int insertIndex = srcIndex < tgtIndex ? Math.Max(0, tgtIndex - 1) : tgtIndex;
            insertIndex = Math.Min(insertIndex, list.Count);
            list.Insert(insertIndex, item);

            SelectedItem = item;

            UpdateEffectiveItems();
            EnsureUniqueHeaders(EffectiveItems);
            UpdateTabSubscriptions();
            InvokeSaveCommand(null);

            e.Handled = true;
        }

        private static bool IsValidDragSource(DependencyObject? original)
        {
            if (original == null) return false;

            var current = original;
            while (current != null)
            {
                if (current is Button || current is TextBox) return false;
                if (current is ContentPresenter || current is TabItem) return true;
                current = VisualTreeHelper.GetParent(current);
            }

            return false;
        }

        private void UpdateEffectiveItems()
        {
            if (ItemsSource == null || IsEnumerableEmpty(ItemsSource))
            {
                EffectiveItems = _fallbackItems;
                if (SelectedItem == null && _fallbackItems.Count > 0) SelectedItem = _fallbackItems[0];
            }
            else
            {
                EffectiveItems = ItemsSource!;
            }

            UpdateTabSubscriptions();
        }

        private static bool IsEnumerableEmpty(IEnumerable items)
        {
            if (items is ICollection col) return col.Count == 0;
            foreach (var _ in items) return false;
            return true;
        }

        private void SubscribeToNotifier(IEnumerable? items)
        {
            if (items is INotifyCollectionChanged incc)
            {
                _itemsSourceNotifier = incc;
                _itemsSourceNotifier.CollectionChanged += ItemsSource_CollectionChanged;
            }
        }

        private void UnsubscribeFromOldNotifier(IEnumerable? _)
        {
            if (_itemsSourceNotifier != null)
            {
                _itemsSourceNotifier.CollectionChanged -= ItemsSource_CollectionChanged;
                _itemsSourceNotifier = null;
            }
        }

        private void ItemsSource_CollectionChanged(object? _, NotifyCollectionChangedEventArgs __)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateEffectiveItems();
                EnsureUniqueHeaders(EffectiveItems);
                InvokeSaveCommand(null);
            });
        }

        private sealed class DelegateCommand : ICommand
        {
            private readonly Action<object?> _execute;
            private readonly Predicate<object?>? _canExecute;

            public DelegateCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
            {
                _execute = execute ?? throw new ArgumentNullException(nameof(execute));
                _canExecute = canExecute;
            }

            public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
            public void Execute(object? parameter) => _execute(parameter);
            public event EventHandler? CanExecuteChanged
            {
                add => CommandManager.RequerySuggested += value;
                remove => CommandManager.RequerySuggested -= value;
            }
        }

        private void CloseButton_Click(object? sender, RoutedEventArgs e)
        {
            if (sender is not FrameworkElement fe) return;
            var item = fe.DataContext;
            if (item == null) return;

            var result = WinUxMessageBox.Show(
                "Are you sure you want to close this tab?",
                "Confirm Close",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);

            if (result != MessageBoxResults.OK) return;

            if (CloseCommand?.CanExecute(item) == true) { CloseCommand.Execute(item); return; }
            if (CloseRequested != null) { CloseRequested.Invoke(this, item); return; }

            DefaultClose(item);
            InvokeSaveCommand(null);
        }

        private static T? FindDescendantByName<T>(DependencyObject root, string name) where T : FrameworkElement
        {
            if (root == null) return null;

            int count = VisualTreeHelper.GetChildrenCount(root);
            for (int i = 0; i < count; i++)
            {
                var child = VisualTreeHelper.GetChild(root, i);
                if (child is T fe && fe.Name == name) return fe;
                var found = FindDescendantByName<T>(child, name);
                if (found != null) return found;
            }

            return null;
        }

        private void UpdateTabSubscriptions()
        {
            // Unsubscribe previous handlers
            foreach (var kv in _tabSubscriptions)
            {
                var tvm = kv.Key;
                try
                {
                    tvm.Items.CollectionChanged -= kv.Value.collectionHandler;
                    tvm.PropertyChanged -= kv.Value.headerHandler;

                    var itemHandler = kv.Value.itemHandler;
                    foreach (var existingItem in tvm.Items.OfType<INotifyPropertyChanged>())
                    {
                        SafeAction(() => existingItem.PropertyChanged -= itemHandler);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UpdateTabSubscriptions: unsubscribe failed: {ex.Message}");
                }
            }
            _tabSubscriptions.Clear();

            var items = EffectiveItems;
            if (items == null) return;

            foreach (var it in items)
            {
                if (it is TabViewModel tvm)
                {
                    PropertyChangedEventHandler itemHandler = (_, __) => OnTabItemsChanged(tvm);

                    PropertyChangedEventHandler headerHandler = (_, pe) =>
                    {
                        if (pe?.PropertyName == nameof(TabViewModel.Header)) OnTabHeaderChanged(tvm);
                    };

                    NotifyCollectionChangedEventHandler collectionHandler = (_, args) =>
                    {
                        try
                        {
                            if (args.Action == NotifyCollectionChangedAction.Add && args.NewItems != null)
                            {
                                foreach (var newIt in args.NewItems.OfType<INotifyPropertyChanged>())
                                    SafeAction(() => newIt.PropertyChanged += itemHandler);
                            }
                            else if (args.Action == NotifyCollectionChangedAction.Remove && args.OldItems != null)
                            {
                                foreach (var oldIt in args.OldItems.OfType<INotifyPropertyChanged>())
                                    SafeAction(() => oldIt.PropertyChanged -= itemHandler);
                            }
                            else if (args.Action == NotifyCollectionChangedAction.Replace)
                            {
                                if (args.OldItems != null)
                                {
                                    foreach (var oldIt in args.OldItems.OfType<INotifyPropertyChanged>())
                                        SafeAction(() => oldIt.PropertyChanged -= itemHandler);
                                }
                                if (args.NewItems != null)
                                {
                                    foreach (var newIt in args.NewItems.OfType<INotifyPropertyChanged>())
                                        SafeAction(() => newIt.PropertyChanged += itemHandler);
                                }
                            }
                            else if (args.Action == NotifyCollectionChangedAction.Reset)
                            {
                                SafeAction(() =>
                                {
                                    foreach (var existing in tvm.Items.OfType<INotifyPropertyChanged>()) existing.PropertyChanged -= itemHandler;
                                });

                                foreach (var existing in tvm.Items.OfType<INotifyPropertyChanged>())
                                    SafeAction(() => existing.PropertyChanged += itemHandler);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"UpdateTabSubscriptions.collectionHandler: {ex.Message}");
                        }

                        OnTabItemsChanged(tvm);
                    };

                    tvm.Items.CollectionChanged += collectionHandler;
                    tvm.PropertyChanged += headerHandler;

                    foreach (var existingItem in tvm.Items.OfType<INotifyPropertyChanged>())
                    {
                        SafeAction(() => existingItem.PropertyChanged += itemHandler);
                    }

                    _tabSubscriptions[tvm] = (collectionHandler, headerHandler, itemHandler);
                }
            }
        }

        private void OnTabItemsChanged(TabViewModel tvm) => InvokeSaveCommand(tvm);

        private void OnTabHeaderChanged(TabViewModel tvm)
        {
            // when header changes, re-initialize per-tab Pdf folder using exe path,
            // so the folder stays in "exe path\Pdf Files\{TabName}"
            SafeAction(() => tvm.InitializePdfFolder(_exeBaseDir));
            InvokeSaveCommand(null);
        }

        public void SaveToFile(string filePath)
        {
            var tabs = CollectTabModels();
            TabPersistence.SaveTabsToFile(filePath, tabs);
        }

        /// <summary>
        /// Asynchronously load CRP (.crp) file and apply tabs on the UI thread.
        /// Also initializes per-tab PdfFolder and loads PdfFiles into each TabViewModel.
        /// </summary>
        public async Task LoadFromFileAsync(string filePath)
        {
            var tabs = await Task.Run(() => TabPersistence.LoadTabsFromFile(filePath)).ConfigureAwait(false);
            if (tabs == null) return;

            // Create TabViewModel instances and initialize PdfFolder on UI thread,
            // then load PdfFiles asynchronously after the UI work is done.
            var op = Dispatcher.InvokeAsync(() =>
            {
                var vmListLocal = tabs.Select(tm =>
                {
                    var tvm = new TabViewModel(tm.Header, null);
                    foreach (var ci in tm.Items ?? Enumerable.Empty<CodeReportTracker.Core.Models.CodeItem>()) tvm.Items.Add(ci);
                    try { tvm.InitializePdfFolder(_exeBaseDir); } catch { }
                    return (object)tvm;
                }).ToList();

                EnsureUniqueHeaders(vmListLocal);

                if (ItemsSource is IList list && !list.IsReadOnly)
                {
                    list.Clear();
                    foreach (var it in vmListLocal) list.Add(it);
                    SelectedItem = vmListLocal.FirstOrDefault();
                }
                else
                {
                    _fallbackItems.Clear();
                    foreach (var it in vmListLocal) _fallbackItems.Add(it);
                    SelectedItem = _fallbackItems.FirstOrDefault();
                    UpdateEffectiveItems();
                    UpdateTabSubscriptions();
                }

                return vmListLocal;
            });

            var vmList = await op.Task.ConfigureAwait(false);

            // Load pdf file lists for each TabViewModel (background I/O)
            try
            {
                var loadTasks = vmList.OfType<TabViewModel>().Select(tvm => tvm.LoadPdfFilesAsync(_exeBaseDir));
                await Task.WhenAll(loadTasks).ConfigureAwait(false);
            }
            catch
            {
                // ignore per-tab load failures
            }
        }

        /// <summary>
        /// Synchronous entry that starts async load without blocking caller.
        /// </summary>
        public void LoadFromFile(string filePath) => _ = LoadFromFileAsync(filePath);

        private List<TabModel> CollectTabModels()
        {
            var result = new List<TabModel>();
            var items = EffectiveItems;
            if (items == null) return result;

            CodeItem? MapRawToCodeItem(object? raw)
            {
                if (raw == null) return null;

                if (raw is CodeItemDto dto) return dto.ToCodeItem();

                if (raw is CodeItem ci)
                {
                    return new CodeItem
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
                        HasUpdate = ci.HasUpdate,
                        CodeExists = ci.CodeExists
                    };
                }

                try
                {
                    var rawType = raw.GetType();
                    var ci2 = new CodeItem
                    {
                        Number = rawType.GetProperty("Number")?.GetValue(raw) as string ?? string.Empty,
                        Link = rawType.GetProperty("Link")?.GetValue(raw) as string ?? string.Empty,
                        WebType = rawType.GetProperty("WebType")?.GetValue(raw) as string ?? string.Empty,
                        ProductCategory = rawType.GetProperty("ProductCategory")?.GetValue(raw) as string ?? string.Empty,
                        Description = rawType.GetProperty("Description")?.GetValue(raw) as string ?? string.Empty,
                        ProductsListed = rawType.GetProperty("ProductsListed")?.GetValue(raw) as string ?? string.Empty,
                        LatestCode = rawType.GetProperty("LatestCode")?.GetValue(raw) as string ?? string.Empty,
                        LatestCode_Old = rawType.GetProperty("LatestCode_Old")?.GetValue(raw) as string ?? string.Empty,
                        IssueDate = rawType.GetProperty("IssueDate")?.GetValue(raw) as string ?? string.Empty,
                        IssueDate_Old = rawType.GetProperty("IssueDate_Old")?.GetValue(raw) as string ?? string.Empty,
                        ExpirationDate = rawType.GetProperty("ExpirationDate")?.GetValue(raw) as string ?? string.Empty,
                        ExpirationDate_Old = rawType.GetProperty("ExpirationDate_Old")?.GetValue(raw) as string ?? string.Empty,
                        LastCheck = rawType.GetProperty("LastCheck")?.GetValue(raw) as string ?? string.Empty
                    };

                    try { ci2.DownloadProcess = Convert.ToInt32(rawType.GetProperty("DownloadProcess")?.GetValue(raw) ?? 0); } catch { }
                    try { ci2.HasCheck = Convert.ToBoolean(rawType.GetProperty("HasCheck")?.GetValue(raw) ?? false); } catch { }
                    try { ci2.HasUpdate = Convert.ToBoolean(rawType.GetProperty("HasUpdate")?.GetValue(raw) ?? false); } catch { }
                    try { ci2.CodeExists = Convert.ToBoolean(rawType.GetProperty("CodeExists")?.GetValue(raw) ?? false); } catch { }

                    if (!string.IsNullOrEmpty(ci2.Number) ||
                        !string.IsNullOrEmpty(ci2.Description) ||
                        !string.IsNullOrEmpty(ci2.Link) ||
                        ci2.DownloadProcess != 0 ||
                        ci2.HasCheck || ci2.HasUpdate || ci2.CodeExists)
                    {
                        return ci2;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"MapRawToCodeItem reflection failed: {ex.Message}");
                }

                return null;
            }

            foreach (var it in items)
            {
                if (it is TabViewModel tvm)
                {
                    var codeList = new List<CodeItem>();

                    SafeAction(() =>
                    {
                        if (tvm.Content is CodeReportTable hostedTable)
                        {
                            foreach (var raw in hostedTable.Items)
                            {
                                var mapped = MapRawToCodeItem(raw);
                                if (mapped != null) codeList.Add(mapped);
                            }
                        }
                    });

                    if (codeList.Count == 0)
                    {
                        foreach (var ci in tvm.Items)
                        {
                            var mapped = MapRawToCodeItem(ci);
                            if (mapped != null) codeList.Add(mapped);
                        }
                    }

                    result.Add(new TabModel { Header = tvm.Header, Items = codeList });
                }
                else if (it is TabModel tm)
                {
                    var cloned = (tm.Items ?? new List<CodeItem>()).Select(i => MapRawToCodeItem(i) ?? new CodeItem()).ToList();
                    result.Add(new TabModel { Header = tm.Header, Items = cloned });
                }
                else
                {
                    SafeAction(() =>
                    {
                        var type = it.GetType();
                        var headerProp = type.GetProperty("Header");
                        var itemsProp = type.GetProperty("Items");
                        if (headerProp != null)
                        {
                            var headerVal = headerProp.GetValue(it) as string ?? string.Empty;
                            var codeList = new List<CodeItem>();

                            if (itemsProp != null)
                            {
                                var rawItems = itemsProp.GetValue(it) as IEnumerable;
                                if (rawItems != null)
                                {
                                    foreach (var raw in rawItems)
                                    {
                                        var mapped = MapRawToCodeItem(raw);
                                        if (mapped != null) codeList.Add(mapped);
                                    }
                                }
                            }

                            result.Add(new TabModel { Header = headerVal, Items = codeList });
                        }
                    });
                }
            }

            return result;
        }

        #region Helper utilities

        private static void Bind(DependencyObject target, DependencyProperty property, string path, object source, BindingMode mode)
        {
            var binding = new Binding(path) { Source = source, Mode = mode };
            BindingOperations.SetBinding(target, property, binding);
        }

        private static void SafeAction(Action? action)
        {
            if (action == null) return;
            try { action(); } catch (Exception ex) { Debug.WriteLine($"SafeAction exception: {ex.Message}"); }
        }

        private static Image? LoadIcon(string uriString)
        {
            try
            {
                var uri = new Uri(uriString, UriKind.Absolute);
                var bmp = new BitmapImage(uri);
                return new Image { Source = bmp, Width = 16, Height = 16, VerticalAlignment = VerticalAlignment.Center };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"LoadIcon failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Centralized housekeeping to run after a tab is removed.
        /// Ensures EffectiveItems, subscriptions, SelectedItem and save-command invocation are updated.
        /// </summary>
        private void RunPostRemovalHousekeeping()
        {
            UpdateEffectiveItems();
            EnsureUniqueHeaders(EffectiveItems);
            UpdateTabSubscriptions();
            InvokeSaveCommand(null);

            SafeAction(() => CollectionViewSource.GetDefaultView(EffectiveItems)?.Refresh());

            if (SelectedItem != null)
            {
                var found = false;
                foreach (var it in EffectiveItems)
                {
                    if (ReferenceEquals(it, SelectedItem) || Equals(it, SelectedItem)) { found = true; break; }
                }

                if (!found)
                {
                    SelectedItem = (EffectiveItems as IEnumerable)?.OfType<object>().FirstOrDefault();
                    SafeAction(() => { if (PART_TabControl != null) PART_TabControl.SelectedItem = SelectedItem; });
                }
            }
        }

        #endregion

        private void MenuItem_OpenPdfFolder_Click(object? sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem mi) return;

            var item = mi.Tag;
            if (item == null)
            {
                var cm = mi.Parent as ContextMenu ?? (mi.CommandParameter as ContextMenu);
                var placementTarget = cm?.PlacementTarget as FrameworkElement;
                item = placementTarget?.DataContext;
            }

            if (item == null) return;

            // Prefer new TabViewModel.PdfFolder property if available, fall back to PdfFolderPath for older compatibility
            string? folderPath = null;
            try
            {
                var prop = item.GetType().GetProperty("PdfFolder");
                if (prop != null) folderPath = prop.GetValue(item) as string;
            }
            catch { /* ignore reflection errors */ }

            if (string.IsNullOrEmpty(folderPath))
            {
                try
                {
                    var prop = item.GetType().GetProperty("PdfFolderPath");
                    if (prop != null) folderPath = prop.GetValue(item) as string;
                }
                catch { /* ignore reflection errors */ }
            }

            if (string.IsNullOrEmpty(folderPath))
            {
                // Use exe base directory + "Pdf Files" + sanitized tab name
                var header = item is TabViewModel tvm ? (tvm.Header ?? "Tab") : (item.GetType().GetProperty("Header")?.GetValue(item) as string ?? "Tab");
                foreach (var c in Path.GetInvalidFileNameChars()) header = header.Replace(c, '_');
                folderPath = Path.Combine(_exeBaseDir, "Pdf Files", header);
            }

            if (!Directory.Exists(folderPath))
            {
                var create = WinUxMessageBox.Show(
                    $"Folder does not exist:\n{folderPath}\n\nCreate it now?",
                    "Open Pdf Folder",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question);

                if (create == MessageBoxResults.OK)
                {
                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch (Exception ex)
                    {
                        WinUxMessageBox.Show($"Unable to create folder:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    return;
                }
            }

            try
            {
                Process.Start(new ProcessStartInfo { FileName = folderPath, UseShellExecute = true });
            }
            catch (Exception ex)
            {
                WinUxMessageBox.Show($"Failed to open folder:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}