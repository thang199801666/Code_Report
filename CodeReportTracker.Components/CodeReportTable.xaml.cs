using CodeReportTracker.Components.Helpers;
using CodeReportTracker.Components.ViewModels;
using CodeReportTracker.Core.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using WinUx;

namespace CodeReportTracker.Components
{
    public partial class CodeReportTable : UserControl
    {
        public CodeReportTable()
        {
            InitializeComponent();

            dgMain.CanUserAddRows = true;
            dgMain.SelectionChanged += DgMain_SelectionChanged;

            if (dgMain.ContextMenu != null)
                dgMain.ContextMenu.Opened += DgMain_ContextMenuOpened;

            if (DesignerProperties.GetIsInDesignMode(this) && DataContext == null)
                DataContext = new CodeReportTableViewModel();

            // Ensure grid has an add-capable collection so the new-item placeholder is shown.
            if (ItemsSource == null)
            {
                var initial = new ObservableCollection<CodeItem>();
                ItemsSource = initial;
                dgMain.ItemsSource = initial;
            }
        }

        // Expose DataGrid Items so callers can iterate
        public ItemCollection Items => dgMain.Items;

        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(
                nameof(ItemsSource),
                typeof(IEnumerable),
                typeof(CodeReportTable),
                new PropertyMetadata(null, OnItemsSourceChanged));

        public IEnumerable ItemsSource
        {
            get => (IEnumerable)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }

        private static void OnItemsSourceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is CodeReportTable ctrl)
                ctrl.ApplyItemsSource(e.NewValue as IEnumerable);
        }

        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register(
                nameof(SelectedItem),
                typeof(object),
                typeof(CodeReportTable),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnSelectedItemChanged));

        public object SelectedItem
        {
            get => GetValue(SelectedItemProperty);
            set => SetValue(SelectedItemProperty, value);
        }

        private static void OnSelectedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is CodeReportTable ctrl && ctrl.dgMain.SelectedItem != e.NewValue)
                ctrl.dgMain.SelectedItem = e.NewValue;
        }

        private void DgMain_SelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (SelectedItem != dgMain.SelectedItem)
                SelectedItem = dgMain.SelectedItem;
        }

        // Ensure the DataGrid gets an add-capable collection.
        private void ApplyItemsSource(IEnumerable? source)
        {
            if (source == null)
            {
                var oc = new ObservableCollection<CodeItem>();
                ItemsSource = oc;
                dgMain.ItemsSource = oc;
                return;
            }

            // If source is IList and writable, use it directly so additions affect original collection.
            if (source is IList list)
            {
                try
                {
                    if (!list.IsReadOnly)
                    {
                        dgMain.ItemsSource = list;
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"ApplyItemsSource: IList interrogation failed: {ex.Message}");
                    // fall through to wrapping
                }
            }

            // If already an ObservableCollection<CodeItem>, use it directly.
            if (source is ObservableCollection<CodeItem> occ)
            {
                dgMain.ItemsSource = occ;
                return;
            }

            // Wrap any CodeItem instances into an ObservableCollection<CodeItem>.
            var wrapped = new ObservableCollection<CodeItem>();
            foreach (var o in source)
            {
                if (o is CodeItem ci) wrapped.Add(ci);
            }

            dgMain.ItemsSource = wrapped;
        }

        // Public API used by MainWindow
        public void SetData(IEnumerable<CodeItem> data, bool append)
        {
            if (data == null) return;

            var itemsSource = dgMain.ItemsSource as IList;

            if (!append)
            {
                if (itemsSource != null)
                {
                    try { itemsSource.Clear(); } catch (Exception ex) { Debug.WriteLine($"SetData.Clear itemsSource: {ex.Message}"); }
                }
                else
                {
                    dgMain.Items.Clear();
                }
            }

            foreach (var item in data)
            {
                if (itemsSource != null)
                {
                    try { itemsSource.Add(item); }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"SetData.Add to itemsSource failed: {ex.Message}");
                        try { dgMain.Items.Add(item); } catch (Exception ex2) { Debug.WriteLine($"SetData.Add fallback failed: {ex2.Message}"); }
                    }
                }
                else
                {
                    try { dgMain.Items.Add(item); } catch (Exception ex) { Debug.WriteLine($"SetData.Add direct failed: {ex.Message}"); }
                }
            }

            try { dgMain.Items.Refresh(); } catch { }
        }

        public void ClearData()
        {
            var itemsSource = dgMain.ItemsSource as IList;
            if (itemsSource != null)
            {
                try { itemsSource.Clear(); } catch (Exception ex) { Debug.WriteLine($"ClearData: {ex.Message}"); }
            }
            else
            {
                dgMain.Items.Clear();
            }

            try { dgMain.Items.Refresh(); } catch { }
        }

        public void SetSource(IEnumerable<CodeItem> data)
        {
            if (data == null) return;

            var itemsSource = dgMain.ItemsSource as IList;
            if (itemsSource != null)
            {
                try
                {
                    itemsSource.Clear();
                    foreach (var it in data) itemsSource.Add(it);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"SetSource fallback: {ex.Message}");
                    dgMain.Items.Clear();
                    foreach (var it in data) dgMain.Items.Add(it);
                }
            }
            else
            {
                dgMain.Items.Clear();
                foreach (var it in data) dgMain.Items.Add(it);
            }

            try { dgMain.Items.Refresh(); } catch { }
        }

        public void Refresh()
        {
            try { dgMain.Items.Refresh(); } catch { }
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (ItemsSource != null && dgMain.ItemsSource == null)
                ApplyItemsSource(ItemsSource);
        }

        private void DataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void RowDoubleClick(object sender, RoutedEventArgs e)
        {
            dgMain.CommitEdit();

            if (sender is DataGridRow row && row.Item is CodeItem codeItem)
            {
                // ensure the row is selected
                try { dgMain.SelectedItem = codeItem; } catch { }

                try
                {
                    var dlg = new CodeItemEditDialog(codeItem)
                    {
                        Owner = Window.GetWindow(this),
                        WindowStartupLocation = WindowStartupLocation.CenterOwner
                    };

                    if (dlg.ShowDialog() == true)
                    {
                        try { dgMain.Items.Refresh(); } catch { }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"RowDoubleClick: failed to open editor: {ex.Message}");
                    WinUx.Controls.WinUxMessageBox.Show($"Edit requested for {codeItem.Number}", 
                                                        "Edit",
                                                        WinUx.Controls.MessageBoxButtons.OK,
                                                        WinUx.Controls.MessageBoxIcon.Information);
                }
            }

            dgMain.CancelEdit();
        }

        private void DataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = true;
                PasteFromClipboard();
                return;
            }
            if (e.Key == Key.F2)
            {
                e.Handled = true;
                return;
            }
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                return;
            }
        }

        private void PasteMenu_Click(object sender, RoutedEventArgs e) => PasteFromClipboard();

        private void PasteCommand_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            try
            {
                var dobj = Clipboard.GetDataObject();
                e.CanExecute = dobj != null && (dobj.GetDataPresent(DataFormats.UnicodeText) || dobj.GetDataPresent(DataFormats.Text) || dobj.GetDataPresent(DataFormats.Html));
            }
            catch
            {
                e.CanExecute = false;
            }

            e.Handled = true;
        }

        private void PasteCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            PasteFromClipboard();
            e.Handled = true;
        }

        private void PasteFromClipboard()
        {
            IDataObject dobj;
            try { dobj = Clipboard.GetDataObject(); }
            catch { return; }

            if (dobj == null) return;

            var columnsFromGrid = BuildColumnsFromGrid();

            var parsed = ClipboardHelpers.ParseClipboardToCodeItemsWithSetProps(dobj, columnsFromGrid);
            if (parsed == null || parsed.Count == 0) return;

            var controlItemsSource = ItemsSource as IList;
            var dgItemsSource = dgMain.ItemsSource as IList;

            if (controlItemsSource == null && dgItemsSource == null)
            {
                var oc = new ObservableCollection<CodeItem>();
                foreach (var it in dgMain.Items)
                    if (it is CodeItem ci) oc.Add(ci);

                ItemsSource = oc;
                dgMain.ItemsSource = oc;
                controlItemsSource = oc;
                dgItemsSource = oc;
            }

            IEnumerable<object> searchCollection = controlItemsSource != null
                ? controlItemsSource.Cast<object>()
                : dgItemsSource != null
                    ? dgItemsSource.Cast<object>()
                    : dgMain.Items.Cast<object>();

            bool newItemSelected = dgMain.SelectedItem == CollectionView.NewItemPlaceholder
                                   || (dgMain.SelectedIndex == dgMain.Items.Count - 1 && dgMain.SelectedIndex >= 0);

            if (parsed.Count == 1 && newItemSelected)
            {
                var (singleParsedItem, _) = parsed[0];
                TryAddItem(singleParsedItem, controlItemsSource, dgItemsSource);
                dgMain.SelectedItem = singleParsedItem;
                dgMain.ScrollIntoView(singleParsedItem);

                if (dgMain.Columns.Count > 0)
                {
                    dgMain.CurrentCell = new DataGridCellInfo(singleParsedItem, dgMain.Columns[0]);
                    try { dgMain.BeginEdit(); } catch { }
                }

                try { dgMain.Items.Refresh(); } catch { }
                return;
            }

            foreach (var (parsedItem, setProps) in parsed)
            {
                var found = FindExisting(parsedItem, searchCollection);
                if (found != null)
                {
                    UpdateProperties(found, parsedItem, setProps);
                }
                else
                {
                    TryAddItem(parsedItem, controlItemsSource, dgItemsSource);
                }
            }

            try { dgMain.CommitEdit(DataGridEditingUnit.Row, true); } catch { }
            try { dgMain.Items.Refresh(); } catch { }
        }

        private List<string> BuildColumnsFromGrid()
        {
            var list = new List<string>();
            foreach (var col in dgMain.Columns)
            {
                if (col is DataGridBoundColumn boundCol)
                {
                    var bind = boundCol.Binding as Binding;
                    list.Add(bind?.Path?.Path ?? string.Empty);
                }
                else
                {
                    list.Add(col.Header?.ToString() ?? string.Empty);
                }
            }

            return list;
        }

        private static object FindExisting(CodeItem parsedItem, IEnumerable<object> searchCollection)
        {
            if (!string.IsNullOrWhiteSpace(parsedItem.Number))
            {
                var key = parsedItem.Number.Trim();
                var found = searchCollection.FirstOrDefault(o =>
                    o is CodeItem ci && !string.IsNullOrWhiteSpace(ci.Number) && string.Equals(ci.Number.Trim(), key, StringComparison.OrdinalIgnoreCase));
                if (found != null) return found;
            }

            if (!string.IsNullOrWhiteSpace(parsedItem.Link))
            {
                var key = parsedItem.Link.Trim();
                return searchCollection.FirstOrDefault(o =>
                    o is CodeItem ci && !string.IsNullOrWhiteSpace(ci.Link) && string.Equals(ci.Link.Trim(), key, StringComparison.OrdinalIgnoreCase));
            }

            return null;
        }

        private static void UpdateProperties(object target, CodeItem source, IEnumerable<string> propertyNames)
        {
            var targetType = target.GetType();
            var sourceType = typeof(CodeItem);

            foreach (var propName in propertyNames)
            {
                var targetProp = targetType.GetProperty(propName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                if (targetProp == null || !targetProp.CanWrite) continue;

                var sourceProp = sourceType.GetProperty(propName, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.IgnoreCase);
                if (sourceProp == null) continue;

                var value = sourceProp.GetValue(source);
                try { targetProp.SetValue(target, value); }
                catch (Exception ex) { Debug.WriteLine($"UpdateProperties.SetValue failed for {propName}: {ex.Message}"); }
            }
        }

        private void TryAddItem(CodeItem item, IList controlItemsSource, IList dgItemsSource)
        {
            try
            {
                if (controlItemsSource != null && !controlItemsSource.IsReadOnly)
                {
                    controlItemsSource.Add(item);
                    return;
                }

                if (dgItemsSource != null && !dgItemsSource.IsReadOnly)
                {
                    dgItemsSource.Add(item);
                    return;
                }

                dgMain.Items.Add(item);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TryAddItem failed: {ex.Message}");
                try { dgMain.Items.Add(item); } catch (Exception ex2) { Debug.WriteLine($"TryAddItem fallback failed: {ex2.Message}"); }
            }
        }

        private static bool ParseBoolLike(string? text)
        {
            var lower = text?.Trim().ToLowerInvariant() ?? string.Empty;
            if (string.IsNullOrEmpty(lower)) return false;
            if (lower is "1" or "true" or "yes" or "y" or "t") return true;
            if (lower is "0" or "false" or "no" or "n" or "f") return false;
            return bool.TryParse(lower, out var b) && b;
        }

        private static object? ConvertStringToType(string? text, Type targetType)
        {
            if (string.IsNullOrEmpty(text))
            {
                if (Nullable.GetUnderlyingType(targetType) != null || !targetType.IsValueType) return null;
                return Activator.CreateInstance(targetType);
            }

            var nonNullableType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (nonNullableType == typeof(string)) return text;
                if (nonNullableType == typeof(int))
                {
                    if (int.TryParse(text, out var i)) return i;
                    if (text.EndsWith("%") && int.TryParse(text.TrimEnd('%').Trim(), out i)) return i;
                    if (double.TryParse(text, out var d)) return (int)d;
                    return 0;
                }

                if (nonNullableType == typeof(long))
                {
                    if (long.TryParse(text, out var l)) return l;
                    return 0L;
                }

                if (nonNullableType == typeof(bool))
                    return ParseBoolLike(text);

                if (nonNullableType == typeof(DateTime))
                {
                    if (DateTime.TryParse(text, out var dt)) return dt;
                    return DateTime.MinValue;
                }

                if (nonNullableType.IsEnum)
                {
                    try { return Enum.Parse(nonNullableType, text, true); }
                    catch { return Activator.CreateInstance(nonNullableType); }
                }

                return Convert.ChangeType(text, nonNullableType, CultureInfo.InvariantCulture);
            }
            catch
            {
                return nonNullableType.IsValueType ? Activator.CreateInstance(nonNullableType) : null;
            }
        }

        private void DgMain_ContextMenuOpened(object? sender, RoutedEventArgs e)
        {
            if (dgMain?.ContextMenu == null) return;

            MenuItem? columnsMenu = null;
            foreach (var it in dgMain.ContextMenu.Items)
            {
                if (it is MenuItem mi && (mi.Name == "miColumns" || (mi.Header is string s && s == "Columns")))
                {
                    columnsMenu = mi;
                    break;
                }
            }

            if (columnsMenu == null) return;

            columnsMenu.Items.Clear();

            foreach (var col in dgMain.Columns)
            {
                string headerText = col.Header?.ToString() ?? $"Column {dgMain.Columns.IndexOf(col) + 1}";
                var entry = new MenuItem
                {
                    Header = headerText,
                    IsCheckable = true,
                    IsChecked = col.Visibility == Visibility.Visible,
                    Tag = col
                };

                entry.Click += (s, _) =>
                {
                    if (entry.Tag is DataGridColumn c)
                        c.Visibility = entry.IsChecked ? Visibility.Visible : Visibility.Collapsed;
                };

                columnsMenu.Items.Add(entry);
            }

            columnsMenu.Items.Add(new Separator());

            var showAll = new MenuItem { Header = "Show All" };
            showAll.Click += (_, _) => { foreach (var c in dgMain.Columns) c.Visibility = Visibility.Visible; };
            columnsMenu.Items.Add(showAll);

            var hideAll = new MenuItem { Header = "Hide All" };
            hideAll.Click += (_, _) => { foreach (var c in dgMain.Columns) c.Visibility = Visibility.Collapsed; };
            columnsMenu.Items.Add(hideAll);
        }

        private void ColumnHeader_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is DataGridColumnHeader header)) return;

            var parentGrid = FindParent<DataGrid>(header) ?? dgMain;
            if (parentGrid == null) return;

            var cm = new ContextMenu();

            foreach (var col in parentGrid.Columns)
            {
                string colHeaderText = col.Header?.ToString() ?? $"Column {parentGrid.Columns.IndexOf(col) + 1}";

                var mi = new MenuItem
                {
                    Header = colHeaderText,
                    IsCheckable = true,
                    IsChecked = col.Visibility == Visibility.Visible,
                    Tag = col
                };

                mi.Click += (s, _) =>
                {
                    if (mi.Tag is DataGridColumn c)
                        c.Visibility = mi.IsChecked ? Visibility.Visible : Visibility.Collapsed;
                };

                cm.Items.Add(mi);
            }

            cm.Items.Add(new Separator());

            var showAll = new MenuItem { Header = "Show All" };
            showAll.Click += (_, _) => { foreach (var c in parentGrid.Columns) c.Visibility = Visibility.Visible; };
            cm.Items.Add(showAll);

            var hideAll = new MenuItem { Header = "Hide All" };
            hideAll.Click += (_, _) => { foreach (var c in parentGrid.Columns) c.Visibility = Visibility.Collapsed; };
            cm.Items.Add(hideAll);

            header.ContextMenu = cm;
            cm.IsOpen = true;
            e.Handled = true;
        }

        private static T? FindParent<T>(DependencyObject? child) where T : DependencyObject
        {
            if (child == null) return null;

            DependencyObject? parent = null;

            try
            {
                parent = VisualTreeHelper.GetParent(child);
            }
            catch (ArgumentException)
            {
                parent = LogicalTreeHelper.GetParent(child);
            }

            while (parent != null && parent is not T)
            {
                DependencyObject? next = null;
                try
                {
                    next = VisualTreeHelper.GetParent(parent);
                }
                catch (ArgumentException)
                {
                    next = LogicalTreeHelper.GetParent(parent);
                }

                if (next == null && parent is FrameworkElement fe)
                    next = fe.Parent;

                parent = next;
            }

            return parent as T;
        }
    }
}