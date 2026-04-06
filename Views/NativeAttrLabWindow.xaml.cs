using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using NwApp = Autodesk.Navisworks.Api.Application;
using WpfColor = System.Windows.Media.Color;

namespace SetAtributesToolkit
{
    public partial class NativeAttrLabWindow : Window
    {
        // ── Constantes ──────────────────────────────────────────────────────────
        private const string DefaultStorageCategory = "Toolkit_Attributes";
        private const string AwpFieldName           = "Set_Name";

        // ── Cores por modo ──────────────────────────────────────────────────────
        private static readonly WpfColor ColorWrite       = (WpfColor)ColorConverter.ConvertFromString("#E8F5E9");
        private static readonly WpfColor ColorWriteBorder = (WpfColor)ColorConverter.ConvertFromString("#A5D6A7");
        private static readonly WpfColor ColorEdit        = (WpfColor)ColorConverter.ConvertFromString("#FFF8E1");
        private static readonly WpfColor ColorEditBorder  = (WpfColor)ColorConverter.ConvertFromString("#FFE082");
        private static readonly WpfColor ColorDelete      = (WpfColor)ColorConverter.ConvertFromString("#FFEBEE");
        private static readonly WpfColor ColorDeleteBorder= (WpfColor)ColorConverter.ConvertFromString("#EF9A9A");

        // ── Estado: Sets ────────────────────────────────────────────────────────
        private readonly ObservableCollection<SetEntry> _allSets      = new ObservableCollection<SetEntry>();
        private readonly ObservableCollection<SetEntry> _filteredSets = new ObservableCollection<SetEntry>();

        // ── Estado: Grade de atributos ──────────────────────────────────────────
        private readonly ObservableCollection<AttributeEntry> _entries = new ObservableCollection<AttributeEntry>();

        // ── Estado: Árvore nativa interativa ────────────────────────────────────
        private readonly ObservableCollection<CatNode> _nativeCats = new ObservableCollection<CatNode>();
        private Dictionary<string, Dictionary<string, string>> _nativeCatMap
            = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        // ────────────────────────────────────────────────────────────────────────

        public NativeAttrLabWindow()
        {
            InitializeComponent();

            StorageCategoryBox.Text    = DefaultStorageCategory;
            SetsListView.ItemsSource   = _filteredSets;
            AttributesGrid.ItemsSource = _entries;
            NativeTree.ItemsSource     = _nativeCats;

            Loaded += async (s, e) =>
            {
                ApplyModeUI();
                await LoadSetsAsync();
                await ReloadNativeAsync();
                Activate();
                Topmost = true;
                Topmost = false;
            };
        }

        // ══ MODO ════════════════════════════════════════════════════════════════

        private void RadioModo_Changed(object sender, RoutedEventArgs e)
        {
            if (CustomCategoryBox == null) return;
            ApplyModeUI();
        }

        private void ApplyModeUI()
        {
            bool isWrite  = RadioGravar?.IsChecked == true;
            bool isEdit   = RadioEditar?.IsChecked == true;
            bool isDelete = RadioExcluir?.IsChecked == true;

            // Header de modo
            if (isWrite)
            {
                ModeHeaderBorder.Background  = new SolidColorBrush(ColorWrite);
                ModeHeaderBorder.BorderBrush = new SolidColorBrush(ColorWriteBorder);
                ModeHeaderBorder.BorderThickness = new Thickness(1);
                ModeHeaderIcon.Text       = "●";
                ModeHeaderIcon.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E7D32"));
                ModeHeaderText.Text       = "Mode: Write — fill attributes and click Write & Verify";
                ModeHeaderText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#1B5E20"));
                SectionTitle.Text = "Attributes to Write";
            }
            else if (isEdit)
            {
                ModeHeaderBorder.Background  = new SolidColorBrush(ColorEdit);
                ModeHeaderBorder.BorderBrush = new SolidColorBrush(ColorEditBorder);
                ModeHeaderBorder.BorderThickness = new Thickness(1);
                ModeHeaderIcon.Text       = "●";
                ModeHeaderIcon.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F57F17"));
                ModeHeaderText.Text       = "Mode: Edit — check props in the tree, modify values, then write";
                ModeHeaderText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#E65100"));
                SectionTitle.Text = "Attributes to Edit";
            }
            else // Delete
            {
                ModeHeaderBorder.Background  = new SolidColorBrush(ColorDelete);
                ModeHeaderBorder.BorderBrush = new SolidColorBrush(ColorDeleteBorder);
                ModeHeaderBorder.BorderThickness = new Thickness(1);
                ModeHeaderIcon.Text       = "●";
                ModeHeaderIcon.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#C62828"));
                ModeHeaderText.Text       = "Mode: Delete — the category below will be REMOVED from selected elements";
                ModeHeaderText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#B71C1C"));
                SectionTitle.Text = "Category to Delete";
            }

            // Bloqueio de edição no modo Delete
            bool editable = isWrite || isEdit;
            if (CustomCategoryBox != null)
            {
                if (!isDelete) SyncCategoryNameBox(); // gerencia IsEnabled dinamicamente
                else
                {
                    CustomCategoryBox.IsEnabled = false;
                    CustomCategoryBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F0F0F0"));
                }
            }
            if (AttributesGrid != null) AttributesGrid.IsReadOnly = !editable;
            if (BtnAddRow     != null) BtnAddRow.IsEnabled   = editable;
            if (BtnRemoveRow  != null) BtnRemoveRow.IsEnabled = editable;
            if (BtnNewCategory!= null) BtnNewCategory.IsEnabled = editable;

            // Botão Write & Verify
            if (BtnWriteAndVerify != null)
            {
                if (isDelete)
                {
                    BtnWriteAndVerify.Content    = "Delete from Selection";
                    BtnWriteAndVerify.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FFEBEE"));
                    BtnWriteAndVerify.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#C62828"));
                }
                else
                {
                    BtnWriteAndVerify.Content    = isEdit ? "Edit & Verify (selection)" : "Write & Verify (selection)";
                    BtnWriteAndVerify.Background = new SolidColorBrush(ColorWrite);
                    BtnWriteAndVerify.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E7D32"));
                }
            }

            UpdateCategoryInputLabel();
            UpdateStatus();
        }

        // ══ STATUS DA SELEÇÃO ATUAL ══════════════════════════════════════════════

        private void UpdateStatus()
        {
            if (StatusText == null) return;
            try
            {
                int count = NwApp.ActiveDocument?.CurrentSelection.SelectedItems.Count ?? 0;
                StatusText.Text = $"Elements affected by Write/Delete: {count}";
                ZeroSelectionWarning.Visibility = count > 0 ? Visibility.Collapsed : Visibility.Visible;
            }
            catch
            {
                StatusText.Text = "Could not read selection.";
                ZeroSelectionWarning.Visibility = Visibility.Visible;
            }
        }

        // ══ SETS DE SELEÇÃO ═════════════════════════════════════════════════════

        private async Task LoadSetsAsync()
        {
            _allSets.Clear();
            _filteredSets.Clear();
            FooterStatus.Text = "Carregando sets...";

            var doc = NwApp.ActiveDocument;
            if (doc == null)
            {
                FooterStatus.Text = "Nenhum documento ativo.";
                HeaderSubtitle.Text = "Nenhum documento ativo.";
                return;
            }

            var sets = await Task.Run(() => CollectSets(doc));

            foreach (var entry in sets)
            {
                entry.PropertyChanged += OnSetCheckedChanged;
                _allSets.Add(entry);
            }

            ApplyFilter(SetsSearchBox.Text);
            UpdateSetsUI();
        }

        private static List<SetEntry> CollectSets(Document doc)
        {
            var result = new List<SetEntry>();
            try
            {
                foreach (SavedItem item in doc.SelectionSets.Value)
                    CollectRecursive(item, doc, result);
            }
            catch { }
            return result;
        }

        private static void CollectRecursive(SavedItem item, Document doc, List<SetEntry> result)
        {
            if (item is SelectionSet ss)
            {
                int count = 0;
                try { count = ss.GetSelectedItems(doc).Count; } catch { }
                result.Add(new SetEntry { Name = ss.DisplayName, ItemCount = count, Source = ss });
            }
            else if (item is Autodesk.Navisworks.Api.GroupItem gi)
            {
                foreach (SavedItem child in gi.Children)
                    CollectRecursive(child, doc, result);
            }
        }

        private void OnSetCheckedChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SetEntry.IsChecked))
                UpdateSetsUI();
        }

        private void UpdateSetsUI()
        {
            int total    = _allSets.Count;
            int selected = _allSets.Count(s => s.IsChecked);
            string first = _allSets.FirstOrDefault(s => s.IsChecked)?.Name ?? "—";

            SelectedCountText.Text = $"{selected} selected";
            HeaderSubtitle.Text    = total == 0
                ? "Nenhum set de seleção encontrado"
                : $"{selected}/{total} set(s) selecionado(s) — First: {first}";

            SetsCountText.Text = selected > 0
                ? $"{selected} of {total} set name(s) selected for the {AwpFieldName} field."
                : string.Empty;
        }

        private void ApplyFilter(string query)
        {
            _filteredSets.Clear();
            foreach (var s in _allSets)
                if (string.IsNullOrWhiteSpace(query) ||
                    s.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                    _filteredSets.Add(s);
        }

        private void SetsSearchBox_TextChanged(object sender, TextChangedEventArgs e)
            => ApplyFilter(SetsSearchBox.Text);

        private void BtnResetCategory_Click(object sender, RoutedEventArgs e)
            => StorageCategoryBox.Text = DefaultStorageCategory;

        private void BtnClearSets_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in _allSets) s.IsChecked = false;
        }

        private void BtnSelectAllSets_Click(object sender, RoutedEventArgs e)
        {
            foreach (var s in _filteredSets) s.IsChecked = true;
        }

        // ══ ÁRVORE NATIVA INTERATIVA ════════════════════════════════════════════

        private async Task ReloadNativeAsync()
        {
            // Cancela subscrições anteriores
            foreach (var cat in _nativeCats)
                foreach (var prop in cat.Properties)
                    prop.PropertyChanged -= OnPropCheckedChanged;

            _nativeCats.Clear();
            _nativeCatMap.Clear();
            _entries.Clear();

            var doc = NwApp.ActiveDocument;
            if (doc == null || doc.CurrentSelection.IsEmpty)
            {
                NativeSubtitle.Text = "(nenhuma seleção)";
                UpdateStatus();
                return;
            }

            var items = doc.CurrentSelection.SelectedItems.Cast<ModelItem>().ToList();

            var catMap = await Task.Run(() =>
            {
                var map = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in items)
                {
                    foreach (PropertyCategory cat in item.PropertyCategories)
                    {
                        if (!map.ContainsKey(cat.DisplayName))
                            map[cat.DisplayName] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (DataProperty prop in cat.Properties)
                        {
                            if (string.IsNullOrEmpty(prop.DisplayName)) continue;
                            if (!map[cat.DisplayName].ContainsKey(prop.DisplayName))
                                map[cat.DisplayName][prop.DisplayName] = PluginHelpers.SafeValue(prop.Value);
                        }
                    }
                }
                return map;
            });

            _nativeCatMap = catMap;

            foreach (var catKv in catMap.OrderBy(c => c.Key))
            {
                var catNode = new CatNode { Name = catKv.Key };
                foreach (var propKv in catKv.Value.OrderBy(p => p.Key))
                {
                    var pn = new PropNode
                    {
                        Name         = propKv.Key,
                        CategoryName = catKv.Key,
                        CurrentValue = propKv.Value
                    };
                    pn.PropertyChanged += OnPropCheckedChanged;
                    catNode.Properties.Add(pn);
                }
                _nativeCats.Add(catNode);
            }

            NativeSubtitle.Text = $"{catMap.Count} cats • {items.Count} elem.";
            UpdateStatus();
            UpdateCategoryInputLabel();
        }

        private void OnPropCheckedChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(PropNode.IsChecked)) return;
            var pn = (PropNode)sender;

            if (pn.IsChecked)
            {
                // Evita duplicatas
                if (_entries.Any(x => string.Equals(x.Name, pn.Name, StringComparison.OrdinalIgnoreCase)
                                   && string.Equals(x.CategoryName, pn.CategoryName, StringComparison.OrdinalIgnoreCase)))
                {
                    pn.IsChecked = false;
                    return;
                }
                _entries.Add(new AttributeEntry
                {
                    Name         = pn.Name,
                    Value        = pn.CurrentValue,
                    Type         = "string",
                    CategoryName = pn.CategoryName
                });
            }
            else
            {
                var toRemove = _entries.FirstOrDefault(x =>
                    string.Equals(x.Name, pn.Name, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(x.CategoryName, pn.CategoryName, StringComparison.OrdinalIgnoreCase));
                if (toRemove != null) _entries.Remove(toRemove);
            }

            SyncCategoryNameBox();
        }

        private void SyncCategoryNameBox()
        {
            if (RadioExcluir?.IsChecked == true) return; // Delete mode: não altera

            var cats = _entries
                .Select(e => e.CategoryName)
                .Where(c => c != null)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (cats.Count == 1)
            {
                CustomCategoryBox.Text       = cats[0];
                CustomCategoryBox.IsEnabled  = false;
                CustomCategoryBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F0F0F0"));
            }
            else
            {
                CustomCategoryBox.IsEnabled  = true;
                CustomCategoryBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FAFAFA"));
            }

            UpdateCategoryInputLabel();
        }

        private void UpdateCategoryInputLabel()
        {
            if (NewCategoryLabel == null) return;

            bool hasEntries        = _entries.Count > 0;
            bool isSingleCategory  = _entries.Select(e => e.CategoryName).Where(c => c != null).Distinct().Count() == 1;
            bool boxEditable       = CustomCategoryBox?.IsEnabled == true;

            NewCategoryLabel.Visibility = (!hasEntries || !isSingleCategory) && boxEditable
                ? Visibility.Visible : Visibility.Collapsed;
            ExistingCategoryLabel.Visibility = isSingleCategory && hasEntries
                ? Visibility.Visible : Visibility.Collapsed;
        }

        private void BtnReloadNative_Click(object sender, RoutedEventArgs e)
            => _ = ReloadNativeAsync();

        private void BtnClearChecks_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cat in _nativeCats)
                foreach (var prop in cat.Properties)
                    prop.IsChecked = false;

            _entries.Clear();
            CustomCategoryBox.Text      = string.Empty;
            CustomCategoryBox.IsEnabled = true;
            CustomCategoryBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FAFAFA"));
            UpdateCategoryInputLabel();
        }

        // ══ GRADE DE ATRIBUTOS ══════════════════════════════════════════════════

        private void BtnAddRow_Click(object sender, RoutedEventArgs e)
            => _entries.Add(new AttributeEntry { Name = "", Value = "", Type = "string" });

        private void BtnRemoveSelectedRow_Click(object sender, RoutedEventArgs e)
        {
            foreach (var entry in AttributesGrid.SelectedItems.Cast<AttributeEntry>().ToList())
            {
                // Desmarca o checkbox correspondente na árvore
                foreach (var cat in _nativeCats)
                {
                    var pn = cat.Properties.FirstOrDefault(p =>
                        string.Equals(p.Name, entry.Name, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(p.CategoryName, entry.CategoryName, StringComparison.OrdinalIgnoreCase));
                    if (pn != null) { pn.PropertyChanged -= OnPropCheckedChanged; pn.IsChecked = false; pn.PropertyChanged += OnPropCheckedChanged; }
                }
                _entries.Remove(entry);
            }
            UpdateCategoryInputLabel();
        }

        private void BtnClearGrid_Click(object sender, RoutedEventArgs e)
        {
            if (_entries.Count == 0) return;
            var r = MessageBox.Show("Clear all attributes from the grid?", "Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (r == MessageBoxResult.Yes)
            {
                // Desmarca todos na árvore
                foreach (var cat in _nativeCats)
                    foreach (var pn in cat.Properties)
                    { pn.PropertyChanged -= OnPropCheckedChanged; pn.IsChecked = false; pn.PropertyChanged += OnPropCheckedChanged; }

                _entries.Clear();
                UpdateCategoryInputLabel();
            }
        }

        private void BtnNewCategory_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cat in _nativeCats)
                foreach (var pn in cat.Properties)
                { pn.PropertyChanged -= OnPropCheckedChanged; pn.IsChecked = false; pn.PropertyChanged += OnPropCheckedChanged; }

            _entries.Clear();
            CustomCategoryBox.Text      = string.Empty;
            CustomCategoryBox.IsEnabled = true;
            CustomCategoryBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FAFAFA"));
            UpdateCategoryInputLabel();
            Keyboard.Focus(CustomCategoryBox);
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_entries.Count == 0) { MessageBox.Show("The grid is empty.", "Export", MessageBoxButton.OK, MessageBoxImage.Information); return; }
            var dlg = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Export Attributes", Filter = "CSV (*.csv)|*.csv|JSON (*.json)|*.json",
                DefaultExt = "csv", FileName = "custom_attributes"
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            try
            {
                string ext = Path.GetExtension(dlg.FileName).ToLower();
                File.WriteAllText(dlg.FileName,
                    ext == ".json" ? BuildJson() : BuildCsv(),
                    new UTF8Encoding(true));
                FooterStatus.Text = $"Exported: {Path.GetFileName(dlg.FileName)}";
            }
            catch (Exception ex) { MessageBox.Show($"Export error:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.OpenFileDialog
            {
                Title = "Import Attributes", Filter = "CSV (*.csv)|*.csv|JSON (*.json)|*.json"
            };
            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            try
            {
                string ext = Path.GetExtension(dlg.FileName).ToLower();
                var imported = ext == ".json"
                    ? ImportJson(File.ReadAllText(dlg.FileName))
                    : ImportCsv(File.ReadAllLines(dlg.FileName, Encoding.UTF8));

                if (imported.Count == 0) { MessageBox.Show("No attributes found.", "Import", MessageBoxButton.OK, MessageBoxImage.Warning); return; }

                var r = _entries.Count > 0
                    ? MessageBox.Show("Replace current attributes?", "Import", MessageBoxButton.YesNoCancel, MessageBoxImage.Question)
                    : MessageBoxResult.Yes;
                if (r == MessageBoxResult.Cancel) return;
                if (r == MessageBoxResult.Yes) _entries.Clear();
                foreach (var entry in imported) _entries.Add(entry);
                FooterStatus.Text = $"{imported.Count} attribute(s) imported.";
                SyncCategoryNameBox();
            }
            catch (Exception ex) { MessageBox.Show($"Import error:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private void AttributesGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.EditingElement is TextBox tb)
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input,
                    new Action(() => { tb.Focus(); Keyboard.Focus(tb); tb.SelectAll(); }));
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb) tb.SelectAll();
        }

        // ══ WRITE & VERIFY (seleção atual) ══════════════════════════════════════

        private void BtnWriteAndVerify_Click(object sender, RoutedEventArgs e)
        {
            AttributesGrid.CommitEdit(DataGridEditingUnit.Row, true);

            bool isDelete = RadioExcluir?.IsChecked == true;

            string categoryName = CustomCategoryBox.Text?.Trim();
            if (string.IsNullOrEmpty(categoryName))
            {
                MessageBox.Show("Enter the category name (Category field).", "Attention",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var doc = NwApp.ActiveDocument;
            if (doc == null || doc.CurrentSelection.IsEmpty)
            {
                MessageBox.Show("No elements selected in the model.", "Attention",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                UpdateStatus();
                return;
            }

            if (isDelete)
            {
                var confirm = MessageBox.Show(
                    $"Delete category \"{categoryName}\" from {doc.CurrentSelection.SelectedItems.Count} selected element(s)?\n\nThis cannot be undone.",
                    "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (confirm != MessageBoxResult.Yes) return;

                int deleted = 0, errors = 0;
                try
                {
                    var nwState = (InwOpState3)ComApiBridge.State;
                    foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
                    {
                        try
                        {
                            var nwPath   = (InwOaPath3)ComApiBridge.ToInwOaPath(item);
                            var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);
                            PluginHelpers.DeleteUserDefinedCategory(nwState, propNode, categoryName);
                            deleted++;
                        }
                        catch { errors++; }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"General error:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string msg = $"Delete complete.\n✔ {deleted} element(s) processed.";
                if (errors > 0) msg += $"\n✖ {errors} error(s).";
                MessageBox.Show(msg, "Delete", MessageBoxButton.OK,
                    errors > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
                FooterStatus.Text = $"Delete: {deleted} elem., {errors} errors.";
                return;
            }

            // Write / Edit mode
            if (_entries.Count == 0)
            {
                MessageBox.Show("Add at least one attribute to the grid.", "Attention",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool isNativeName = _nativeCatMap.ContainsKey(categoryName);
            AppendLog("─────────────────────────────────────────────");
            AppendLog($"▶ Start — category: \"{categoryName}\"");
            AppendLog($"  Exists in natives? {(isNativeName ? "YES ⚠" : "NO")}");
            AppendLog($"  Attributes in grid: {_entries.Count} | Mode: merge");

            int written = 0, writeErrors = 0;
            try
            {
                var nwState = (InwOpState3)ComApiBridge.State;
                foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
                {
                    try
                    {
                        var nwPath   = (InwOaPath3)ComApiBridge.ToInwOaPath(item);
                        var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);
                        var newProps = _entries
                            .Where(en => !string.IsNullOrWhiteSpace(en.Name))
                            .Select(en => new KeyValuePair<string, object>(en.Name, PluginHelpers.ConvertValue(en)));
                        PluginHelpers.WriteUserDefinedCategory(nwState, propNode, categoryName, newProps);
                        written++;
                    }
                    catch (Exception itemEx) { writeErrors++; AppendLog($"  ✖ Error: {itemEx.Message}"); }
                }
                AppendLog($"  Write: {written} elem., {writeErrors} error(s).");
            }
            catch (Exception ex) { AppendLog($"✖ General error: {ex.Message}"); return; }

            AppendLog("  Verifying...");
            VerifyAfterWrite(categoryName);
        }

        private void VerifyAfterWrite(string writtenCategoryName)
        {
            try
            {
                var doc = NwApp.ActiveDocument;
                if (doc == null || doc.CurrentSelection.IsEmpty) return;

                var firstItem = doc.CurrentSelection.SelectedItems.Cast<ModelItem>().First();
                var sb = new StringBuilder();
                sb.AppendLine("  ── Categories after write ──");

                foreach (PropertyCategory cat in firstItem.PropertyCategories)
                {
                    bool isTarget = string.Equals(cat.DisplayName, writtenCategoryName, StringComparison.OrdinalIgnoreCase);
                    sb.AppendLine($"  [{cat.DisplayName}]{(isTarget ? $" ◄ \"{writtenCategoryName}\"" : "")}");
                    if (!isTarget) continue;
                    foreach (DataProperty prop in cat.Properties)
                        sb.AppendLine($"      • {prop.DisplayName} = {PluginHelpers.SafeValue(prop.Value)}");
                }

                AppendLog(sb.ToString().TrimEnd());
                AppendLog("  ── COM API (UserDefined flag) ──");

                var nwState  = (InwOpState3)ComApiBridge.State;
                var nwPath   = (InwOaPath3)ComApiBridge.ToInwOaPath(firstItem);
                var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);

                bool foundNative = false, foundCustom = false;
                foreach (object obj in propNode.GUIAttributes())
                {
                    var a = obj as InwGUIAttribute2;
                    if (a == null) continue;
                    if (!string.Equals(a.name, writtenCategoryName, StringComparison.OrdinalIgnoreCase)) continue;
                    string tipo = a.UserDefined ? "[UserDefined]" : "[Native]";
                    AppendLog($"  {tipo}  \"{a.name}\"");
                    foreach (object pObj in a.Properties())
                    {
                        var p = pObj as InwOaProperty;
                        if (p != null) AppendLog($"      • {p.UserName} = {p.value}");
                    }
                    if (a.UserDefined) foundCustom = true; else foundNative = true;
                }

                AppendLog("  ── Result ──");
                if (foundNative && foundCustom)
                    AppendLog($"  ✔ COEXISTENCE: \"{writtenCategoryName}\" exists as NATIVE and UserDefined.");
                else if (foundCustom)
                    AppendLog($"  ✔ Created as UserDefined (no native with same name).");
                else if (foundNative)
                    AppendLog($"  ⚠ Only NATIVE version found. Write may have failed.");
                else
                    AppendLog($"  ⚠ Category not found via COM.");

                _ = ReloadNativeAsync();
            }
            catch (Exception ex) { AppendLog($"  ✖ Verification error: {ex.Message}"); }
        }

        // ══ SAVE / DELETE (por sets) ═════════════════════════════════════════════

        private void BtnCancel_Click(object sender, RoutedEventArgs e) => Close();

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            string storageCategory = StorageCategoryBox.Text?.Trim();
            if (string.IsNullOrEmpty(storageCategory))
            {
                MessageBox.Show("Enter the Storage Category.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var checkedSets = _allSets.Where(s => s.IsChecked).ToList();
            if (checkedSets.Count == 0)
            {
                MessageBox.Show("Select at least one Set to delete.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"Delete category \"{storageCategory}\" from elements of {checkedSets.Count} selected set(s)?\n\nThis cannot be undone.",
                "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            ExecuteWriteBySets(storageCategory, checkedSets, delete: true);
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            AttributesGrid.CommitEdit(DataGridEditingUnit.Row, true);

            string storageCategory = StorageCategoryBox.Text?.Trim();
            if (string.IsNullOrEmpty(storageCategory))
            {
                MessageBox.Show("Enter the Storage Category.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var checkedSets = _allSets.Where(s => s.IsChecked).ToList();
            if (checkedSets.Count == 0)
            {
                MessageBox.Show("Select at least one Selection Set.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            ExecuteWriteBySets(storageCategory, checkedSets, delete: false);
        }

        private void ExecuteWriteBySets(string storageCategory, List<SetEntry> sets, bool delete)
        {
            var doc = NwApp.ActiveDocument;
            if (doc == null) { MessageBox.Show("No active document."); return; }

            var allItems = new HashSet<ModelItem>();
            foreach (var set in sets)
            {
                try { foreach (ModelItem item in set.Source.GetSelectedItems(doc)) allItems.Add(item); }
                catch { }
            }

            if (allItems.Count == 0)
            {
                MessageBox.Show("The selected sets contain no elements.", "Attention", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string awpValue = string.Join(", ", sets.Select(s => s.Name));
            int written = 0, errors = 0;
            string firstError = null;

            try
            {
                var nwState = (InwOpState3)ComApiBridge.State;
                foreach (ModelItem item in allItems)
                {
                    try
                    {
                        var nwPath   = (InwOaPath3)ComApiBridge.ToInwOaPath(item);
                        var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);

                        if (delete)
                        {
                            PluginHelpers.DeleteUserDefinedCategory(nwState, propNode, storageCategory);
                        }
                        else
                        {
                            var newProps = new List<KeyValuePair<string, object>>
                            {
                                new KeyValuePair<string, object>(AwpFieldName, awpValue)
                            };
                            foreach (var entry in _entries)
                                if (!string.IsNullOrWhiteSpace(entry.Name))
                                    newProps.Add(new KeyValuePair<string, object>(entry.Name, PluginHelpers.ConvertValue(entry)));

                            PluginHelpers.WriteUserDefinedCategory(nwState, propNode, storageCategory, newProps);
                        }
                        written++;
                    }
                    catch (Exception itemEx)
                    {
                        errors++;
                        if (firstError == null) firstError = itemEx.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"General error:\n{ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string verb = delete ? "Delete" : "Save";
            string msg  = $"{verb} complete.\n✔ {written} element(s) processed.";
            if (errors > 0) msg += $"\n✖ {errors} error(s).\n\nDetail: {firstError}";

            MessageBox.Show(msg, verb, MessageBoxButton.OK,
                errors > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

            FooterStatus.Text = $"{verb}: {written} elem., {errors} errors.";
            if (!delete && errors == 0) Close();
        }

        // ══ LOG ══════════════════════════════════════════════════════════════════

        private void AppendLog(string line) => LogText.Text += line + "\n";

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
            => LogText.Text = string.Empty;

        // ══ EXPORT / IMPORT ══════════════════════════════════════════════════════

        private string BuildCsv()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Name;Value;Type");
            foreach (var e in _entries)
                sb.AppendLine($"{PluginHelpers.EscapeCsv(e.Name)};{PluginHelpers.EscapeCsv(e.Value)};{PluginHelpers.EscapeCsv(e.Type)}");
            return sb.ToString();
        }

        private string BuildJson()
        {
            var sb   = new StringBuilder();
            var list = _entries.ToList();
            sb.AppendLine("[");
            for (int i = 0; i < list.Count; i++)
            {
                var e = list[i];
                sb.Append($"  {{\"name\":\"{PluginHelpers.EscapeJson(e.Name)}\",\"value\":\"{PluginHelpers.EscapeJson(e.Value)}\",\"type\":\"{PluginHelpers.EscapeJson(e.Type)}\"}}");
                if (i < list.Count - 1) sb.AppendLine(","); else sb.AppendLine();
            }
            sb.AppendLine("]");
            return sb.ToString();
        }

        private static List<AttributeEntry> ImportCsv(string[] lines)
        {
            var result = new List<AttributeEntry>();
            foreach (var line in lines.Skip(1))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                var parts = line.Split(';');
                result.Add(new AttributeEntry
                {
                    Name  = parts.Length > 0 ? parts[0].Trim('"') : "",
                    Value = parts.Length > 1 ? parts[1].Trim('"') : "",
                    Type  = parts.Length > 2 ? parts[2].Trim('"') : "string"
                });
            }
            return result;
        }

        private static List<AttributeEntry> ImportJson(string json)
        {
            var result  = new List<AttributeEntry>();
            var entries = json.Split(new[] { "},{", "}, {" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var block in entries)
            {
                string name  = PluginHelpers.ExtractJsonValue(block, "name");
                string value = PluginHelpers.ExtractJsonValue(block, "value");
                string type  = PluginHelpers.ExtractJsonValue(block, "type");
                if (!string.IsNullOrEmpty(name))
                    result.Add(new AttributeEntry { Name = name, Value = value, Type = type ?? "string" });
            }
            return result;
        }
    }

    // ── MODELOS ──────────────────────────────────────────────────────────────────

    /// <summary>Entrada de atributo na grade. Partilhada por toda a janela.</summary>
    public class AttributeEntry : INotifyPropertyChanged
    {
        private string _name, _value, _type = "string", _categoryName;

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }
        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(nameof(Value)); }
        }
        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(nameof(Type)); }
        }
        /// <summary>Categoria de origem (preenchida quando vinda da árvore nativa).</summary>
        public string CategoryName
        {
            get => _categoryName;
            set { _categoryName = value; OnPropertyChanged(nameof(CategoryName)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>Nó de categoria na árvore nativa interativa.</summary>
    public class CatNode : INotifyPropertyChanged
    {
        private bool _isExpanded = true;

        public string Name { get; set; }

        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(nameof(IsExpanded)); }
        }

        public ObservableCollection<PropNode> Properties { get; set; }
            = new ObservableCollection<PropNode>();

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>Nó de propriedade com checkbox na árvore nativa interativa.</summary>
    public class PropNode : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Name         { get; set; }
        public string CategoryName { get; set; }
        public string CurrentValue { get; set; }

        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(nameof(IsChecked)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>Entry de set de seleção na lista da coluna central.</summary>
    public class SetEntry : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string       Name      { get; set; }
        public int          ItemCount { get; set; }
        public SelectionSet Source    { get; set; }

        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(nameof(IsChecked)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
