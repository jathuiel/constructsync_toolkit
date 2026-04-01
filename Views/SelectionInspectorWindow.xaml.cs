using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using NwApp = Autodesk.Navisworks.Api.Application;
using WpfColor = System.Windows.Media.Color;

namespace SetAtributesToolkit
{
    public partial class SelectionInspectorWindow : Window
    {
        // ── Estado ───────────────────────────────────────────────────────────
        private List<ModelItem> _allItems = new List<ModelItem>();
        private readonly ObservableCollection<CheckableItem> _categories = new ObservableCollection<CheckableItem>();
        private List<ColDef> _currentColumns = new List<ColDef>();
        private bool _rebuilding;

        public SelectionInspectorWindow()
        {
            InitializeComponent();
            CategoriesList.ItemsSource = _categories;

            Loaded += async (s, e) =>
            {
                Activate();
                Topmost = true;
                Topmost = false;
                await LoadSelectionAsync();
            };
        }

        // ── CARREGAMENTO DA SELEÇÃO ──────────────────────────────────────────

        private async Task LoadSelectionAsync()
        {
            SetStatus("Carregando seleção...", StatusLevel.Info);

            var doc = NwApp.ActiveDocument;
            if (doc == null || doc.CurrentSelection.IsEmpty)
            {
                SetStatus("Nenhum elemento selecionado no modelo.", StatusLevel.Warning);
                TxtSubtitle.Text = "Nenhum elemento selecionado.";
                return;
            }

            // Copia os itens selecionados para lista local (operação no thread da UI)
            var selection = doc.CurrentSelection.SelectedItems;
            var items = new List<ModelItem>();
            foreach (var item in selection)
                items.Add(item);

            _allItems = items;

            // Extrai categorias únicas em background
            var catNames = await Task.Run(() =>
            {
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in items)
                    foreach (PropertyCategory cat in item.PropertyCategories)
                        if (!string.IsNullOrWhiteSpace(cat.DisplayName))
                            names.Add(cat.DisplayName);
                return names.OrderBy(n => n).ToList();
            });

            _categories.Clear();
            foreach (var name in catNames)
            {
                var entry = new CheckableItem { Name = name };
                entry.PropertyChanged += OnCategoryCheckedChanged;
                _categories.Add(entry);
            }

            TxtSubtitle.Text = $"{_allItems.Count} elemento(s) selecionado(s) | {_categories.Count} categorias";
            SetStatus($"{_allItems.Count} elemento(s) | {_categories.Count} categoria(s) disponíveis.", StatusLevel.Info);
        }

        private async void OnCategoryCheckedChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CheckableItem.IsChecked))
                await RebuildGridAsync();
        }

        // ── RECONSTRUÇÃO DO GRID ─────────────────────────────────────────────

        private async Task RebuildGridAsync()
        {
            if (_rebuilding) return;
            _rebuilding = true;

            try
            {
                var checkedCats = _categories.Where(c => c.IsChecked).Select(c => c.Name).ToList();

                InspectorGrid.Columns.Clear();
                InspectorGrid.ItemsSource = null;

                if (checkedCats.Count == 0 || _allItems.Count == 0)
                {
                    TxtEmptyHint.Visibility = Visibility.Visible;
                    BtnExportCsv.IsEnabled = false;
                    BtnExportClip.IsEnabled = false;
                    TxtStatus.Text = $"{_allItems.Count} elemento(s) | 0 categorias | 0 colunas";
                    SetStatus("Marque categorias para visualizar.", StatusLevel.Info);
                    return;
                }

                TxtEmptyHint.Visibility = Visibility.Collapsed;
                SetStatus("Construindo tabela...", StatusLevel.Info);

                var result = await Task.Run(() => BuildDataTable(checkedCats));
                _currentColumns = result.Columns;

                // Adiciona colunas no thread da UI
                var headerTemplate = (DataTemplate)FindResource("ColHeaderTemplate");
                foreach (var col in result.Columns)
                {
                    InspectorGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = new ColumnHeaderData { Category = col.Category, Property = col.Property },
                        HeaderTemplate = headerTemplate,
                        Binding = new Binding($"[{col.Key}]"),
                        Width = new DataGridLength(1, DataGridLengthUnitType.Auto),
                        MinWidth = 110,
                        IsReadOnly = true
                    });
                }

                InspectorGrid.ItemsSource = result.Table.DefaultView;

                BtnExportCsv.IsEnabled = true;
                BtnExportClip.IsEnabled = true;
                SetStatus($"{_allItems.Count} elemento(s) | {checkedCats.Count} categoria(s) | {result.Columns.Count} coluna(s)", StatusLevel.Success);
            }
            catch (Exception ex)
            {
                SetStatus($"Erro ao construir grid: {ex.Message}", StatusLevel.Error);
            }
            finally
            {
                _rebuilding = false;
            }
        }

        // ── BUILD DATATABLE (background thread) ─────────────────────────────

        private BuildResult BuildDataTable(List<string> checkedCats)
        {
            // 1. Descobre todas as propriedades por categoria (percorre apenas categorias selecionadas)
            var catPropMap = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var catName in checkedCats)
                catPropMap[catName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var item in _allItems)
            {
                foreach (PropertyCategory cat in item.PropertyCategories)
                {
                    if (!catPropMap.TryGetValue(cat.DisplayName, out var propSet)) continue;
                    foreach (DataProperty prop in cat.Properties)
                        if (!string.IsNullOrWhiteSpace(prop.DisplayName))
                            propSet.Add(prop.DisplayName);
                }
            }

            // 2. Cria lista ordenada de colunas
            var columns = new List<ColDef>();
            int idx = 0;
            foreach (var catName in checkedCats)
            {
                foreach (var propName in catPropMap[catName].OrderBy(p => p))
                {
                    columns.Add(new ColDef
                    {
                        Category = catName,
                        Property = propName,
                        Key = $"c{idx++}"
                    });
                }
            }

            // 3. Monta DataTable
            var dt = new DataTable();
            foreach (var col in columns)
                dt.Columns.Add(new DataColumn(col.Key, typeof(string)));

            // 4. Preenche linhas (uma por ModelItem)
            foreach (var item in _allItems)
            {
                // Cache das categorias deste item
                var catCache = new Dictionary<string, PropertyCategory>(StringComparer.OrdinalIgnoreCase);
                foreach (PropertyCategory cat in item.PropertyCategories)
                    catCache[cat.DisplayName] = cat;

                var row = dt.NewRow();
                foreach (var col in columns)
                {
                    if (!catCache.TryGetValue(col.Category, out var category)) continue;
                    foreach (DataProperty prop in category.Properties)
                    {
                        if (!string.Equals(prop.DisplayName, col.Property, StringComparison.OrdinalIgnoreCase)) continue;
                        row[col.Key] = SafeValue(prop.Value);
                        break;
                    }
                }
                dt.Rows.Add(row);
            }

            return new BuildResult { Table = dt, Columns = columns };
        }

        private static string SafeValue(VariantData v)
        {
            if (v == null || v.DataType == VariantDataType.None) return string.Empty;
            try { return v.ToDisplayString() ?? string.Empty; }
            catch { return string.Empty; }
        }

        // ── BOTÕES DE SELEÇÃO ────────────────────────────────────────────────

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cat in _categories) cat.IsChecked = true;
        }

        private void BtnClearAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (var cat in _categories) cat.IsChecked = false;
        }

        // ── EXPORT ───────────────────────────────────────────────────────────

        private void BtnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.SaveFileDialog
            {
                Title = "Exportar CSV",
                Filter = "CSV (*.csv)|*.csv",
                DefaultExt = "csv",
                FileName = "selection_inspector"
            };

            if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

            try
            {
                ExportToCsv(dlg.FileName);
                SetStatus($"CSV exportado: {System.IO.Path.GetFileName(dlg.FileName)}", StatusLevel.Success);
            }
            catch (Exception ex)
            {
                SetStatus($"Erro ao exportar: {ex.Message}", StatusLevel.Error);
            }
        }

        private void BtnExportClip_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var sb = BuildTsv();
                Clipboard.SetText(sb.ToString());
                SetStatus("Dados copiados para a área de transferência.", StatusLevel.Success);
            }
            catch (Exception ex)
            {
                SetStatus($"Erro ao copiar: {ex.Message}", StatusLevel.Error);
            }
        }

        private void ExportToCsv(string path)
        {
            var sb = BuildCsv();
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        private StringBuilder BuildCsv()
        {
            var sb = new StringBuilder();
            if (!(InspectorGrid.ItemsSource is DataView dv)) return sb;

            // Cabeçalho
            var headers = _currentColumns.Select(c => EscapeCsv($"{c.Category} — {c.Property}"));
            sb.AppendLine(string.Join(";", headers));

            // Linhas
            foreach (DataRowView rv in dv)
            {
                var vals = _currentColumns.Select(c => EscapeCsv(rv.Row[c.Key]?.ToString() ?? ""));
                sb.AppendLine(string.Join(";", vals));
            }
            return sb;
        }

        private StringBuilder BuildTsv()
        {
            var sb = new StringBuilder();
            if (!(InspectorGrid.ItemsSource is DataView dv)) return sb;

            sb.AppendLine(string.Join("\t", _currentColumns.Select(c => $"{c.Category} — {c.Property}")));
            foreach (DataRowView rv in dv)
                sb.AppendLine(string.Join("\t", _currentColumns.Select(c => rv.Row[c.Key]?.ToString() ?? "")));

            return sb;
        }

        private static string EscapeCsv(string s)
        {
            if (s == null) return "";
            return (s.Contains(';') || s.Contains('"') || s.Contains('\n'))
                ? $"\"{s.Replace("\"", "\"\"")}\""
                : s;
        }

        // ── FECHAR ───────────────────────────────────────────────────────────

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        // ── STATUS ───────────────────────────────────────────────────────────

        private enum StatusLevel { Info, Success, Warning, Error }

        private void SetStatus(string message, StatusLevel level)
        {
            TxtStatus.Text = message;
            StatusIndicator.Fill = level switch
            {
                StatusLevel.Success => new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E7D32")),
                StatusLevel.Warning => new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#ED6C02")),
                StatusLevel.Error => new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#D32F2F")),
                _ => new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#0288D1")),
            };
        }
    }

    // ── Modelos específicos desta view ───────────────────────────────────────
    // CheckableItem está em Models/CheckableItem.cs (compartilhado)

    public class ColumnHeaderData
    {
        public string Category { get; set; }
        public string Property { get; set; }
    }

    internal class ColDef
    {
        public string Category { get; set; }
        public string Property { get; set; }
        public string Key { get; set; }
    }

    internal class BuildResult
    {
        public DataTable Table { get; set; }
        public List<ColDef> Columns { get; set; }
    }
}
