using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Navisworks.Api;
using NwApp = Autodesk.Navisworks.Api.Application;

namespace TestePlugin
{
    // ══════════════════════════════════════════════════════════════════════
    //  ReadAttributesWindow.xaml.cs
    //
    //  Responsabilidades:
    //    • Exibir categorias e propriedades BIM dos elementos selecionados.
    //    • Navegar entre múltiplos elementos (◄◄ ◄ ► ►►).
    //    • Filtrar categorias e propriedades por texto.
    //    • Copiar propriedades selecionadas para o clipboard no formato TSV.
    //
    //  Modelo de dados:
    //    _data            — dicionário categoria → lista de PropertyRow
    //    _visibleCategories — lista filtrada exibida no painel esquerdo
    //    _selectedCatRows — propriedades da categoria ativa (pré-filtro)
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Janela de leitura de atributos BIM: exibe as categorias e propriedades
    /// dos elementos selecionados no Navisworks com suporte a filtro e navegação.
    /// </summary>
    public partial class ReadAttributesWindow : Window
    {
        // ── Campos ────────────────────────────────────────────────────────────────

        /// <summary>Dados do elemento ativo: categoria → lista de propriedades ordenadas alfabeticamente.</summary>
        private Dictionary<string, List<PropertyRow>> _data =
            new Dictionary<string, List<PropertyRow>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Lista visível no painel esquerdo (resultado do filtro de categoria).</summary>
        private List<CategoryEntry> _visibleCategories = new List<CategoryEntry>();

        /// <summary>Propriedades da categoria selecionada antes de aplicar o filtro de propriedade.</summary>
        private List<PropertyRow> _selectedCatRows = new List<PropertyRow>();

        /// <summary>Todos os elementos atualmente selecionados no Navisworks.</summary>
        private List<ModelItem> _allItems = new List<ModelItem>();

        /// <summary>Índice (0-based) do elemento exibido no momento.</summary>
        private int _currentIndex = 0;

        // ── Construtor ────────────────────────────────────────────────────────────

        public ReadAttributesWindow()
        {
            InitializeComponent();
            Loaded += (s, e) =>
            {
                RefreshData();
                // Força a janela a receber foco do sistema operacional (necessário no host Win32 do Navisworks)
                Activate();
                Topmost = true;
                Topmost = false;
                Keyboard.Focus(CatSearchBox);
            };
        }

        // ── Carga de dados ────────────────────────────────────────────────────────

        /// <summary>
        /// Lê os elementos selecionados no Navisworks, inicializa _allItems
        /// e carrega o primeiro elemento.
        /// </summary>
        private void RefreshData()
        {
            _allItems.Clear();
            _currentIndex = 0;
            NavPanel.Visibility = Visibility.Collapsed;
            ClearPanels();

            try
            {
                var doc = NwApp.ActiveDocument;
                if (doc == null) { StatusText.Text = "Nenhum documento ativo."; return; }

                var selected = doc.CurrentSelection.SelectedItems;
                if (selected.Count == 0) { StatusText.Text = "Nenhum elemento selecionado."; return; }

                _allItems = selected.ToList();

                // Barra de navegação só aparece com múltiplos elementos
                if (_allItems.Count > 1)
                    NavPanel.Visibility = Visibility.Visible;

                LoadItemData(0);
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Erro: {ex.Message}";
            }
        }

        /// <summary>
        /// Carrega e exibe as propriedades do elemento no índice informado.
        /// </summary>
        private void LoadItemData(int index)
        {
            _data.Clear();
            ClearPanels();

            try
            {
                var item = _allItems[index];

                // Agrega propriedades do elemento pelo nome de categoria
                var catDict = new Dictionary<string, Dictionary<string, PropertyRow>>(StringComparer.OrdinalIgnoreCase);

                foreach (PropertyCategory cat in item.PropertyCategories)
                {
                    string catName = cat.DisplayName ?? cat.Name ?? "(sem nome)";
                    if (!catDict.ContainsKey(catName))
                        catDict[catName] = new Dictionary<string, PropertyRow>(StringComparer.OrdinalIgnoreCase);

                    foreach (DataProperty prop in cat.Properties)
                    {
                        string pName  = prop.DisplayName ?? prop.Name ?? "(sem nome)";
                        string pValue = SafeValue(prop.Value);
                        string pType  = SafeType(prop.Value);
                        catDict[catName][pName] = new PropertyRow { Property = pName, Value = pValue, Type = pType };
                    }
                }

                _data = catDict.ToDictionary(
                    kv => kv.Key,
                    kv => kv.Value.Values.OrderBy(r => r.Property).ToList(),
                    StringComparer.OrdinalIgnoreCase);

                string displayName = item.DisplayName ?? $"Elemento {index + 1}";
                int totalProps = _data.Values.Sum(l => l.Count);
                StatusText.Text = $"{displayName}  |  {_data.Count} categorias  |  {totalProps} propriedades";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Erro ao carregar elemento: {ex.Message}";
            }

            UpdateNavigation();
            FilterCategories(CatSearchBox.Text);
        }

        /// <summary>
        /// Limpa os painéis de categorias e propriedades para o estado inicial.
        /// </summary>
        private void ClearPanels()
        {
            CategoriesListView.ItemsSource = null;
            PropertiesGrid.ItemsSource     = null;
            PropPanelTitle.Text            = "Selecione uma categoria";
            _selectedCatRows               = new List<PropertyRow>();
        }

        // ── Navegação entre elementos ─────────────────────────────────────────────

        /// <summary>
        /// Atualiza o contador "X / N" e o estado habilitado dos botões de navegação.
        /// </summary>
        private void UpdateNavigation()
        {
            int total          = _allItems.Count;
            NavText.Text       = total > 0 ? $"{_currentIndex + 1} / {total}" : "—";
            BtnFirst.IsEnabled = _currentIndex > 0;
            BtnPrev.IsEnabled  = _currentIndex > 0;
            BtnNext.IsEnabled  = _currentIndex < total - 1;
            BtnLast.IsEnabled  = _currentIndex < total - 1;
        }

        /// <summary>Navega para o primeiro elemento da seleção.</summary>
        private void BtnFirst_Click(object sender, RoutedEventArgs e)
        {
            _currentIndex = 0;
            LoadItemData(_currentIndex);
        }

        /// <summary>Navega para o elemento anterior (decrementa o índice).</summary>
        private void BtnPrev_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex > 0) { _currentIndex--; LoadItemData(_currentIndex); }
        }

        /// <summary>Navega para o próximo elemento (incrementa o índice).</summary>
        private void BtnNext_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex < _allItems.Count - 1) { _currentIndex++; LoadItemData(_currentIndex); }
        }

        /// <summary>Navega para o último elemento da seleção.</summary>
        private void BtnLast_Click(object sender, RoutedEventArgs e)
        {
            _currentIndex = _allItems.Count - 1;
            LoadItemData(_currentIndex);
        }

        // ── Filtros ───────────────────────────────────────────────────────────────

        /// <summary>
        /// Filtra a lista de categorias pelo texto digitado em <paramref name="query"/>.
        /// Se <paramref name="query"/> estiver em branco, exibe todas as categorias.
        /// </summary>
        private void FilterCategories(string query)
        {
            IEnumerable<string> keys = _data.Keys.OrderBy(k => k);

            if (!string.IsNullOrWhiteSpace(query))
            {
                var q = query.Trim().ToLowerInvariant();
                keys = keys.Where(k => k.ToLowerInvariant().Contains(q));
            }

            _visibleCategories = keys
                .Select(k => new CategoryEntry { Name = k, Count = _data[k].Count })
                .ToList();

            CategoriesListView.ItemsSource = _visibleCategories;
        }

        /// <summary>
        /// Filtra as propriedades da categoria ativa pelo texto digitado em <paramref name="query"/>.
        /// A busca abrange os campos Propriedade, Valor e Tipo simultaneamente.
        /// </summary>
        private void FilterProperties(string query)
        {
            IEnumerable<PropertyRow> rows = _selectedCatRows;

            if (!string.IsNullOrWhiteSpace(query))
            {
                var q = query.Trim().ToLowerInvariant();
                rows = rows.Where(r =>
                    (r.Property?.ToLowerInvariant().Contains(q) == true) ||
                    (r.Value?.ToLowerInvariant().Contains(q)    == true) ||
                    (r.Type?.ToLowerInvariant().Contains(q)     == true));
            }

            PropertiesGrid.ItemsSource = rows.ToList();
        }

        // ── Eventos ──────────────────────────────────────────────────────────────

        /// <summary>Seleciona todo o texto ao clicar em qualquer campo de busca.</summary>
        private void SearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb) tb.SelectAll();
        }

        /// <summary>Dispara o filtro de categorias sempre que o texto do campo muda.</summary>
        private void CatSearch_TextChanged(object sender, TextChangedEventArgs e)
            => FilterCategories(CatSearchBox.Text);

        /// <summary>Dispara o filtro de propriedades sempre que o texto do campo muda.</summary>
        private void PropSearch_TextChanged(object sender, TextChangedEventArgs e)
            => FilterProperties(PropSearchBox.Text);

        private void CategoriesListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CategoriesListView.SelectedItem is CategoryEntry entry &&
                _data.TryGetValue(entry.Name, out var rows))
            {
                _selectedCatRows    = rows;
                PropPanelTitle.Text = entry.Name;

                // Desconecta o evento antes de limpar para evitar chamada dupla de FilterProperties
                PropSearchBox.TextChanged -= PropSearch_TextChanged;
                PropSearchBox.Text         = "";
                PropSearchBox.TextChanged += PropSearch_TextChanged;

                FilterProperties("");
            }
        }

        /// <summary>Recarrega os dados da seleção atual do Navisworks.</summary>
        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
            => RefreshData();

        /// <summary>Fecha a janela.</summary>
        private void BtnClose_Click(object sender, RoutedEventArgs e)
            => Close();

        /// <summary>
        /// Copia as linhas selecionadas na grade (ou todas, se nenhuma estiver marcada)
        /// para o clipboard no formato TSV (Tab-Separated Values).
        /// </summary>
        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            // Usa linhas selecionadas; se nenhuma, copia tudo o que está visível (filtrado)
            IEnumerable<PropertyRow> source;

            if (PropertiesGrid.SelectedItems.Count > 0)
                source = PropertiesGrid.SelectedItems.OfType<PropertyRow>();
            else
                source = PropertiesGrid.ItemsSource as IEnumerable<PropertyRow>
                         ?? Enumerable.Empty<PropertyRow>();

            var sb = new StringBuilder();
            sb.AppendLine("Propriedade\tValor\tTipo");
            foreach (var row in source)
                sb.AppendLine($"{row.Property}\t{row.Value}\t{row.Type}");

            try
            {
                Clipboard.SetText(sb.ToString());
                MessageBox.Show("Dados copiados (formato TSV).", "Copiar",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao copiar: {ex.Message}", "Erro",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ── Helpers de leitura segura ─────────────────────────────────────────────

        /// <summary>
        /// Retorna a representação textual de <paramref name="v"/> sem lançar exceções.
        /// Tenta <c>ToDisplayString()</c> primeiro; cai para <c>ToString()</c>;
        /// retorna <c>"(erro)"</c> se ambas falharem.
        /// </summary>
        private static string SafeValue(VariantData v)
        {
            if (v == null) return "";
            try   { return v.ToDisplayString() ?? ""; }
            catch { }
            try   { return v.ToString() ?? ""; }
            catch { return "(erro)"; }
        }

        /// <summary>
        /// Retorna o nome do tipo de dado de <paramref name="v"/> sem lançar exceções.
        /// </summary>
        private static string SafeType(VariantData v)
        {
            if (v == null) return "";
            try   { return v.DataType.ToString(); }
            catch { return ""; }
        }
    }

    // ── Modelos de dados ──────────────────────────────────────────────────────────

    /// <summary>
    /// Representa uma categoria de propriedades BIM exibida no painel esquerdo.
    /// Contém o nome da categoria e a quantidade de propriedades que ela possui.
    /// </summary>
    public class CategoryEntry
    {
        /// <summary>Nome da categoria (ex.: "Element", "Identity Data").</summary>
        public string Name  { get; set; }

        /// <summary>Número de propriedades nesta categoria.</summary>
        public int    Count { get; set; }
    }

    /// <summary>
    /// Representa uma linha do DataGrid de propriedades no painel direito.
    /// </summary>
    public class PropertyRow
    {
        /// <summary>Nome da propriedade (ex.: "Area", "Level").</summary>
        public string Property { get; set; }

        /// <summary>Valor da propriedade como texto formatado.</summary>
        public string Value    { get; set; }

        /// <summary>Tipo do dado (ex.: "String", "Double", "Integer").</summary>
        public string Type     { get; set; }
    }
}
