using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using NwApp = Autodesk.Navisworks.Api.Application;
using WpfColor = System.Windows.Media.Color;

namespace SetAtributesToolkit
{
    public partial class WriteAttributesWindow : Window
    {
        private readonly ObservableCollection<AttributeEntry> _entries =
            new ObservableCollection<AttributeEntry>();

        private readonly ObservableCollection<CatNode> _catNodes =
            new ObservableCollection<CatNode>();

        private List<ModelItem> _allItems = new List<ModelItem>();

        // Cores por modo
        private static readonly WpfColor ColorGravar = (WpfColor)ColorConverter.ConvertFromString("#E8F5E9");
        private static readonly WpfColor ColorEditar = (WpfColor)ColorConverter.ConvertFromString("#FFF8E1");
        private static readonly WpfColor ColorExcluir = (WpfColor)ColorConverter.ConvertFromString("#FFEBEE");
        private static readonly WpfColor ColorGravarBorder = (WpfColor)ColorConverter.ConvertFromString("#A5D6A7");
        private static readonly WpfColor ColorEditarBorder = (WpfColor)ColorConverter.ConvertFromString("#FFE082");
        private static readonly WpfColor ColorExcluirBorder = (WpfColor)ColorConverter.ConvertFromString("#EF9A9A");

        public WriteAttributesWindow()
        {
            InitializeComponent();
            AttributesGrid.ItemsSource = _entries;
            SetsTreeView.ItemsSource = _catNodes;

            Loaded += async (s, e) =>
            {
                ApplyModeUI();
                await LoadSelectionAsync();
                UpdateStatus();
                Activate();
                Topmost = true;
                Topmost = false;
                Keyboard.Focus(CategoryNameBox);
            };
        }

        // ── MODO ────────────────────────────────────────────────────────────────

        private void RadioModo_Changed(object sender, RoutedEventArgs e)
        {
            if (CategoryNameBox == null) return;
            ApplyModeUI();
            UpdateStatus();
        }

        private void ApplyModeUI()
        {
            bool isGravar = RadioGravar?.IsChecked == true;
            bool isEditar = RadioEditar?.IsChecked == true;
            bool isExcluir = RadioExcluir?.IsChecked == true;

            UpdateCategoryInputLabel();

            // Header de estado
            if (isGravar)
            {
                ModeHeaderBorder.Background = new SolidColorBrush(ColorGravar);
                ModeHeaderBorder.BorderBrush = new SolidColorBrush(ColorGravarBorder);
                ModeHeaderBorder.BorderThickness = new Thickness(1);
                ModeHeaderIcon.Text = "●";
                ModeHeaderIcon.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E7D32"));
                ModeHeaderText.Text = "Modo: Gravação — preencha os atributos e clique em Gravar";
                ModeHeaderText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#1B5E20"));
                SectionTitle.Text = "Atributos a Gravar";
            }
            else if (isEditar)
            {
                ModeHeaderBorder.Background = new SolidColorBrush(ColorEditar);
                ModeHeaderBorder.BorderBrush = new SolidColorBrush(ColorEditarBorder);
                ModeHeaderBorder.BorderThickness = new Thickness(1);
                ModeHeaderIcon.Text = "●";
                ModeHeaderIcon.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F57F17"));
                ModeHeaderText.Text = "Modo: Edição — carregue um set e edite os valores";
                ModeHeaderText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#E65100"));
                SectionTitle.Text = "Atributos a Editar";
            }
            else // Excluir
            {
                ModeHeaderBorder.Background = new SolidColorBrush(ColorExcluir);
                ModeHeaderBorder.BorderBrush = new SolidColorBrush(ColorExcluirBorder);
                ModeHeaderBorder.BorderThickness = new Thickness(1);
                ModeHeaderIcon.Text = "●";
                ModeHeaderIcon.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#C62828"));
                ModeHeaderText.Text = "Modo: Exclusão — os atributos abaixo serão REMOVIDOS dos elementos";
                ModeHeaderText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#B71C1C"));
                SectionTitle.Text = "Atributos a Excluir";
            }

            // Bloqueio de edição no modo Excluir
            bool editable = isGravar || isEditar;
            CategoryNameBox.IsEnabled = editable;
            CategoryNameBox.Background = editable
                ? new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FAFAFA"))
                : new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F0F0F0"));
            AttributesGrid.IsReadOnly = !editable;
            BtnAddRow.IsEnabled = editable;

            // Coluna X: ocultar no modo Excluir (a exclusão do set todo já está clara)
            if (RemoveColumn != null)
                RemoveColumn.Visibility = editable ? Visibility.Visible : Visibility.Collapsed;

            // Painel de contexto
            if (isExcluir)
            {
                ContextPanel.Background = new SolidColorBrush(ColorExcluir);
                ContextPanel.BorderBrush = new SolidColorBrush(ColorExcluirBorder);
                ContextPanel.BorderThickness = new Thickness(1);
                ContextActionText.Text = "Você está prestes a: Excluir atributos";
                ContextActionText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#B71C1C"));
            }
            else
            {
                ContextPanel.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F5F5F5"));
                ContextPanel.BorderBrush = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#E0E0E0"));
                ContextPanel.BorderThickness = new Thickness(1);
                ContextActionText.Text = isGravar ? "Você está prestes a: Gravar atributos" : "Você está prestes a: Editar atributos";
                ContextActionText.Foreground = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E2E2E"));
            }

            // Botão principal
            if (isExcluir)
            {
                BtnWrite.Content = "Excluir Atributos dos Elementos Selecionados";
                BtnWrite.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#D32F2F"));
            }
            else
            {
                BtnWrite.Content = "Gravar nos Elementos Selecionados";
                BtnWrite.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E7D32"));
            }
        }

        // ── CARREGAMENTO DA SELEÇÃO ──────────────────────────────────────────

        private async Task LoadSelectionAsync()
        {
            _catNodes.Clear();
            _entries.Clear();
            CategoryNameBox.Text = "";

            var doc = NwApp.ActiveDocument;
            if (doc == null || doc.CurrentSelection.IsEmpty)
            {
                UpdateStatus();
                return;
            }

            // Copia os itens selecionados
            var items = doc.CurrentSelection.SelectedItems.Cast<ModelItem>().ToList();
            _allItems = items;

            // Extrai categorias e propriedades em background
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
                            if (!map[cat.DisplayName].ContainsKey(prop.DisplayName))
                                map[cat.DisplayName][prop.DisplayName] = SafeValue(prop.Value);
                        }
                    }
                }
                return map;
            });

            // Cria CatNodes com PropNodes
            foreach (var catKv in catMap.OrderBy(c => c.Key))
            {
                var catNode = new CatNode { Name = catKv.Key };
                foreach (var propKv in catKv.Value.OrderBy(p => p.Key))
                {
                    var pn = new PropNode
                    {
                        Name = propKv.Key,
                        CategoryName = catKv.Key,
                        CurrentValue = propKv.Value
                    };
                    pn.PropertyChanged += OnPropCheckedChanged;
                    catNode.Properties.Add(pn);
                }
                _catNodes.Add(catNode);
            }
        }

        private static string SafeValue(VariantData v)
        {
            if (v == null || v.DataType == VariantDataType.None) return string.Empty;
            try { return v.ToDisplayString() ?? string.Empty; }
            catch { return string.Empty; }
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
                    Name = pn.Name,
                    Value = pn.CurrentValue,
                    Type = "string",
                    CategoryName = pn.CategoryName
                });
                SyncCategoryNameBox();
            }
            else
            {
                var toRemove = _entries.FirstOrDefault(x => string.Equals(x.Name, pn.Name, StringComparison.OrdinalIgnoreCase)
                                                          && string.Equals(x.CategoryName, pn.CategoryName, StringComparison.OrdinalIgnoreCase));
                if (toRemove != null) _entries.Remove(toRemove);
                SyncCategoryNameBox();
            }
        }

        private void SyncCategoryNameBox()
        {
            var cats = _entries
                .Select(e => e.CategoryName)
                .Where(c => c != null)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (cats.Count == 1)
            {
                CategoryNameBox.Text = cats[0];
                CategoryNameBox.IsEnabled = false;
                CategoryNameBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#F0F0F0"));
            }
            else
            {
                CategoryNameBox.IsEnabled = true;
                CategoryNameBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FAFAFA"));
            }

            UpdateCategoryInputLabel();
        }

        // ── STATUS / SELEÇÃO ────────────────────────────────────────────────────

        private void UpdateStatus()
        {
            try
            {
                int count = NwApp.ActiveDocument?.CurrentSelection.SelectedItems.Count ?? 0;

                StatusText.Text = $"Elementos afetados: {count}";

                bool hasSelection = count > 0;
                BtnWrite.IsEnabled = hasSelection;
                ZeroSelectionWarning.Visibility = hasSelection ? Visibility.Collapsed : Visibility.Visible;

                if (!hasSelection)
                {
                    BtnWrite.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#BDBDBD"));
                    BtnWrite.ToolTip = "Selecione ao menos um elemento no modelo";
                }
                else
                {
                    bool isExcluir = RadioExcluir?.IsChecked == true;
                    BtnWrite.Background = isExcluir
                        ? new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#D32F2F"))
                        : new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#2E7D32"));
                    BtnWrite.ToolTip = null;
                }
            }
            catch
            {
                StatusText.Text = "Não foi possível ler a seleção.";
                BtnWrite.IsEnabled = false;
                ZeroSelectionWarning.Visibility = Visibility.Visible;
            }
        }



        private void BtnNewSet_Click(object sender, RoutedEventArgs e)
        {
            // Desmarcar todos os checkboxes
            foreach (var cat in _catNodes)
                foreach (var prop in cat.Properties)
                    prop.IsChecked = false;

            _entries.Clear();
            CategoryNameBox.Text = "";
            CategoryNameBox.IsEnabled = true;
            CategoryNameBox.Background = new SolidColorBrush((WpfColor)ColorConverter.ConvertFromString("#FAFAFA"));
            UpdateCategoryInputLabel();
            Keyboard.Focus(CategoryNameBox);
        }

        private void UpdateCategoryInputLabel()
        {
            bool hasEntries = _entries.Count > 0;
            bool isSingleCategory = _entries.Select(e => e.CategoryName).Where(c => c != null).Distinct().Count() == 1;

            if (NewCategoryLabel != null)
                NewCategoryLabel.Visibility = (!hasEntries || !isSingleCategory) && CategoryNameBox.IsEnabled
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            if (ExistingCategoryLabel != null)
                ExistingCategoryLabel.Visibility = (isSingleCategory && hasEntries)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
        }

        // ── GRADE ───────────────────────────────────────────────────────────────

        private void AttributesGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e) { }

        /// <summary>Foco via Dispatcher — necessário porque o host Win32 do Navisworks não repassa teclado ao WPF DataGrid.</summary>
        private void AttributesGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.EditingElement is TextBox tb)
            {
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
                {
                    tb.Focus();
                    Keyboard.Focus(tb);
                    tb.SelectAll();
                }));
            }
        }

        private void EditingTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb) tb.SelectAll();
        }

        private void BtnAddRow_Click(object sender, RoutedEventArgs e)
            => _entries.Add(new AttributeEntry { Name = "", Value = "", Type = "string" });

        private void BtnRemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if (((Button)sender).Tag is AttributeEntry entry)
            {
                var result = MessageBox.Show(
                    $"Remover o atributo \"{entry.Name}\"?",
                    "Confirmar Remoção",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                    _entries.Remove(entry);
            }
        }

        // ── HELPERS ─────────────────────────────────────────────────────────────

        private Dictionary<string, object> ReadExistingProps(InwGUIPropertyNode2 propNode, string categoryName)
        {
            var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            try
            {
                foreach (object obj in propNode.GUIAttributes())
                {
                    var a = obj as InwGUIAttribute2;
                    if (a == null || !a.UserDefined) continue;
                    if (!string.Equals(a.name, categoryName, StringComparison.OrdinalIgnoreCase)) continue;

                    foreach (object pObj in a.Properties())
                    {
                        var p = pObj as InwOaProperty;
                        if (p != null)
                            result[p.UserName] = p.value;
                    }
                    break;
                }
            }
            catch { }

            return result;
        }

        // ── EXECUTAR ────────────────────────────────────────────────────────────

        private void BtnWrite_Click(object sender, RoutedEventArgs e)
        {
            AttributesGrid.CommitEdit(DataGridEditingUnit.Row, true);

            bool isExcluir = RadioExcluir?.IsChecked == true;

            string categoryName = CategoryNameBox.Text?.Trim();
            if (string.IsNullOrEmpty(categoryName))
            {
                MessageBox.Show("Informe o nome da categoria.", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var doc = NwApp.ActiveDocument;
            if (doc == null) { MessageBox.Show("Nenhum documento ativo."); return; }

            var selected = doc.CurrentSelection.SelectedItems;
            if (selected.Count == 0)
            {
                MessageBox.Show("Nenhum elemento selecionado no modelo.", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                UpdateStatus();
                return;
            }

            if (!isExcluir && _entries.Count == 0)
            {
                MessageBox.Show("Adicione ao menos um atributo.", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Confirmação extra para Excluir
            if (isExcluir)
            {
                var confirm = MessageBox.Show(
                    $"Você está prestes a EXCLUIR a categoria \"{categoryName}\" de {selected.Count} elemento(s).\n\nEssa ação não pode ser desfeita. Continuar?",
                    "Confirmar Exclusão",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (confirm != MessageBoxResult.Yes) return;
            }

            int written = 0, errors = 0;
            string firstError = null;

            try
            {
                var nwState = (InwOpState3)ComApiBridge.State;

                foreach (ModelItem item in selected)
                {
                    try
                    {
                        var nwPath = (InwOaPath3)ComApiBridge.ToInwOaPath(item);
                        var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);

                        if (isExcluir)
                        {
                            // Excluir: localizar índice da categoria e remover
                            var guiAttribs = propNode.GUIAttributes();
                            int udCount = 0;
                            int targetNdx = -1;
                            foreach (object obj in guiAttribs)
                            {
                                var a = obj as InwGUIAttribute2;
                                if (a == null || !a.UserDefined) continue;
                                if (string.Equals(a.name, categoryName, StringComparison.OrdinalIgnoreCase))
                                {
                                    targetNdx = udCount;
                                    break;
                                }
                                udCount++;
                            }

                            if (targetNdx >= 0)
                            {
                                // Gravar vetor vazio para apagar a categoria
                                var emptyVec = (InwOaPropertyVec)nwState.ObjectFactory(
                                    nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                                propNode.SetUserDefined(targetNdx, categoryName, categoryName, emptyVec);
                                written++;
                            }
                        }
                        else
                        {
                            // Gravar / Editar — com merge incremental
                            var guiAttribs = propNode.GUIAttributes();
                            int udCount = 0;
                            int targetNdx = -1;
                            foreach (object obj in guiAttribs)
                            {
                                var a = obj as InwGUIAttribute2;
                                if (a == null || !a.UserDefined) continue;
                                udCount++;
                                if (targetNdx < 0 &&
                                    string.Equals(a.name, categoryName, StringComparison.OrdinalIgnoreCase))
                                    targetNdx = udCount - 1;
                            }
                            if (targetNdx < 0) targetNdx = udCount;

                            // Lê atributos existentes e faz merge
                            var existingProps = ReadExistingProps(propNode, categoryName);

                            // Upsert: grid tem prioridade
                            foreach (var entry in _entries)
                            {
                                if (string.IsNullOrWhiteSpace(entry.Name)) continue;
                                existingProps[entry.Name] = ConvertValue(entry);
                            }

                            // Constrói propVec com o resultado completo
                            var propVec = (InwOaPropertyVec)nwState.ObjectFactory(
                                nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                            var propColl = propVec.Properties();

                            foreach (var kvp in existingProps)
                            {
                                var prop = (InwOaProperty)nwState.ObjectFactory(
                                    nwEObjectType.eObjectType_nwOaProperty, null, null);
                                prop.UserName = kvp.Key;
                                prop.value = kvp.Value;
                                propColl.Add(prop);
                            }

                            propNode.SetUserDefined(targetNdx, categoryName, categoryName, propVec);
                            written++;
                        }
                    }
                    catch (Exception itemEx)
                    {
                        errors++;
                        if (firstError == null) firstError = itemEx.Message;
                        System.Diagnostics.Debug.WriteLine($"[WriteAttributes] Erro no item: {itemEx.Message}");
                    }
                }

                string verb = isExcluir ? "Exclusão" : "Gravação";
                string msg = $"{verb} concluída.\n✔ {written} elemento(s) processados.";
                if (errors > 0)
                {
                    msg += $"\n✖ {errors} erro(s).";
                    if (firstError != null) msg += $"\n\nDetalhe: {firstError}";
                }
                MessageBox.Show(msg, verb + " de Atributos", MessageBoxButton.OK,
                    errors > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);

                UpdateStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao processar atributos:\n{ex.Message}", "Erro",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>Converte valor texto para tipo nativo. Aceita: string, int, double, boolean ("true"/"1"/"sim").</summary>
        private static object ConvertValue(AttributeEntry entry)
        {
            string raw = entry.Value ?? "";
            switch (entry.Type)
            {
                case "int":
                    return int.TryParse(raw, out int i) ? (object)i : 0;
                case "double":
                    return double.TryParse(raw,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out double d) ? (object)d : 0.0;
                case "boolean":
                    return raw.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || raw == "1"
                        || raw.Equals("sim", StringComparison.OrdinalIgnoreCase);
                default: return raw;
            }
        }
    }

    // ── MODELOS ─────────────────────────────────────────────────────────────────

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

        public string CategoryName
        {
            get => _categoryName;
            set { _categoryName = value; OnPropertyChanged(nameof(CategoryName)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class CatNode : INotifyPropertyChanged
    {
        private bool _isExpanded = true;

        public string Name { get; set; }

        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(nameof(IsExpanded)); }
        }

        public ObservableCollection<PropNode> Properties { get; set; } = new ObservableCollection<PropNode>();

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PropNode : INotifyPropertyChanged
    {
        private bool _isChecked;

        public string Name { get; set; }
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

}
