using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using NwApp = Autodesk.Navisworks.Api.Application;

namespace SetAtributesToolkit
{
    public partial class NativeAttrLabWindow : Window
    {
        private readonly ObservableCollection<LabCatNode> _nativeCats = new ObservableCollection<LabCatNode>();
        private readonly ObservableCollection<AttributeEntry> _labEntries = new ObservableCollection<AttributeEntry>();
        private List<ModelItem> _selectedItems = new List<ModelItem>();

        // Mapa completo de categorias nativas: nome → (propName → value)
        private Dictionary<string, Dictionary<string, string>> _nativeCatMap
            = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        public NativeAttrLabWindow()
        {
            InitializeComponent();
            NativeTree.ItemsSource = _nativeCats;
            LabGrid.ItemsSource = _labEntries;

            Loaded += async (s, e) =>
            {
                await ReloadNativeAsync();
                Activate();
                Topmost = true;
                Topmost = false;
            };
        }

        // ── CARREGAMENTO ────────────────────────────────────────────────────────

        private async Task ReloadNativeAsync()
        {
            _nativeCats.Clear();
            _nativeCatMap.Clear();
            NativeCategoryCombo.Items.Clear();
            FooterStatus.Text = "Carregando...";

            var doc = NwApp.ActiveDocument;
            if (doc == null || doc.CurrentSelection.IsEmpty)
            {
                NativeSubtitle.Text = "(nenhum elemento selecionado)";
                FooterStatus.Text = "Nenhum elemento selecionado no modelo.";
                return;
            }

            var items = doc.CurrentSelection.SelectedItems.Cast<ModelItem>().ToList();
            _selectedItems = items;

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
                                map[cat.DisplayName][prop.DisplayName] = SafeValue(prop.Value);
                        }
                    }
                }
                return map;
            });

            _nativeCatMap = catMap;

            foreach (var catKv in catMap.OrderBy(c => c.Key))
            {
                var node = new LabCatNode { Name = catKv.Key };
                foreach (var propKv in catKv.Value.OrderBy(p => p.Key))
                    node.Properties.Add(new LabPropNode { Name = propKv.Key, Value = propKv.Value });
                _nativeCats.Add(node);

                NativeCategoryCombo.Items.Add(catKv.Key);
            }

            NativeSubtitle.Text = $"{catMap.Count} categorias • {items.Count} elemento(s)";
            FooterStatus.Text = $"{items.Count} elemento(s) selecionado(s) — {catMap.Count} categorias nativas encontradas";
        }

        private void BtnReloadNative_Click(object sender, RoutedEventArgs e)
            => _ = ReloadNativeAsync();

        // ── IMPORTAR CATEGORIA NATIVA ────────────────────────────────────────────

        private void BtnImportNative_Click(object sender, RoutedEventArgs e)
        {
            var selected = NativeCategoryCombo.SelectedItem as string;
            if (selected == null)
            {
                AppendLog("⚠ Selecione uma categoria nativa no ComboBox antes de importar.");
                return;
            }

            if (!_nativeCatMap.TryGetValue(selected, out var props))
            {
                AppendLog($"⚠ Categoria '{selected}' não encontrada no mapa nativo.");
                return;
            }

            // Preenche o campo de nome da categoria custom com o mesmo nome da nativa
            CustomCategoryNameBox.Text = selected;

            // Limpa e repopula a grade
            _labEntries.Clear();
            foreach (var kvp in props.OrderBy(p => p.Key))
            {
                _labEntries.Add(new AttributeEntry
                {
                    Name = kvp.Key,
                    Value = kvp.Value,
                    Type = "string"
                });
            }

            AppendLog($"✔ {props.Count} propriedade(s) importada(s) da categoria nativa \"{selected}\".");
            AppendLog($"  → Nome da categoria custom definido como: \"{selected}\" (você pode alterar acima).");
        }

        // ── GRADE ────────────────────────────────────────────────────────────────

        private void LabGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
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

        private void BtnAddLabRow_Click(object sender, RoutedEventArgs e)
            => _labEntries.Add(new AttributeEntry { Name = "", Value = "", Type = "string" });

        private void BtnRemoveLabRow_Click(object sender, RoutedEventArgs e)
        {
            if (((Button)sender).Tag is AttributeEntry entry)
                _labEntries.Remove(entry);
        }

        // ── ESCREVER E VERIFICAR ─────────────────────────────────────────────────

        private void BtnWriteAndVerify_Click(object sender, RoutedEventArgs e)
        {
            LabGrid.CommitEdit(DataGridEditingUnit.Row, true);

            string categoryName = CustomCategoryNameBox.Text?.Trim();
            if (string.IsNullOrEmpty(categoryName))
            {
                MessageBox.Show("Informe o nome da categoria custom (campo 2).", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_labEntries.Count == 0)
            {
                MessageBox.Show("Adicione ao menos um atributo na grade (seção 3).", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var doc = NwApp.ActiveDocument;
            if (doc == null || doc.CurrentSelection.IsEmpty)
            {
                MessageBox.Show("Nenhum elemento selecionado no modelo.", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool isNativeName = _nativeCatMap.ContainsKey(categoryName);
            AppendLog("─────────────────────────────────────────────");
            AppendLog($"▶ Início do teste — categoria: \"{categoryName}\"");
            AppendLog($"  Categoria existe nos atributos nativos? {(isNativeName ? "SIM ⚠" : "NÃO")}");
            AppendLog($"  Atributos na grade (novos/editados): {_labEntries.Count}");
            AppendLog($"  Modo: merge — props existentes na categoria serão preservadas");

            int written = 0, errors = 0;

            try
            {
                var nwState = (InwOpState3)ComApiBridge.State;

                foreach (ModelItem item in doc.CurrentSelection.SelectedItems)
                {
                    try
                    {
                        var nwPath = (InwOaPath3)ComApiBridge.ToInwOaPath(item);
                        var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);

                        // Localiza índice da categoria UserDefined alvo
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

                        // Lê props UserDefined já existentes nessa categoria e faz merge
                        // (propriedades da grade têm prioridade; as demais são preservadas)
                        var existingProps = ReadExistingUserDefinedProps(propNode, categoryName);
                        foreach (var entry in _labEntries)
                        {
                            if (!string.IsNullOrWhiteSpace(entry.Name))
                                existingProps[entry.Name] = ConvertLabValue(entry);
                        }

                        // Constrói o vetor com o resultado merged
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
                    catch (Exception itemEx)
                    {
                        errors++;
                        AppendLog($"  ✖ Erro em elemento: {itemEx.Message}");
                    }
                }

                AppendLog($"  Gravação concluída: {written} elemento(s) processado(s), {errors} erro(s).");
            }
            catch (Exception ex)
            {
                AppendLog($"✖ Erro geral na gravação: {ex.Message}");
                return;
            }

            // Verificação: reler todos os atributos e checar coexistência
            AppendLog("  Verificando atributos após gravação...");
            VerifyAfterWrite(categoryName);
        }

        private void VerifyAfterWrite(string writtenCategoryName)
        {
            try
            {
                var doc = NwApp.ActiveDocument;
                if (doc == null || doc.CurrentSelection.IsEmpty) return;

                // Reler via Managed API para ver o que ficou gravado
                var firstItem = doc.CurrentSelection.SelectedItems.Cast<ModelItem>().First();

                var foundNative = false;
                var foundCustom = false;
                var sb = new StringBuilder();

                sb.AppendLine("  ── Categorias encontradas no elemento após gravação ──");

                foreach (PropertyCategory cat in firstItem.PropertyCategories)
                {
                    bool isTargetCat = string.Equals(cat.DisplayName, writtenCategoryName,
                                                     StringComparison.OrdinalIgnoreCase);

                    string tag = isTargetCat ? $" ◄ \"{writtenCategoryName}\"" : "";
                    sb.AppendLine($"  [{cat.DisplayName}]{tag}");

                    if (!isTargetCat) continue;

                    // Detalha as propriedades da categoria alvo
                    foreach (DataProperty prop in cat.Properties)
                        sb.AppendLine($"      • {prop.DisplayName} = {SafeValue(prop.Value)}");
                }

                AppendLog(sb.ToString().TrimEnd());

                // Verificação via COM para distinguir nativo vs UserDefined
                AppendLog("  ── Verificação via COM API (UserDefined flag) ──");
                var nwState = (InwOpState3)ComApiBridge.State;
                var nwPath = (InwOaPath3)ComApiBridge.ToInwOaPath(firstItem);
                var propNode = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);

                foreach (object obj in propNode.GUIAttributes())
                {
                    var a = obj as InwGUIAttribute2;
                    if (a == null) continue;

                    bool isTarget = string.Equals(a.name, writtenCategoryName,
                                                  StringComparison.OrdinalIgnoreCase);
                    if (!isTarget) continue;

                    string tipo = a.UserDefined ? "[UserDefined]" : "[Nativo]";
                    AppendLog($"  {tipo}  categoria: \"{a.name}\"");

                    foreach (object pObj in a.Properties())
                    {
                        var p = pObj as InwOaProperty;
                        if (p != null)
                            AppendLog($"      • {p.UserName} = {p.value}");
                    }

                    if (a.UserDefined) foundCustom = true;
                    else foundNative = true;
                }

                AppendLog("  ── Conclusão ──");
                if (foundNative && foundCustom)
                    AppendLog($"  ✔ COEXISTÊNCIA: categoria \"{writtenCategoryName}\" existe como NATIVA e como USERDEFINED simultaneamente.");
                else if (foundCustom && !foundNative)
                    AppendLog($"  ✔ Categoria \"{writtenCategoryName}\" criada apenas como UserDefined (sem nativa com mesmo nome).");
                else if (foundNative && !foundCustom)
                    AppendLog($"  ⚠ Apenas a versão NATIVA de \"{writtenCategoryName}\" foi encontrada via COM. A gravação pode ter falhado silenciosamente.");
                else
                    AppendLog($"  ⚠ Categoria \"{writtenCategoryName}\" não encontrada via COM após gravação.");

                // Atualiza a árvore nativa com os dados mais recentes
                _ = ReloadNativeAsync();
            }
            catch (Exception ex)
            {
                AppendLog($"  ✖ Erro na verificação: {ex.Message}");
            }
        }

        // ── HELPERS DE LEITURA ───────────────────────────────────────────────────

        /// <summary>
        /// Lê as propriedades já gravadas como UserDefined para a categoria informada.
        /// Retorna dicionário vazio se a categoria ainda não existir.
        /// </summary>
        private static Dictionary<string, object> ReadExistingUserDefinedProps(
            InwGUIPropertyNode2 propNode, string categoryName)
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

        // ── LOG ──────────────────────────────────────────────────────────────────

        private void AppendLog(string line)
        {
            LogText.Text += line + "\n";
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
            => LogText.Text = string.Empty;

        // ── HELPERS ──────────────────────────────────────────────────────────────

        private static string SafeValue(VariantData v)
        {
            if (v == null || v.DataType == VariantDataType.None) return string.Empty;
            try { return v.ToDisplayString() ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static object ConvertLabValue(AttributeEntry entry)
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

    // ── MODELOS DO LAB ───────────────────────────────────────────────────────────

    public class LabCatNode : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public ObservableCollection<LabPropNode> Properties { get; set; } = new ObservableCollection<LabPropNode>();

        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class LabPropNode
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
}
