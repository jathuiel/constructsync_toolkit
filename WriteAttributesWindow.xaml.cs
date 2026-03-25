// ══════════════════════════════════════════════════════════════════════
//  WriteAttributesWindow.xaml.cs
//
//  Responsabilidade: criar ou atualizar atributos personalizados em
//  elementos BIM selecionados no Navisworks via COM API.
//
//  Fluxo principal (BtnWrite_Click):
//    1. Valida entradas (nome da categoria e lista de atributos).
//    2. Obtém os elementos selecionados no documento ativo.
//    3. Para cada elemento, localiza ou cria a categoria InwOaPropertyAttribute.
//    4. Para cada atributo, localiza ou cria a propriedade InwOaProperty.
//    5. Converte o valor de string para o tipo nativo (ConvertValue).
//
//  Tipos suportados na coluna "Tipo": string | int | double | boolean.
// ══════════════════════════════════════════════════════════════════════

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;           // ComApiBridge — Autodesk.Navisworks.ComApi.dll
using Autodesk.Navisworks.Api.Interop.ComApi;   // InwOpState3, InwOaPath3, InwGUIPropertyNode2, InwGUIAttribute2, InwOaPropertyVec, InwOaProperty, nwEObjectType
using NwApp = Autodesk.Navisworks.Api.Application;

namespace TestePlugin
{
    /// <summary>
    /// Janela para criar ou sobrescrever atributos personalizados
    /// (categorias e propriedades) em elementos BIM selecionados.
    /// </summary>
    public partial class WriteAttributesWindow : Window
    {
        private readonly ObservableCollection<AttributeEntry> _entries =
            new ObservableCollection<AttributeEntry>();

        /// <summary>
        /// Inicializa a janela, popula a grade com uma linha de exemplo
        /// e atualiza o status com a contagem de elementos selecionados.
        /// </summary>
        public WriteAttributesWindow()
        {
            InitializeComponent();
            AttributesGrid.ItemsSource = _entries;
            _entries.Add(new AttributeEntry { Name = "Exemplo", Value = "Valor", Type = "string" });
            Loaded += (s, e) =>
            {
                UpdateStatus();
                // Força a janela a receber foco do sistema operacional (necessário no host Win32 do Navisworks)
                Activate();
                Topmost = true;
                Topmost = false;
                Keyboard.Focus(CategoryNameBox);
            };
        }

        /// <summary>
        /// Atualiza <c>StatusText</c> e o conteúdo do botão Gravar
        /// com a quantidade de elementos atualmente selecionados no modelo.
        /// </summary>
        private void UpdateStatus()
        {
            try
            {
                int count = NwApp.ActiveDocument?.CurrentSelection.SelectedItems.Count ?? 0;
                StatusText.Text = $"{count} elemento(s) selecionado(s) serão afetados.";
                BtnWrite.Content = $"Gravar em {count} Elemento(s) Selecionado(s)";
            }
            catch { StatusText.Text = "Não foi possível ler a seleção."; }
        }

        /// <summary>
        /// Inicia a edição ao clicar com um único clique na célula,
        /// eliminando a necessidade de duplo-clique para entrar em modo de edição.
        /// </summary>
        private void AttributesGrid_BeginningEdit(object sender, System.Windows.Controls.DataGridBeginningEditEventArgs e)
        {
            // Permite entrar em modo de edição com um único clique
        }

        /// <summary>
        /// Garante foco de teclado no TextBox de edição da célula via Dispatcher.
        /// Necessário porque o Navisworks (host Win32) não repassa mensagens de
        /// teclado automaticamente para o WPF DataGrid.
        /// </summary>
        private void AttributesGrid_PreparingCellForEdit(object sender, System.Windows.Controls.DataGridPreparingCellForEditEventArgs e)
        {
            if (e.EditingElement is System.Windows.Controls.TextBox tb)
            {
                // Dispatcher garante que o foco seja aplicado após o WPF terminar de montar o elemento
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
                {
                    tb.Focus();
                    Keyboard.Focus(tb);
                    tb.SelectAll();
                }));
            }
        }

        /// <summary>Seleciona todo o texto ao entrar em qualquer TextBox editável.</summary>
        private void EditingTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.TextBox tb)
                tb.SelectAll();
        }

        /// <summary>Adiciona uma nova linha vazia à grade de atributos.</summary>
        private void BtnAddRow_Click(object sender, RoutedEventArgs e)
            => _entries.Add(new AttributeEntry { Name = "", Value = "", Type = "string" });

        /// <summary>
        /// Remove a linha cujo botão ✕ foi clicado.
        /// O item é passado via <c>Tag</c> do botão na DataTemplate.
        /// </summary>
        private void BtnRemoveRow_Click(object sender, RoutedEventArgs e)
        {
            if (((System.Windows.Controls.Button)sender).Tag is AttributeEntry entry)
                _entries.Remove(entry);
        }

        /// <summary>
        /// Valida as entradas e grava os atributos via COM API em todos os
        /// elementos selecionados. Exibe um resumo (✔ gravados / ✖ erros) ao final.
        /// </summary>
        private void BtnWrite_Click(object sender, RoutedEventArgs e)
        {
            // Força o commit de qualquer célula ainda em modo de edição.
            // Sem isso, o valor digitado (mas não confirmado com Enter/Tab) é perdido.
            AttributesGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);

            string categoryName = CategoryNameBox.Text?.Trim();
            if (string.IsNullOrEmpty(categoryName))
            {
                MessageBox.Show("Informe o nome da categoria.", "Atenção",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (_entries.Count == 0)
            {
                MessageBox.Show("Adicione ao menos um atributo.", "Atenção",
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
                return;
            }

            int written = 0, errors = 0;
            string firstError = null;

            try
            {
                // State retorna InwOpState10 — compatível com InwOpState3
                var nwState = (InwOpState3)ComApiBridge.State;

                foreach (ModelItem item in selected)
                {
                    try
                    {
                        var nwPath    = (InwOaPath3)ComApiBridge.ToInwOaPath(item);
                        var propNode   = (InwGUIPropertyNode2)nwState.GetGUIPropertyNode(nwPath, false);
                        var guiAttribs = propNode.GUIAttributes();

                        // O ndx de SetUserDefined é posição ENTRE user-defined (não o índice total).
                        // Conta apenas atributos user-defined e verifica se já existe um com o mesmo nome.
                        int udCount   = 0;
                        int targetNdx = -1;
                        foreach (object obj in guiAttribs)
                        {
                            var a = obj as InwGUIAttribute2;
                            if (a == null || !a.UserDefined) continue;
                            udCount++;
                            if (targetNdx < 0 &&
                                string.Equals(a.name, categoryName, StringComparison.OrdinalIgnoreCase))
                                targetNdx = udCount - 1; // 0-based: udCount já incrementou
                        }
                        if (targetNdx < 0)
                            targetNdx = udCount; // 0-based: udCount == próximo índice livre

                        // Monta o vetor de propriedades
                        var propVec  = (InwOaPropertyVec)nwState.ObjectFactory(
                            nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                        var propColl = propVec.Properties();
                        foreach (var entry in _entries)
                        {
                            if (string.IsNullOrWhiteSpace(entry.Name)) continue;

                            var prop = (InwOaProperty)nwState.ObjectFactory(
                                nwEObjectType.eObjectType_nwOaProperty, null, null);
                            prop.UserName = entry.Name;
                            prop.value    = ConvertValue(entry);
                            propColl.Add(prop);
                        }

                        // Grava/atualiza o atributo user-defined via API correta
                        propNode.SetUserDefined(targetNdx, categoryName, categoryName, propVec);
                        written++;
                    }
                    catch (Exception itemEx)
                    {
                        // Elementos sem suporte a atributos COM (ex.: grupos, câmeras)
                        // são contados como erro sem interromper o loop.
                        errors++;
                        if (firstError == null) firstError = itemEx.Message;
                        System.Diagnostics.Debug.WriteLine(
                            $"[WriteAttributes] Erro no item: {itemEx.Message}");
                    }
                }

                string msg = $"Gravação concluída.\n✔ {written} elemento(s) atualizados.";
                if (errors > 0)
                {
                    msg += $"\n✖ {errors} erro(s).";
                    if (firstError != null) msg += $"\n\nDetalhe: {firstError}";
                }
                MessageBox.Show(msg, "Gravar Atributos", MessageBoxButton.OK,
                    errors > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao gravar atributos:\n{ex.Message}", "Erro",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Converte o valor de texto de <paramref name="entry"/> para o tipo
        /// nativo esperado pela COM API do Navisworks.
        /// </summary>
        /// <remarks>
        /// Tipos suportados: <c>string</c> (padrão), <c>int</c>, <c>double</c>, <c>boolean</c>.<br/>
        /// Valores inválidos para int/double retornam 0/0.0.<br/>
        /// Para boolean, aceita "true", "1" e "sim" (case-insensitive).
        /// </remarks>
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

    /// <summary>
    /// Representa uma linha na grade de atributos da janela de gravação.
    /// Implementa <see cref="INotifyPropertyChanged"/> para refletir edições
    /// feitas diretamente na célula do DataGrid.
    /// </summary>
    public class AttributeEntry : INotifyPropertyChanged
    {
        private string _name, _value, _type = "string";

        /// <summary>Nome do atributo (propriedade) a gravar.</summary>
        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }

        /// <summary>Valor do atributo como string; convertido por <c>ConvertValue</c> antes de gravar.</summary>
        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(nameof(Value)); }
        }

        /// <summary>Tipo do dado: "string" | "int" | "double" | "boolean".</summary>
        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(nameof(Type)); }
        }

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
