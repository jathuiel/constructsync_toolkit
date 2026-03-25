using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;           // ComApiBridge
using Autodesk.Navisworks.Api.Interop.ComApi;   // InwOpState3, InwLVec3f, nwEObjectType
using NwApp   = Autodesk.Navisworks.Api.Application;
using NwColor = Autodesk.Navisworks.Api.Color;
using WpfColor = System.Windows.Media.Color;

namespace TestePlugin
{
    // ══════════════════════════════════════════════════════════════════════
    //  ColorizerWindow.xaml.cs
    //
    //  Responsabilidades:
    //    • Listar Selection Sets / Search Sets / todos os elementos do modelo.
    //    • Permitir atribuição de cor individual (color picker) ou em lote
    //      (paleta predefinida / cor aleatória).
    //    • Aplicar as cores via COM API (InwOpState3.OverrideColor).
    //    • Restaurar as cores originais via OverrideResetAll.
    //
    //  Modelo de dados: SetColorEntry (INotifyPropertyChanged) garante que
    //  o swatch de cor na ListView reflita mudanças em tempo real.
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Janela de colorização: aplica cores a elementos agrupados por
    /// Selection Sets ou Search Sets do Navisworks.
    /// </summary>
    public partial class ColorizerWindow : Window
    {
        // ── Constantes de cor padrão ─────────────────────────────────────

        /// <summary>Cor padrão atribuída a cada entrada ao criá-la (azul Catppuccin).</summary>
        private static readonly WpfColor DefaultEntryColor = WpfColor.FromRgb(137, 180, 250);

        // ── Paleta predefinida ───────────────────────────────────────────

        /// <summary>
        /// Paleta de 20 cores vibrantes usada por "Paleta Predefinida".
        /// As cores são aplicadas ciclicamente quando há mais sets do que entradas.
        /// </summary>
        private static readonly WpfColor[] Palette =
        {
            WpfColor.FromRgb(235,  64,  52),  WpfColor.FromRgb(235, 149,  52),
            WpfColor.FromRgb(235, 226,  52),  WpfColor.FromRgb(119, 235,  52),
            WpfColor.FromRgb( 52, 235, 131),  WpfColor.FromRgb( 52, 235, 223),
            WpfColor.FromRgb( 52, 152, 235),  WpfColor.FromRgb( 52,  64, 235),
            WpfColor.FromRgb(149,  52, 235),  WpfColor.FromRgb(235,  52, 208),
            WpfColor.FromRgb(255, 128, 128),  WpfColor.FromRgb(255, 200, 100),
            WpfColor.FromRgb(100, 220, 100),  WpfColor.FromRgb( 80, 200, 200),
            WpfColor.FromRgb(100, 140, 255),  WpfColor.FromRgb(200, 100, 255),
            WpfColor.FromRgb(255, 160,  60),  WpfColor.FromRgb( 60, 180, 120),
            WpfColor.FromRgb(180,  80,  80),  WpfColor.FromRgb( 80, 120, 180),
        };

        // ── Estado interno ───────────────────────────────────────────────

        /// <summary>Coleção observável ligada ao ItemsSource do SetsListView.</summary>
        private readonly ObservableCollection<SetColorEntry> _entries =
            new ObservableCollection<SetColorEntry>();

        /// <summary>Gerador de números aleatórios compartilhado (instância estática é thread-safe para leitura).</summary>
        private static readonly Random _rng = new Random();

        // ── Construtor ───────────────────────────────────────────────────

        /// <summary>
        /// Inicializa a janela, liga a lista de entries ao ListView
        /// e carrega os sets assim que a janela é exibida.
        /// </summary>
        public ColorizerWindow()
        {
            InitializeComponent();
            SetsListView.ItemsSource = _entries;
            Loaded += (s, e) =>
            {
                LoadSets();
                // Força a janela a receber foco do sistema operacional (necessário no host Win32 do Navisworks)
                Activate();
                Topmost = true;
                Topmost = false;
                Keyboard.Focus(SetsListView);
            };
        }

        // ── Carregar sets ────────────────────────────────────────────────

        /// <summary>
        /// Popula <see cref="_entries"/> com os sets visíveis conforme o
        /// radio button selecionado (Selection Sets, Search Sets ou todos).
        /// </summary>
        private void LoadSets()
        {
            _entries.Clear();

            var doc = NwApp.ActiveDocument;
            if (doc == null)
            {
                StatusText.Text = "Nenhum documento ativo.";
                return;
            }

            if (RadioAllElements.IsChecked == true)
            {
                // Modo "Todos os Elementos": uma entrada única representa o modelo inteiro
                _entries.Add(new SetColorEntry
                {
                    SetName   = "[ Todos os Elementos ]",
                    SetType   = "Modelo",
                    WpfColor  = DefaultEntryColor,
                    SavedItem = null   // null = raiz de todos os modelos (ver BtnApplySelected_Click)
                });
            }
            else
            {
                bool showSel    = RadioSelectionSets.IsChecked == true;
                bool showSearch = RadioSearchSets.IsChecked    == true;
                CollectSets(doc.SelectionSets, showSel, showSearch);
            }

            StatusText.Text = $"{_entries.Count} set(s) — duplo-clique para selecionar no modelo.";
        }

        /// <summary>
        /// Percorre recursivamente <paramref name="items"/> e adiciona à lista
        /// os sets do tipo solicitado (seleção explícita e/ou busca).
        /// </summary>
        /// <param name="items">Coleção de SavedItems a inspecionar (pode conter pastas).</param>
        /// <param name="showSel">Se <c>true</c>, inclui Selection Sets explícitos.</param>
        /// <param name="showSearch">Se <c>true</c>, inclui Search Sets.</param>
        private void CollectSets(SavedItemCollection items, bool showSel, bool showSearch)
        {
            foreach (SavedItem item in items)
            {
                if (item is FolderItem folder)
                {
                    // Desce na hierarquia de pastas recursivamente
                    CollectSets(folder.Children, showSel, showSearch);
                }
                else if (item is SelectionSet ss)
                {
                    // HasSearch == true → Search Set; false → Selection Set explícito
                    bool isSearch = ss.HasSearch;
                    if ((isSearch && showSearch) || (!isSearch && showSel))
                        _entries.Add(CreateEntry(ss, isSearch ? "Search" : "Selection"));
                }
            }
        }

        /// <summary>
        /// Cria uma nova <see cref="SetColorEntry"/> com a cor padrão.
        /// </summary>
        private static SetColorEntry CreateEntry(SavedItem item, string type) =>
            new SetColorEntry
            {
                SetName   = item.DisplayName ?? "(sem nome)",
                SetType   = type,
                WpfColor  = DefaultEntryColor,
                SavedItem = item
            };

        // ── Eventos de fonte ─────────────────────────────────────────────

        /// <summary>Recarrega a lista ao mudar o radio button de fonte.</summary>
        private void RadioSource_Checked(object sender, RoutedEventArgs e)
        {
            if (IsLoaded) LoadSets();
        }

        // ── Swatch — clique para color picker ───────────────────────────

        /// <summary>
        /// Abre o <see cref="System.Windows.Forms.ColorDialog"/> ao clicar no
        /// swatch de cor de uma linha; atualiza a cor da entrada se confirmado.
        /// </summary>
        private void Swatch_Click(object sender, MouseButtonEventArgs e)
        {
            if (((FrameworkElement)sender).Tag is SetColorEntry entry)
            {
                var dlg = new System.Windows.Forms.ColorDialog
                {
                    Color    = System.Drawing.Color.FromArgb(entry.WpfColor.R,
                                                             entry.WpfColor.G,
                                                             entry.WpfColor.B),
                    FullOpen = true   // exibe o painel expandido com controles RGB/HSL
                };

                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    entry.WpfColor = WpfColor.FromRgb(dlg.Color.R, dlg.Color.G, dlg.Color.B);
            }

            // Impede que o clique no swatch propague para o item da ListView
            e.Handled = true;
        }

        // ── Duplo-clique — seleciona elementos do set no modelo ──────────

        /// <summary>
        /// Ao dar duplo-clique em um set, substitui a seleção atual do Navisworks
        /// pelos elementos daquele set e exibe a contagem no status.
        /// </summary>
        private void SetsListView_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            var entry = SetsListView.SelectedItem as SetColorEntry;
            if (!(entry?.SavedItem is SelectionSet ss)) return;

            var doc = NwApp.ActiveDocument;
            if (doc == null) return;

            try
            {
                var items = ss.GetSelectedItems(doc);
                doc.CurrentSelection.Clear();
                doc.CurrentSelection.AddRange(items);
                StatusText.Text = $"{items.Count} elemento(s) selecionado(s) — {entry.SetName}";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Erro ao selecionar elementos: {ex.Message}";
            }
        }

        // ── Cores aleatórias ─────────────────────────────────────────────

        /// <summary>
        /// Atribui uma cor aleatória (R, G, B ∈ [60, 239]) a cada entry,
        /// evitando tons muito escuros ou claros demais.
        /// </summary>
        private void BtnRandomColor_Click(object sender, RoutedEventArgs e)
        {
            foreach (var entry in _entries)
                entry.WpfColor = WpfColor.FromRgb(
                    (byte)_rng.Next(60, 240),
                    (byte)_rng.Next(60, 240),
                    (byte)_rng.Next(60, 240));
        }

        // ── Paleta predefinida ───────────────────────────────────────────

        /// <summary>
        /// Aplica as cores da <see cref="Palette"/> ciclicamente a cada entry.
        /// </summary>
        private void BtnApplyPalette_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < _entries.Count; i++)
                _entries[i].WpfColor = Palette[i % Palette.Length];
        }

        // ── Aplicar cores nos sets selecionados (ou todos) ───────────────

        /// <summary>
        /// Aplica as cores definidas nos sets marcados na ListView (ou em todos,
        /// se nenhum estiver selecionado) via COM API do Navisworks.
        /// </summary>
        private void BtnApplySelected_Click(object sender, RoutedEventArgs e)
        {
            var doc = NwApp.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("Nenhum documento ativo.", "Codificação Visual",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Se nenhum item estiver selecionado na ListView, aplica em todos
            var toApply = SetsListView.SelectedItems.Cast<SetColorEntry>().ToList();
            if (toApply.Count == 0)
                toApply = _entries.ToList();

            int applied = 0, skipped = 0;
            var state = (InwOpState3)ComApiBridge.State;

            foreach (var entry in toApply)
            {
                try
                {
                    var items = ResolveItems(doc, entry);
                    if (items == null || items.Count == 0)
                    {
                        skipped++;
                        continue;
                    }

                    ApplyColorCom(state, items, entry.WpfColor);
                    applied++;
                }
                catch (Exception ex)
                {
                    // Elementos sem geometria ou com geometria corrompida são ignorados;
                    // o detalhe do erro é acumulado no status para não bloquear o loop.
                    skipped++;
                    StatusText.Text = $"Item ignorado: {ex.Message}";
                }
            }

            StatusText.Text = $"Cores aplicadas em {applied} set(s){(skipped > 0 ? $", {skipped} ignorado(s)." : ".")}";
        }

        /// <summary>
        /// Retorna os <see cref="ModelItemCollection"/> correspondentes a uma entry.
        /// Para entrada "Todos os Elementos" (<c>SavedItem == null</c>), retorna
        /// a raiz de cada modelo carregado no documento.
        /// </summary>
        private static ModelItemCollection ResolveItems(Document doc, SetColorEntry entry)
        {
            if (entry.SavedItem == null)
            {
                // Entrada especial: representa todos os modelos abertos
                var all = new ModelItemCollection();
                foreach (var model in doc.Models)
                    all.Add(model.RootItem);
                return all;
            }

            if (entry.SavedItem is SelectionSet ss)
                return ss.GetSelectedItems(doc);

            // Tipo de SavedItem não suportado
            return null;
        }

        // ── Restaurar todas as cores ─────────────────────────────────────

        /// <summary>
        /// Remove todas as sobreposições de cor aplicadas pelo Colorizer,
        /// restaurando as cores originais do modelo.
        /// </summary>
        private void BtnRestore_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ((InwOpState3)ComApiBridge.State).OverrideResetAll();
                StatusText.Text = "Todas as cores restauradas.";
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Erro ao restaurar cores: {ex.Message}";
            }
        }

        // ── Helper: aplica cor via COM API ───────────────────────────────

        /// <summary>
        /// Aplica a cor <paramref name="c"/> a todos os elementos em
        /// <paramref name="items"/> usando <c>InwOpState3.OverrideColor</c>.
        /// Os valores RGB são normalizados para o intervalo [0.0, 1.0].
        /// </summary>
        /// <param name="state">Estado COM do Navisworks.</param>
        /// <param name="items">Elementos que receberão a cor.</param>
        /// <param name="c">Cor WPF a aplicar.</param>
        private static void ApplyColorCom(InwOpState3 state, ModelItemCollection items, WpfColor c)
        {
            var sel = ComApiBridge.ToInwOpSelection(items);

            // Cria o vetor de cor RGB normalizado (0.0–1.0) via fábrica de objetos COM
            var colorVec = (InwLVec3f)state.ObjectFactory(
                nwEObjectType.eObjectType_nwLVec3f, null, null);

            colorVec.data1 = c.R / 255.0;   // componente vermelho
            colorVec.data2 = c.G / 255.0;   // componente verde
            colorVec.data3 = c.B / 255.0;   // componente azul

            state.OverrideColor(sel, colorVec);
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  Modelo de dados: SetColorEntry
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Representa uma linha na ListView do Colorizer.
    /// Implementa <see cref="INotifyPropertyChanged"/> para atualizar o swatch
    /// de cor automaticamente quando <see cref="WpfColor"/> é modificado.
    /// </summary>
    public class SetColorEntry : INotifyPropertyChanged
    {
        private WpfColor _wpfColor;

        /// <summary>Nome exibido do set ou grupo.</summary>
        public string    SetName   { get; set; }

        /// <summary>Tipo do set: "Selection", "Search" ou "Modelo".</summary>
        public string    SetType   { get; set; }

        /// <summary>
        /// Referência ao <see cref="SavedItem"/> do Navisworks.
        /// <c>null</c> indica a entrada especial "Todos os Elementos".
        /// </summary>
        public SavedItem SavedItem { get; set; }

        /// <summary>
        /// Cor WPF atual do set. Ao ser alterada, recria <see cref="ColorBrush"/>
        /// e notifica a UI via <see cref="PropertyChanged"/>.
        /// </summary>
        public WpfColor WpfColor
        {
            get => _wpfColor;
            set
            {
                _wpfColor  = value;
                ColorBrush = new SolidColorBrush(value);
                OnPropertyChanged(nameof(WpfColor));
                OnPropertyChanged(nameof(ColorBrush));
            }
        }

        /// <summary>
        /// Brush derivado de <see cref="WpfColor"/>, vinculado ao swatch na ListView.
        /// Recriado automaticamente a cada mudança de cor.
        /// </summary>
        public SolidColorBrush ColorBrush { get; private set; } =
            new SolidColorBrush(WpfColor.FromRgb(137, 180, 250));

        /// <inheritdoc/>
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
