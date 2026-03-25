// ══════════════════════════════════════════════════════════════════════
//  AboutWindow.xaml.cs — Construct Sync Toolkit
//
//  Code-behind da janela "Sobre". Não contém lógica de negócio;
//  apenas fecha a janela quando o botão Fechar é acionado.
// ══════════════════════════════════════════════════════════════════════

using System.Windows;

namespace TestePlugin
{
    /// <summary>
    /// Janela "Sobre" do Construct Sync Toolkit.
    /// Exibe a descrição do plugin e suas três ferramentas principais:
    /// Leitura de Atributos, Gravação de Atributos e Colorizer.
    /// </summary>
    public partial class AboutWindow : Window
    {
        /// <summary>
        /// Inicializa os componentes gerados pelo compilador XAML.
        /// </summary>
        public AboutWindow()
        {
            InitializeComponent();
            Loaded += (s, e) =>
            {
                Activate();
                Topmost = true;
                Topmost = false;
            };
        }

        // ── HANDLERS ───────────────────────────────────────────────────

        /// <summary>
        /// Fecha a janela ao clicar em "Fechar".
        /// </summary>
        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}
