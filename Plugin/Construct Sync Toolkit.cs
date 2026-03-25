using System;
using System.Windows;
using Autodesk.Navisworks.Api.Plugins;

namespace TestePlugin
{
    // ══════════════════════════════════════════════════════════════════════
    //  Construct Sync Toolkit.cs — Ponto de entrada do plugin (namespace TestePlugin)
    //
    //  Estrutura:
    //    • RibbonPlugin          — registra a aba e o layout XAML do ribbon
    //    • AboutPlugin           — abre AboutWindow
    //    • ReadAttributesPlugin  — abre ReadAttributesWindow
    //    • WriteAttributesPlugin — abre WriteAttributesWindow
    //    • ColorizerPlugin       — abre ColorizerWindow
    //    • PluginHelpers         — OpenWindow<T>() compartilhado
    //
    //  IDs alinhados entre [Plugin("id")], ribbon XAML e .addin:
    //    TestePlugin.RibbonPlugin | TestePlugin.About
    //    TestePlugin.ReadAttributes | TestePlugin.WriteAttributes | TestePlugin.Colorizer
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Plugin que abre a janela <see cref="AboutWindow"/>,
    /// exibindo a descrição e os recursos do Construct Sync Toolkit.
    /// </summary>
    [Plugin("TestePlugin.About", "JCA", DisplayName = "Sobre esta aplicativo")]
    public class AboutPlugin : AddInPlugin
    {
        /// <inheritdoc/>
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<AboutWindow>();
            return 0;
        }
    }

    /// <summary>
    /// Plugin que abre a janela <see cref="ReadAttributesWindow"/>,
    /// permitindo inspecionar atributos BIM dos elementos selecionados.
    /// </summary>
    [Plugin("TestePlugin.ReadAttributes", "JCA", DisplayName = "Ler Atributos")]
    public class ReadAttributesPlugin : AddInPlugin
    {
        /// <inheritdoc/>
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<ReadAttributesWindow>();
            return 0;
        }
    }

    /// <summary>
    /// Plugin que abre a janela <see cref="WriteAttributesWindow"/>,
    /// permitindo criar ou atualizar atributos personalizados nos
    /// elementos selecionados.
    /// </summary>
    [Plugin("TestePlugin.WriteAttributes", "JCA", DisplayName = "Atributos")]
    public class WriteAttributesPlugin : AddInPlugin
    {
        /// <inheritdoc/>
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<WriteAttributesWindow>();
            return 0;
        }
    }

    /// <summary>
    /// Plugin que abre a janela <see cref="ColorizerWindow"/>,
    /// permitindo aplicar cores a elementos por Selection Sets ou Search Sets.
    /// </summary>
    [Plugin("TestePlugin.Colorizer", "JCA", DisplayName = "Codificação Visual")]
    public class ColorizerPlugin : AddInPlugin
    {
        /// <inheritdoc/>
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<ColorizerWindow>();
            return 0;
        }
    }

    // ══════════════════════════════════════════════════════════════════════
    //  Utilitários compartilhados pelos plugins
    // ══════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Métodos de extensão e helpers internos usados pelos plugins.
    /// </summary>
    internal static class PluginHelpers
    {
        /// <summary>
        /// Instancia e exibe uma janela WPF do tipo <typeparamref name="T"/>.
        /// Em caso de exceção, exibe uma MessageBox com a mensagem de erro.
        /// </summary>
        /// <typeparam name="T">
        /// Tipo da janela WPF a ser criada; deve ter um construtor sem parâmetros.
        /// </typeparam>
        internal static void OpenWindow<T>() where T : Window, new()
        {
            try
            {
                new T().ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Erro ao abrir a janela {typeof(T).Name}:\n{ex.Message}",
                    "Construct Sync",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}
