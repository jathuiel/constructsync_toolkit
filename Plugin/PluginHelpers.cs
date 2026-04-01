using System;
using System.Windows;

namespace SetAtributesToolkit
{
    /// <summary>
    /// Abre janelas WPF dentro do host Win32 do Navisworks.
    /// Garante que o esquema pack:// esteja registrado antes de criar qualquer
    /// janela que referencie recursos XAML por URI absoluta.
    /// </summary>
    internal static class PluginHelpers
    {
        internal static void OpenWindow<T>() where T : Window, new()
        {
            try
            {
                if (!UriParser.IsKnownScheme("pack"))
                    new Application();

                var window = new T();
                window.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.ToString());
                MessageBox.Show(
                    $"Erro ao abrir a janela {typeof(T).Name}:\n{ex.Message}\n\n{ex.InnerException?.Message}",
                    "Construct Sync",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}
