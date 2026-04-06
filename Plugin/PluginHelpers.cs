using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using System;
using System.Collections.Generic;
using System.Windows;

namespace SetAtributesToolkit
{
    internal static class PluginHelpers
    {
        // ── JANELAS ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Abre janelas WPF dentro do host Win32 do Navisworks.
        /// Garante que o esquema pack:// esteja registrado antes de criar qualquer
        /// janela que referencie recursos XAML por URI absoluta.
        /// </summary>
        internal static void OpenWindow<T>() where T : Window, new()
        {
            try
            {
                if (!UriParser.IsKnownScheme("pack"))
                    new System.Windows.Application();

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

        // ── API NAVISWORKS: leitura ───────────────────────────────────────────────

        /// <summary>Converte VariantData para string de forma segura.</summary>
        internal static string SafeValue(VariantData v)
        {
            if (v == null || v.DataType == VariantDataType.None) return string.Empty;
            try { return v.ToDisplayString() ?? string.Empty; }
            catch { return string.Empty; }
        }

        // ── API NAVISWORKS: gravação ──────────────────────────────────────────────

        /// <summary>
        /// Grava (ou atualiza via merge) uma categoria UserDefined em um elemento.
        /// Faz uma única iteração sobre GUIAttributes() para localizar o índice e
        /// ler as props existentes, evitando dois calls COM separados.
        /// As props do dicionário <paramref name="newProps"/> têm prioridade (upsert).
        /// </summary>
        internal static void WriteUserDefinedCategory(
            InwOpState3 nwState,
            InwGUIPropertyNode2 propNode,
            string categoryName,
            IEnumerable<KeyValuePair<string, object>> newProps)
        {
            int udCount = 0, targetNdx = -1;
            var merged = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            foreach (object obj in propNode.GUIAttributes())
            {
                var a = obj as InwGUIAttribute2;
                if (a == null || !a.UserDefined) continue;

                if (string.Equals(a.name, categoryName, StringComparison.OrdinalIgnoreCase))
                {
                    targetNdx = udCount;
                    foreach (object pObj in a.Properties())
                    {
                        var p = pObj as InwOaProperty;
                        if (p != null && !string.IsNullOrWhiteSpace(p.UserName))
                            merged[p.UserName] = p.value;
                    }
                }
                udCount++;
            }
            if (targetNdx < 0) targetNdx = udCount;

            // Upsert: novas props sobrescrevem as existentes
            foreach (var kvp in newProps)
                merged[kvp.Key] = kvp.Value;

            var propVec  = (InwOaPropertyVec)nwState.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
            var propColl = propVec.Properties();

            foreach (var kvp in merged)
            {
                var prop = (InwOaProperty)nwState.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
                prop.UserName = kvp.Key;
                prop.value    = kvp.Value;
                propColl.Add(prop);
            }

            propNode.SetUserDefined(targetNdx, categoryName, categoryName, propVec);
        }

        /// <summary>
        /// Exclui uma categoria UserDefined de um elemento gravando um vetor vazio.
        /// Não faz nada se a categoria não existir.
        /// </summary>
        internal static void DeleteUserDefinedCategory(
            InwOpState3 nwState,
            InwGUIPropertyNode2 propNode,
            string categoryName)
        {
            int udCount = 0, targetNdx = -1;

            foreach (object obj in propNode.GUIAttributes())
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

            if (targetNdx < 0) return;

            var emptyVec = (InwOaPropertyVec)nwState.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
            propNode.SetUserDefined(targetNdx, categoryName, categoryName, emptyVec);
        }

        // ── CONVERSÃO DE TIPOS ────────────────────────────────────────────────────

        /// <summary>
        /// Converte o valor texto de um AttributeEntry para o tipo nativo correspondente.
        /// Suporta: string, int, double, boolean ("true" / "1" / "sim").
        /// </summary>
        internal static object ConvertValue(AttributeEntry entry)
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

        // ── SERIALIZAÇÃO ──────────────────────────────────────────────────────────

        /// <summary>Escapa um campo para CSV (separador ponto-e-vírgula).</summary>
        internal static string EscapeCsv(string s)
        {
            if (s == null) return "";
            return (s.IndexOf(';') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0)
                ? $"\"{s.Replace("\"", "\"\"")}\""
                : s;
        }

        /// <summary>Escapa um valor para uso em string JSON.</summary>
        internal static string EscapeJson(string s)
            => s?.Replace("\\", "\\\\").Replace("\"", "\\\"") ?? "";

        /// <summary>Extrai o valor de uma chave em um bloco JSON simples (sem Newtonsoft).</summary>
        internal static string ExtractJsonValue(string json, string key)
        {
            string search = $"\"{key}\":\"";
            int start = json.IndexOf(search, StringComparison.OrdinalIgnoreCase);
            if (start < 0) return null;
            start += search.Length;
            int end = json.IndexOf('"', start);
            return end < 0 ? null : json.Substring(start, end - start);
        }
    }
}
