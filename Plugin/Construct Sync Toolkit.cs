using Autodesk.Navisworks.Api.Plugins;

namespace SetAtributesToolkit
{
    [Plugin("SetAtributesToolkit.WriteAttributes", "JCA", DisplayName = "Atributos")]
    public class WriteAttributesPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<WriteAttributesWindow>();
            return 0;
        }
    }

    [Plugin("SetAtributesToolkit.SelectionInspector", "JCA", DisplayName = "Selection Inspector")]
    public class SelectionInspectorPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<SelectionInspectorWindow>();
            return 0;
        }
    }

    [Plugin("SetAtributesToolkit.NativeAttrLab", "JCA", DisplayName = "Native Attr Lab")]
    public class NativeAttrLabPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            PluginHelpers.OpenWindow<NativeAttrLabWindow>();
            return 0;
        }
    }
}
