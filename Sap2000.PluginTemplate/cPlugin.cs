using SAP2000v20;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Sap2000.PluginTemplate
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class cPlugin
    {
        public void main(ref cSapModel sapModel, ref cPluginCallback pluginCallback)
        {
            MessageBox.Show("ok");
            pluginCallback.Finish(0);
        }
        public int Info(ref string text)
        {
            text += "ZY.FrameLoadHelper v1.0.0.0\n";
            text += "Copyright(C) ShangHai ZhanYun,Inc,2019\n";
            text += "..........................................\n";
            return 0;
        }
    }
}

