using CdrApp = Corel.Interop.VGCore.Application;

namespace AIVectorHelper
{
    // 全局持有 CorelDRAW 进程内 Application 对象(由插件入口 OnLoad 注入)。
    // 进程内直接持有真对象 => 不经过坏掉的进程外 COM, 导入/描摹/关文档都安全。
    public static class PluginHost
    {
        public static CdrApp App;
    }

    // CorelDRAW 插件入口。若 CDR 通过 IVGAppPlugin 契约加载本类, 会在启动时调用 OnLoad,
    // 把进程内 app 对象递进来。(契约来自 Corel.Interop.VGCore.IVGAppPlugin, CDR2018 = 20.1.0.708)
    public class AIPlugin : Corel.Interop.VGCore.IVGAppPlugin
    {
        public void OnLoad(CdrApp Application)
        {
            PluginHost.App = Application;
            int ver = 0;
            try { ver = Application.VersionMajor; } catch { }
            Log.W("IVGAppPlugin.OnLoad 已触发 ✓  内部版本号=" + ver);
        }
        public void StartSession() { Log.W("StartSession"); }
        public void StopSession() { Log.W("StopSession"); }
        public void OnUnload() { Log.W("OnUnload"); PluginHost.App = null; }
    }
}
