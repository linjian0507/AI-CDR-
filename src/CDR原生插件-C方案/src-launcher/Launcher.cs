using System;
using System.IO;
using System.Runtime.InteropServices;
using CdrApp = Corel.Interop.VGCore.Application;
using RGiesecke.DllExport;

namespace AIVectorLauncher
{
    // CDR 原生插件入口。编译产物改名为 .cpg 放进 addon 文件夹后, CDR 启动会加载它并调用 OnLoad,
    // 把[进程内] Application 递进来(不依赖绿色版缺失的 COM 运行对象表注册)。
    // 存进 AppDomain 进程级数据, 供 wpfhost 界面程序集 AIVectorHelper.dll 读取。
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class Plugin : Corel.Interop.VGCore.IVGAppPlugin
    {
        public void OnLoad(CdrApp app)
        {
            try
            {
                AppDomain.CurrentDomain.SetData("AIVectorHelper.CorelApp", app);
                int v = 0; try { v = app.VersionMajor; } catch { }
                Log("Launcher.OnLoad 触发 ✓ 已注入 app, 内部版本号=" + v);
            }
            catch (Exception e) { Log("Launcher.OnLoad 出错: " + e.Message); }
        }
        public void StartSession() { Log("Launcher.StartSession"); }
        public void StopSession() { Log("Launcher.StopSession"); }
        public void OnUnload() { Log("Launcher.OnUnload"); AppDomain.CurrentDomain.SetData("AIVectorHelper.CorelApp", null); }

        static void Log(string m)
        {
            try
            {
                File.AppendAllText(Path.Combine(Path.GetTempPath(), "AIVectorHelper.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + m + Environment.NewLine);
            }
            catch { }
        }
    }

    // .cpg 加载器入口。CorelDRAW 通过这个类把 IVGAppPlugin 的 COM 接口指针交给宿主。
    // 该结构与现有可工作的 QiukuLauncher.cpg 保持一致：成功返回 256，失败返回 0。
    public class CpgEntry
    {
        [DllExport("AttachPlugin", CallingConvention = CallingConvention.StdCall)]
        public static uint AttachPlugin(ref IntPtr pluginPtr)
        {
            try
            {
                var plugin = new Plugin();
                pluginPtr = Marshal.GetComInterfaceForObject(
                    plugin,
                    typeof(Corel.Interop.VGCore.IVGAppPlugin));
                return 256;
            }
            catch
            {
                pluginPtr = IntPtr.Zero;
                return 0;
            }
        }

    }
}
