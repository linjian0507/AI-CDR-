using System;
using System.IO;

namespace AIVectorHelper
{
    // 简易文件日志: 用于在无界面情况下自动验证插件被 CDR 加载到哪一步。
    public static class Log
    {
        private static readonly string PathFile =
            Path.Combine(Path.GetTempPath(), "AIVectorHelper.log");

        public static void W(string msg)
        {
            try
            {
                File.AppendAllText(PathFile,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + msg + Environment.NewLine);
            }
            catch { }
        }
    }
}
