using System;
using System.Collections.Generic;
using System.Text;

namespace DrugFlow
{
    public static class Report
    {
        public static string Build(ScanResult r)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("==================================================================");
            sb.AppendLine("        药房流水导出助手 · 环境诊断报告");
            sb.AppendLine("==================================================================");
            sb.AppendLine("生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            sb.AppendLine("程序版本：" + AppInfo.Version);
            sb.AppendLine();
            sb.AppendLine("【一、系统信息】");
            foreach (string n in r.Notes) sb.AppendLine("  " + n);
            sb.AppendLine();

            sb.AppendLine("【二、可以直接使用的数据源】 " + r.Sources.Count + " 个");
            foreach (SourceItem s in r.Sources)
            {
                sb.AppendLine("  * [" + s.Kind + "] " + s.Title);
                if (!string.IsNullOrEmpty(s.Connection)) sb.AppendLine("      连接：" + Tx.MaskSecret(s.Connection));
            }
            if (r.Sources.Count <= 1)
            {
                sb.AppendLine("  （除示例数据外没有发现可用数据源，请看下面的线索和表结构）");
            }
            sb.AppendLine();

            sb.AppendLine("【三、扫描线索】 " + r.Findings.Count + " 条");
            string lastKind = null;
            foreach (Finding f in r.Findings)
            {
                if (f.Kind != lastKind)
                {
                    sb.AppendLine();
                    sb.AppendLine("  -- " + f.Kind + " --");
                    lastKind = f.Kind;
                }
                sb.AppendLine("  · " + f.Title);
                if (!string.IsNullOrEmpty(f.Detail)) sb.AppendLine("      " + f.Detail);
                if (!string.IsNullOrEmpty(f.Path) && f.Path != f.Detail) sb.AppendLine("      位置：" + f.Path);
            }
            sb.AppendLine();

            sb.AppendLine("【四、数据库表结构（按“像不像销售流水”排序；只有表名和字段名，没有数据）】");
            if (r.Schema.Count == 0) sb.AppendLine("  （没有找到可以直接读取的数据库）");
            foreach (string line in r.Schema) sb.AppendLine("  " + line);
            sb.AppendLine();

            sb.AppendLine("【五、疑似安装目录】");
            foreach (string d in r.CandidateDirs) sb.AppendLine("  " + d);
            sb.AppendLine();

            sb.AppendLine("【六、进程列表】");
            sb.AppendLine("  （为减少杀毒软件误报，本程序不再读取其它进程的信息）");
            sb.AppendLine();

            sb.AppendLine("【七、给用户的话】");
            sb.AppendLine("  1. 请把本文件原样发给帮你做工具的人（微信发送即可）。");
            sb.AppendLine("  2. 报告里没有药品销售的具体数据，只有软件名字、路径和数据库位置。");
            sb.AppendLine("  3. 如果“可以直接使用的数据源”只有示例数据，说明这个软件的数据不是常见的");
            sb.AppendLine("     SQL Server / Access / SQLite / DBF 格式，需要技术人员根据上面的线索再适配。");
            sb.AppendLine();
            return sb.ToString();
        }

        public static string Save(ScanResult r)
        {
            string name = "药房流水诊断报告_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt";
            string text = Build(r);
            string[] dirs = new string[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                AppDomain.CurrentDomain.BaseDirectory,
                System.IO.Path.GetTempPath()
            };
            Exception last = null;
            for (int i = 0; i < dirs.Length; i++)
            {
                try
                {
                    if (string.IsNullOrEmpty(dirs[i])) continue;
                    string path = System.IO.Path.Combine(dirs[i], name);
                    System.IO.File.WriteAllText(path, text, Encoding.UTF8);
                    return path;
                }
                catch (Exception ex) { last = ex; }
            }
            throw new Exception("无法保存报告文件：" + (last == null ? "未知原因" : last.Message));
        }
    }

    public static class AppInfo
    {
        public const string Name = "药房流水导出助手";
        public const string Version = "1.0";
    }
}
