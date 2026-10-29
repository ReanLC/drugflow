using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DrugFlow
{
    /// <summary>
    /// 一个“方案”：数据源 + 数据表 + 字段映射（+ 可选的主表关联），
    /// 可以存成 .ini 发给别人，放到 profiles 目录里就能直接用。
    /// </summary>
    public class Profile
    {
        public string Name = "";
        public string Kind = "";          // sqlserver / access / sqlite / odbc
        public string Connection = "";
        public string Table = "";
        public string Date = "";
        public string Drug = "";
        public string Spec = "";
        public string Qty = "";
        public string Price = "";
        public string Amount = "";
        public string Unit = "";
        public string Bill = "";
        public string DateStyle = "auto";
        public string JoinMaster = "";
        public string JoinDetailKey = "";
        public string JoinMasterKey = "";
        public string JoinDate = "";

        public bool HasJoin
        {
            get
            {
                return !string.IsNullOrEmpty(JoinMaster) && !string.IsNullOrEmpty(JoinDetailKey)
                    && !string.IsNullOrEmpty(JoinMasterKey) && !string.IsNullOrEmpty(JoinDate);
            }
        }

        public JoinSpec ToJoin()
        {
            if (!HasJoin) return null;
            JoinSpec j = new JoinSpec();
            j.MasterTable = JoinMaster;
            j.DetailKey = JoinDetailKey;
            j.MasterKey = JoinMasterKey;
            j.DateColumn = JoinDate;
            return j;
        }

        public void FromJoin(JoinSpec j)
        {
            if (j == null)
            {
                JoinMaster = ""; JoinDetailKey = ""; JoinMasterKey = ""; JoinDate = "";
                return;
            }
            JoinMaster = j.MasterTable;
            JoinDetailKey = j.DetailKey;
            JoinMasterKey = j.MasterKey;
            JoinDate = j.DateColumn;
        }

        // ---------------- 读写 ----------------
        public static Profile Load(string path)
        {
            Profile p = new Profile();
            string[] lines = File.ReadAllLines(path, Encoding.UTF8);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0) continue;
                if (line[0] == '#' || line[0] == ';') continue;
                int eq = line.IndexOf('=');
                if (eq <= 0) continue;
                string k = line.Substring(0, eq).Trim().ToLowerInvariant();
                string v = line.Substring(eq + 1).Trim();
                switch (k)
                {
                    case "name": p.Name = v; break;
                    case "kind": p.Kind = v; break;
                    case "connection": p.Connection = v; break;
                    case "table": p.Table = v; break;
                    case "date": p.Date = v; break;
                    case "drug": p.Drug = v; break;
                    case "spec": p.Spec = v; break;
                    case "qty": p.Qty = v; break;
                    case "price": p.Price = v; break;
                    case "amount": p.Amount = v; break;
                    case "unit": p.Unit = v; break;
                    case "bill": p.Bill = v; break;
                    case "datestyle": p.DateStyle = v; break;
                    case "joinmaster": p.JoinMaster = v; break;
                    case "joindetailkey": p.JoinDetailKey = v; break;
                    case "joinmasterkey": p.JoinMasterKey = v; break;
                    case "joindate": p.JoinDate = v; break;
                }
            }
            if (string.IsNullOrEmpty(p.Name))
            {
                try { p.Name = Path.GetFileNameWithoutExtension(path); } catch { }
            }
            return p;
        }

        public void Save(string path)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("# 药房流水导出助手 方案文件（UTF-8）");
            sb.AppendLine("name=" + Name);
            sb.AppendLine("kind=" + Kind);
            sb.AppendLine("connection=" + Connection);
            sb.AppendLine("table=" + Table);
            sb.AppendLine("date=" + Date);
            sb.AppendLine("drug=" + Drug);
            sb.AppendLine("spec=" + Spec);
            sb.AppendLine("qty=" + Qty);
            sb.AppendLine("price=" + Price);
            sb.AppendLine("amount=" + Amount);
            sb.AppendLine("unit=" + Unit);
            sb.AppendLine("bill=" + Bill);
            sb.AppendLine("dateStyle=" + DateStyle);
            sb.AppendLine("joinMaster=" + JoinMaster);
            sb.AppendLine("joinDetailKey=" + JoinDetailKey);
            sb.AppendLine("joinMasterKey=" + JoinMasterKey);
            sb.AppendLine("joinDate=" + JoinDate);
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(true));
        }

        public static string UserDir
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "DrugFlow", "profiles");
            }
        }

        public static string ExeDir
        {
            get { return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "profiles"); }
        }

        public static List<Profile> LoadAll()
        {
            List<Profile> list = new List<Profile>();
            string[] dirs = new string[] { UserDir, ExeDir };
            for (int d = 0; d < dirs.Length; d++)
            {
                try
                {
                    if (!Directory.Exists(dirs[d])) continue;
                    string[] files = Directory.GetFiles(dirs[d], "*.ini");
                    Array.Sort(files, StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < files.Length; i++)
                    {
                        try
                        {
                            Profile p = Load(files[i]);
                            if (string.IsNullOrEmpty(p.Kind) || string.IsNullOrEmpty(p.Connection)) continue;
                            bool dup = false;
                            for (int k = 0; k < list.Count; k++)
                            {
                                if (string.Equals(list[k].Name, p.Name, StringComparison.OrdinalIgnoreCase)) { dup = true; break; }
                            }
                            if (!dup) list.Add(p);
                        }
                        catch { }
                    }
                }
                catch { }
            }
            // “上次使用”排最前面
            list.Sort(delegate(Profile a, Profile b)
            {
                bool la = a.Name == "上次使用", lb = b.Name == "上次使用";
                if (la != lb) return la ? -1 : 1;
                return string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
            });
            return list;
        }

        public static string SaveLast(Profile p)
        {
            try
            {
                if (!Directory.Exists(UserDir)) Directory.CreateDirectory(UserDir);
                string path = Path.Combine(UserDir, "last.ini");
                p.Name = "上次使用";
                p.Save(path);
                return path;
            }
            catch { return ""; }
        }
    }
}
