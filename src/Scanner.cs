using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace DrugFlow
{
    public static class Scanner
    {
        public static readonly string[] Keywords = new string[]
        {
            "医药", "药店", "药品", "药房", "卫生", "诊所", "门诊", "医院", "处方", "收费", "医保",
            "his", "pos", "pharmacy", "drug", "medicine", "clinic", "hospital", "retail", "处方药"
        };

        public static bool LooksMedical(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            for (int i = 0; i < Keywords.Length; i++)
            {
                string k = Keywords[i];
                bool ascii = true;
                for (int j = 0; j < k.Length; j++) if (k[j] > 127) { ascii = false; break; }
                if (ascii)
                {
                    // 英文关键词按“整词”匹配，避免 Photoshop 里的 shop、position 里的 pos
                    if (Regex.IsMatch(s, @"\b" + Regex.Escape(k) + @"\b", RegexOptions.IgnoreCase)) return true;
                }
                else
                {
                    if (s.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) return true;
                }
            }
            return false;
        }

        private static void P(Action<string> progress, string msg)
        {
            if (progress != null) { try { progress(msg); } catch { } }
        }

        public static ScanResult Scan(Action<string> progress)
        {
            ScanResult r = new ScanResult();
            HashSet<string> dirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            P(progress, "读取系统信息…");
            r.Notes.Add("操作系统：" + OsName());
            r.Notes.Add("64 位系统：" + (Environment.Is64BitOperatingSystem ? "是（本程序为 32 位，不影响使用）" : "否"));
            r.Notes.Add("计算机名：" + Environment.MachineName + "    用户：" + Environment.UserName);
            r.Notes.Add(".NET Framework：" + SafeNetVersion());
            string office = DetectOffice();
            r.Notes.Add("办公软件：" + office);
            r.Notes.Add("数据库驱动（ODBC）：" + Tx.Join(OdbcDrivers().ToArray(), ", "));
            r.Notes.Add("数据库相关服务：" + DbServices());

            // ---------- 1. 进程 ----------
            // 为了不被杀毒软件当成“枚举进程 / 窃取信息”，本程序不读取其它进程的信息。
            r.Notes.Add("未读取其它进程信息（避免杀毒软件误报）。");

            // ---------- 2. 已安装程序 ----------
            P(progress, "扫描已安装程序列表…");
            RegistryKey[] uninstRoots = new RegistryKey[]
            {
                Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
                Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            };
            foreach (RegistryKey root in uninstRoots)
            {
                if (root == null) continue;
                try
                {
                    string[] subs = root.GetSubKeyNames();
                    for (int i = 0; i < subs.Length; i++)
                    {
                        RegistryKey k = null;
                        try
                        {
                            k = root.OpenSubKey(subs[i]);
                            if (k == null) continue;
                            string name = Convert.ToString(k.GetValue("DisplayName"));
                            if (string.IsNullOrEmpty(name)) continue;
                            string loc = Convert.ToString(k.GetValue("InstallLocation"));
                            string pub = Convert.ToString(k.GetValue("Publisher"));
                            string ver = Convert.ToString(k.GetValue("DisplayVersion"));
                            if (!LooksMedical(name + " " + pub)) { k.Close(); continue; }
                            if (!seen.Add("REG|" + name)) { k.Close(); continue; }
                            r.Findings.Add(new Finding
                            {
                                Kind = "已安装程序",
                                Title = name,
                                Detail = (string.IsNullOrEmpty(ver) ? "" : "v" + ver + "  ") + Tx.Nz(pub),
                                Path = loc
                            });
                            if (!string.IsNullOrEmpty(loc) && Directory.Exists(loc)) dirs.Add(loc);
                        }
                        catch { }
                        finally { if (k != null) k.Close(); }
                    }
                }
                catch { }
                finally { root.Close(); }
            }

            // ---------- 3. 安装目录 ----------
            P(progress, "查找疑似安装目录…");
            string[] pf = new string[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            };
            foreach (string baseDir in pf)
            {
                if (string.IsNullOrEmpty(baseDir) || !Directory.Exists(baseDir)) continue;
                try
                {
                    string[] sub = Directory.GetDirectories(baseDir);
                    for (int i = 0; i < sub.Length; i++)
                    {
                        string n = Path.GetFileName(sub[i]);
                        if (LooksMedical(n)) { dirs.Add(sub[i]); }
                    }
                }
                catch { }
            }
            // 进程目录的下一级也纳入（很多软件数据在子目录）
            List<string> firstLevel = new List<string>(dirs);
            foreach (string d in firstLevel)
            {
                try
                {
                    string[] sub = Directory.GetDirectories(d);
                    for (int i = 0; i < sub.Length && i < 40; i++) dirs.Add(sub[i]);
                }
                catch { }
            }
            foreach (string d in dirs) r.CandidateDirs.Add(d);

            // ---------- 4. 文件扫描 ----------
            P(progress, "扫描数据库文件和配置文件…");
            List<string> dbFiles = new List<string>();
            List<string> cfgFiles = new List<string>();
            string[] dbExt = new string[] { ".mdb", ".accdb", ".db", ".db3", ".sqlite", ".fdb", ".gdb", ".mdf", ".dbf" };
            string[] cfgExt = new string[] { ".ini", ".config", ".xml", ".udl", ".txt", ".properties", ".cnf", ".json", ".cfg", ".set" };

            int fileBudget = 6000;
            foreach (string d in dirs)
            {
                CollectFiles(d, 3, dbExt, cfgExt, dbFiles, cfgFiles, ref fileBudget);
                if (fileBudget <= 0) break;
            }

            // 去重
            dbFiles = Dedup(dbFiles);
            cfgFiles = Dedup(cfgFiles);

            foreach (string f in dbFiles)
            {
                long size = 0;
                try { size = new FileInfo(f).Length; } catch { }
                r.Findings.Add(new Finding
                {
                    Kind = "数据库文件",
                    Title = Path.GetFileName(f),
                    Detail = FmtSize(size) + "  " + f,
                    Path = f
                });
            }

            // ---------- 5. 配置文件中的连接串 ----------
            P(progress, "分析配置文件中的连接串…");
            Regex connRe = new Regex(@"(?i)(provider\s*=|data\s*source\s*=|initial\s*catalog\s*=|dbname\s*=|database\s*=)[^;\r\n]{2,160}");
            Regex pathRe = new Regex(@"(?i)[a-z]:\\[^""'<>\r\n]{1,200}?\.(mdb|accdb|fdb|gdb|db|sqlite)");
            Regex fullRe = new Regex(@"(?i)(provider|data\s*source|server)\s*=");
            List<string> configConns = new List<string>();
            foreach (string f in cfgFiles)
            {
                string text;
                try
                {
                    FileInfo fi = new FileInfo(f);
                    if (fi.Length > 2 * 1024 * 1024) continue;
                    text = File.ReadAllText(f, Encoding.Default);
                }
                catch { continue; }

                MatchCollection mc = connRe.Matches(text);
                int hits = 0;
                for (int i = 0; i < mc.Count && hits < 6; i++)
                {
                    string s = mc[i].Value.Trim();
                    if (s.Length < 6) continue;
                    r.Findings.Add(new Finding { Kind = "配置连接串", Title = Path.GetFileName(f), Detail = Tx.MaskSecret(s), Path = f });
                    hits++;
                }

                // 收集完整的连接串，稍后拿去试连
                if (LooksMedical(f) || f.IndexOf("app", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    MatchCollection mf = fullRe.Matches(text);
                    for (int i = 0; i < mf.Count && i < 10; i++)
                    {
                        string seg = ConnSegment(text, mf[i].Index);
                        if (seg == null) continue;
                        if (configConns.Count < 12 && !configConns.Contains(seg)) configConns.Add(seg);
                    }
                }

                MatchCollection mp = pathRe.Matches(text);
                for (int i = 0; i < mp.Count && i < 4; i++)
                {
                    string s = mp[i].Value.Trim();
                    if (dbFiles.FindIndex(delegate(string x) { return string.Equals(x, s, StringComparison.OrdinalIgnoreCase); }) < 0)
                    {
                        r.Findings.Add(new Finding { Kind = "配置中的数据文件", Title = Path.GetFileName(f), Detail = s, Path = s });
                    }
                }
            }

            // ---------- 6. SQL Server ----------
            P(progress, "探测本机 SQL Server…");
            List<string> instances = DbFactory.LocalSqlInstances();
            List<string> servers = new List<string>();
            if (instances.Count == 0)
            {
                servers.Add(".");
                servers.Add(@".\SQLEXPRESS");
            }
            else
            {
                for (int i = 0; i < instances.Count; i++)
                {
                    if (string.Equals(instances[i], "MSSQLSERVER", StringComparison.OrdinalIgnoreCase)) servers.Add(instances[i]);
                    else servers.Add(@".\" + instances[i]);
                }
            }
            foreach (string server in servers)
            {
                try
                {
                    List<string> dbs = DbFactory.SqlServerDatabases(server);
                    AddSqlServerSources(r, server, dbs, false);
                }
                catch (Exception ex)
                {
                    // 只有“连得上但握手/加密谈不拢”才值得再试 ODBC；
                    // 服务器根本没开、实例名不对，换 ODBC 也一样，徒增等待时间。
                    bool ok = false;
                    if (LooksLikeProtocolError(ex))
                    {
                        try
                        {
                            List<string> dbs = DbFactory.SqlServerDatabasesOdbc(server);
                            AddSqlServerSources(r, server, dbs, true);
                            r.Notes.Add("实例 " + server + "：SqlClient 连不上，改用 ODBC 成功。原因：" + Tx.Short(ex.Message, 100));
                            ok = true;
                        }
                        catch { }
                    }
                    if (!ok) r.Notes.Add("SQL Server 实例 " + server + " 连接失败：" + Tx.Short(ex.Message, 160));
                }
            }

            // ---------- 6.5 配置文件里的连接串（可能是远程库或 SQL 账号） ----------
            if (configConns.Count > 0)
            {
                P(progress, "尝试配置文件里的连接串…");
                int okCount = 0;
                int tried = 0;
                for (int i = 0; i < configConns.Count && tried < 8; i++)
                {
                    tried++;
                    if (TryConfigConnection(r, configConns[i])) okCount++;
                }
                if (okCount > 0) r.Notes.Add("配置文件里的连接串：" + okCount + " 个连接成功");
            }

            // ---------- 7. Access 文件 ----------
            P(progress, "尝试打开 Access 数据库…");
            foreach (string f in dbFiles)
            {
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".mdb" && ext != ".accdb") continue;
                try
                {
                    using (AccessProvider ap = new AccessProvider(f))
                    {
                        List<string> tables = ap.ListTables();
                        r.Findings.Add(new Finding
                        {
                            Kind = "Access 可打开",
                            Title = Path.GetFileName(f),
                            Detail = tables.Count + " 张表：" + Tx.Short(Tx.Join(tables.ToArray(), ", "), 300),
                            Path = f
                        });
                        SourceItem s = new SourceItem();
                        s.Kind = "access";
                        s.Title = "Access  " + Path.GetFileName(f) + "  (" + tables.Count + " 表)";
                        s.Connection = f;
                        s.Path = f;
                        s.Database = Path.GetFileNameWithoutExtension(f);
                        if (!r.Sources.Exists(delegate(SourceItem x) { return x.Connection == s.Connection; })) r.Sources.Add(s);
                    }
                }
                catch (Exception ex)
                {
                    r.Notes.Add("Access 文件打不开：" + f + "  —  " + Tx.Short(ex.Message, 160));
                }
            }

            // ---------- 7.5 SQLite ----------
            P(progress, "尝试打开 SQLite 数据库…");
            int sqliteTried = 0;
            foreach (string f in dbFiles)
            {
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext != ".db" && ext != ".db3" && ext != ".sqlite") continue;
                if (sqliteTried >= 25) break;
                long size = 0;
                try { size = new FileInfo(f).Length; } catch { }
                if (size < 4096) continue;
                sqliteTried++;
                try
                {
                    using (SqliteProvider sp = new SqliteProvider(f))
                    {
                        List<string> tables = sp.ListTables();
                        r.Findings.Add(new Finding
                        {
                            Kind = "SQLite 可打开",
                            Title = Path.GetFileName(f),
                            Detail = tables.Count + " 张表：" + Tx.Short(Tx.Join(tables.ToArray(), ", "), 300),
                            Path = f
                        });
                        SourceItem s = new SourceItem();
                        s.Kind = "sqlite";
                        s.Title = "SQLite  " + Path.GetFileName(f) + "  (" + tables.Count + " 表)";
                        s.Connection = f;
                        s.Path = f;
                        s.Database = Path.GetFileNameWithoutExtension(f);
                        if (!r.Sources.Exists(delegate(SourceItem x) { return x.Connection == s.Connection; })) r.Sources.Add(s);
                    }
                }
                catch (Exception ex)
                {
                    r.Notes.Add("SQLite 文件打不开：" + f + "  —  " + Tx.Short(ex.Message, 120));
                }
            }

            // ---------- 8. 其它数据库文件与 ODBC ----------
            List<string> odbcDrivers = OdbcDrivers();
            foreach (string f in dbFiles)
            {
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (ext == ".fdb" || ext == ".gdb")
                {
                    r.Findings.Add(new Finding
                    {
                        Kind = "Firebird/InterBase",
                        Title = Path.GetFileName(f),
                        Detail = "可能是 Firebird 数据库；本机 ODBC 驱动：" + Tx.Join(odbcDrivers.ToArray(), ", "),
                        Path = f
                    });
                }
                else if (ext == ".db" || ext == ".db3" || ext == ".sqlite")
                {
                    r.Findings.Add(new Finding
                    {
                        Kind = "SQLite 文件",
                        Title = Path.GetFileName(f),
                        Detail = "SQLite 数据库文件（未能直接打开，可能是加密库或不是 SQLite）",
                        Path = f
                    });
                }
                else if (ext == ".mdf")
                {
                    r.Findings.Add(new Finding
                    {
                        Kind = "SQL Server 数据文件",
                        Title = Path.GetFileName(f),
                        Detail = "这是 SQL Server 的数据文件，需要对应的实例在运行才能读取",
                        Path = f
                    });
                }
            }

            // ---------- 8.5 dBASE / FoxPro (*.dbf) ----------
            List<string> dbfDirs = new List<string>();
            foreach (string f in dbFiles)
            {
                if (Path.GetExtension(f).ToLowerInvariant() != ".dbf") continue;
                string d = Path.GetDirectoryName(f);
                if (!string.IsNullOrEmpty(d) && !dbfDirs.Contains(d)) dbfDirs.Add(d);
            }
            int dbfTried = 0;
            foreach (string d in dbfDirs)
            {
                if (dbfTried >= 4) break;
                dbfTried++;
                try
                {
                    string cs = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=" + d + ";";
                    using (OdbcProvider op = new OdbcProvider(cs))
                    {
                        List<string> tables = op.ListTables();
                        if (tables.Count == 0) continue;
                        SourceItem s = new SourceItem();
                        s.Kind = "odbc";
                        s.Title = "DBF 目录  " + d + "  (" + tables.Count + " 表)";
                        s.Connection = cs;
                        s.Path = d;
                        r.Sources.Add(s);
                        r.Findings.Add(new Finding
                        {
                            Kind = "DBF 可打开",
                            Title = Path.GetFileName(d),
                            Detail = tables.Count + " 张表：" + Tx.Short(Tx.Join(tables.ToArray(), ", "), 250),
                            Path = d
                        });
                    }
                }
                catch (Exception ex) { r.Notes.Add("DBF 目录打不开：" + d + " — " + Tx.Short(ex.Message, 120)); }
            }

            // ---------- 9. 示例源 ----------
            SourceItem demo = new SourceItem();
            demo.Kind = "demo";
            demo.Title = "示例数据（先看看导出效果）";
            demo.Connection = "";
            r.Sources.Insert(0, demo);

            P(progress, "整理数据库表结构…");
            DumpSchema(r);

            P(progress, "扫描完成");
            return r;
        }

        /// <summary>把找到的数据库的表名、字段名记进报告，方便远程写 SQL</summary>
        public static void DumpSchema(ScanResult r)
        {
            int tablesBudget = 120;
            for (int si = 0; si < r.Sources.Count; si++)
            {
                SourceItem s = r.Sources[si];
                if (s.Kind == "demo") continue;
                if (tablesBudget <= 0) { r.Schema.Add("（表太多，后面的省略）"); break; }
                IProvider p = null;
                try
                {
                    p = DbFactory.Create(s);
                    List<string> tables = p.ListTables();
                    r.Schema.Add("【" + s.Title + "】共 " + tables.Count + " 张表");

                    // 把“像销售流水”的表排在前面，方便直接定位
                    List<string> ranked = new List<string>(tables);
                    ranked.Sort(delegate(string a, string b)
                    {
                        int sa = TableScore(a), sb = TableScore(b);
                        if (sa != sb) return sb.CompareTo(sa);
                        return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
                    });

                    int detailed = Math.Min(ranked.Count, 10);
                    for (int i = 0; i < detailed && tablesBudget > 0; i++)
                    {
                        tablesBudget--;
                        try
                        {
                            List<ColumnInfo> cols = p.ListColumns(ranked[i]);
                            StringBuilder sb = new StringBuilder();
                            for (int k = 0; k < cols.Count && k < 30; k++)
                            {
                                if (k > 0) sb.Append(", ");
                                sb.Append(cols[k].Name);
                                if (!string.IsNullOrEmpty(cols[k].Type)) sb.Append("(" + cols[k].Type + ")");
                            }
                            if (cols.Count > 30) sb.Append(", …");
                            r.Schema.Add("  " + ranked[i] + " : " + sb.ToString());
                        }
                        catch (Exception ex)
                        {
                            r.Schema.Add("  " + ranked[i] + " : 字段读取失败 " + Tx.Short(ex.Message, 80));
                        }
                    }
                    if (ranked.Count > detailed)
                    {
                        int rest = Math.Min(40, ranked.Count - detailed);
                        string[] names = ranked.GetRange(detailed, rest).ToArray();
                        r.Schema.Add("  其余 " + (ranked.Count - detailed) + " 张表：" + Tx.Join(names, ", "));
                    }
                }
                catch (Exception ex)
                {
                    r.Schema.Add("【" + s.Title + "】打不开：" + Tx.Short(ex.Message, 100));
                }
                finally
                {
                    if (p != null) { try { p.Dispose(); } catch { } }
                }
            }
        }

        /// <summary>按表名给“像不像销售流水”打分</summary>
        public static int TableScore(string name)
        {
            string n = (name == null ? "" : name).ToLowerInvariant();
            int s = 0;
            if (n.IndexOf("销售明细") >= 0) s += 100;
            if (n.IndexOf("saledetail") >= 0 || n.IndexOf("sale_detail") >= 0) s += 100;
            if (n.IndexOf("销售") >= 0 || n.IndexOf("sale") >= 0 || n.IndexOf("零售") >= 0 || n.IndexOf("retail") >= 0) s += 50;
            if (n.IndexOf("明细") >= 0 || n.IndexOf("detail") >= 0 || n.IndexOf("流水") >= 0) s += 40;
            if (n.IndexOf("收费") >= 0 || n.IndexOf("charge") >= 0 || n.IndexOf("bill") >= 0) s += 25;
            if (n.IndexOf("药品") >= 0 || n.IndexOf("drug") >= 0) s += 20;
            if (n.IndexOf("入库") >= 0 || n.IndexOf("进货") >= 0 || n.IndexOf("purchase") >= 0 || n.IndexOf("stock") >= 0 || n.IndexOf("库存") >= 0) s -= 80;
            if (n.IndexOf("日志") >= 0 || n.IndexOf("log") >= 0) s -= 60;
            if (n.IndexOf("sys") >= 0) s -= 30;
            return s;
        }

        private static string SafeNetVersion()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                {
                    if (k != null)
                    {
                        object v = k.GetValue("Version");
                        if (v != null) return Convert.ToString(v);
                    }
                }
            }
            catch { }
            return "未知";
        }

        private static string ConnSegment(string text, int idx)
        {
            int start = text.LastIndexOf('\n', Math.Max(0, idx - 1));
            start = start < 0 ? 0 : start + 1;
            int end = text.IndexOf('\n', idx);
            if (end < 0) end = text.Length;
            string line = text.Substring(start, end - start).Trim();
            line = line.Trim('"', '\'', ' ', '\t', '\r');
            if (line.Length > 400)
            {
                int from = Math.Max(0, idx - start - 120);
                line = line.Substring(from, Math.Min(400, line.Length - from));
            }
            if (line.Length < 15) return null;
            if (line.IndexOf('=') < 0) return null;
            string lower = line.ToLowerInvariant();
            if (lower.IndexOf("initial catalog") < 0 && lower.IndexOf("database") < 0 &&
                lower.IndexOf("user id") < 0 && lower.IndexOf("uid") < 0 &&
                lower.IndexOf("pwd") < 0 && lower.IndexOf("password") < 0)
                return null;
            if (lower.IndexOf(".mdb") >= 0 || lower.IndexOf(".accdb") >= 0) return null;
            return line;
        }

        private static Dictionary<string, string> ParseConn(string raw)
        {
            Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string[] parts = raw.Split(';');
            for (int i = 0; i < parts.Length; i++)
            {
                int eq = parts[i].IndexOf('=');
                if (eq <= 0) continue;
                string k = parts[i].Substring(0, eq).Trim();
                string v = parts[i].Substring(eq + 1).Trim().Trim('"', '\'');
                if (k.Length > 0) d[k] = v;
            }
            return d;
        }

        private static string First(Dictionary<string, string> d, params string[] keys)
        {
            for (int i = 0; i < keys.Length; i++)
            {
                string v;
                if (d.TryGetValue(keys[i], out v) && !string.IsNullOrEmpty(v)) return v;
            }
            return "";
        }

        private static bool TryConfigConnection(ScanResult r, string raw)
        {
            Dictionary<string, string> d;
            string server, db, uid, pwd;
            try
            {
                d = ParseConn(raw);
                server = First(d, "Data Source", "Server", "Address", "Addr", "Network Address");
                if (string.IsNullOrEmpty(server)) return false;
                db = First(d, "Initial Catalog", "Database");
                uid = First(d, "User ID", "UID", "User");
                pwd = First(d, "Password", "PWD");
            }
            catch { return false; }

            // 先按 SqlClient 连
            Exception ex0 = null;
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Data Source=").Append(server).Append(";Connect Timeout=2;Application Name=DrugFlow;");
                if (!string.IsNullOrEmpty(db)) sb.Append("Initial Catalog=").Append(db).Append(';');
                if (!string.IsNullOrEmpty(uid))
                    sb.Append("User ID=").Append(uid).Append(";Password=").Append(pwd).Append(';');
                else
                    sb.Append("Integrated Security=SSPI;");
                string cs = sb.ToString();

                if (ExistsConn(r, cs)) return true;
                using (SqlConnection cn = new SqlConnection(cs))
                {
                    cn.Open();
                    AddConfigSource(r, server, cn.Database, cs, "sqlserver", raw);
                    return true;
                }
            }
            catch (Exception ex) { ex0 = ex; }

            // 再按 ODBC 连（老服务器握手失败时）
            if (ex0 == null || !LooksLikeProtocolError(ex0)) return false;
            try
            {
                string ocs = DbFactory.SqlServerOdbcConnection(server,
                    string.IsNullOrEmpty(db) ? "master" : db, uid, pwd);
                if (ExistsConn(r, ocs)) return true;
                using (OdbcConnection cn = new OdbcConnection(ocs))
                {
                    cn.Open();
                    AddConfigSource(r, server, cn.Database, ocs, "odbc", raw);
                    return true;
                }
            }
            catch { }

            return false;
        }

        private static bool LooksLikeProtocolError(Exception ex)
        {
            if (ex == null) return false;
            string m = (ex.Message == null ? "" : ex.Message).ToLowerInvariant();
            return m.IndexOf("common algorithm") >= 0
                || m.IndexOf("tls") >= 0
                || m.IndexOf("ssl") >= 0
                || m.IndexOf("pre-login") >= 0
                || m.IndexOf("prelogin") >= 0
                || m.IndexOf("handshake") >= 0
                || m.IndexOf("协议") >= 0
                || m.IndexOf("加密") >= 0;
        }

        private static bool ExistsConn(ScanResult r, string cs)
        {
            for (int i = 0; i < r.Sources.Count; i++)
            {
                if (string.Equals(r.Sources[i].Connection, cs, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static void AddConfigSource(ScanResult r, string server, string actual, string cs, string kind, string raw)
        {
            SourceItem s = new SourceItem();
            s.Kind = kind;
            s.Connection = cs;
            s.Database = actual;
            s.Title = "配置里的库  " + server + " / " + actual + "  ★";
            r.Sources.Add(s);
            r.Findings.Add(new Finding
            {
                Kind = "配置连接串可用",
                Title = server + " / " + actual,
                Detail = "用配置文件里的连接串连上了：" + Tx.MaskSecret(raw),
                Path = server
            });
        }

        private static void AddSqlServerSources(ScanResult r, string server, List<string> dbs, bool viaOdbc)
        {
            r.Findings.Add(new Finding
            {
                Kind = "SQL Server",
                Title = "实例 " + server + (viaOdbc ? "（ODBC）" : "") + " 可连接",
                Detail = "数据库：" + Tx.Join(dbs.ToArray(), ", "),
                Path = server
            });
            for (int i = 0; i < dbs.Count; i++)
            {
                SourceItem s = new SourceItem();
                s.Kind = viaOdbc ? "odbc" : "sqlserver";
                s.Database = dbs[i];
                s.Title = "SQL Server" + (viaOdbc ? "(ODBC)" : "") + " " + server + " / " + dbs[i] +
                          (LooksMedical(dbs[i]) ? "  ★" : "");
                s.Connection = viaOdbc
                    ? DbFactory.SqlServerOdbcConnection(server, dbs[i], null, null)
                    : DbFactory.SqlServerConnection(server, dbs[i]);
                r.Sources.Add(s);
            }
        }

        private static string DbServices()
        {
            List<string> list = new List<string>();
            try
            {
                ServiceController[] svcs = ServiceController.GetServices();
                for (int i = 0; i < svcs.Length; i++)
                {
                    string n = (svcs[i].ServiceName == null ? "" : svcs[i].ServiceName).ToLowerInvariant();
                    string d = (svcs[i].DisplayName == null ? "" : svcs[i].DisplayName).ToLowerInvariant();
                    bool hit = n.StartsWith("mssql") || n.StartsWith("mysql") || n.StartsWith("firebird")
                               || n.StartsWith("oracle") || n.StartsWith("sybase")
                               || d.IndexOf("sql server") >= 0 || d.IndexOf("mysql") >= 0
                               || d.IndexOf("firebird") >= 0;
                    if (hit) list.Add(svcs[i].ServiceName + "=" + svcs[i].Status);
                }
            }
            catch { }
            if (list.Count == 0) return "未发现 SQL Server / MySQL / Firebird 等服务";
            return Tx.Join(list.ToArray(), "; ");
        }

        private static string OsName()
        {
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (k != null)
                    {
                        string p = Convert.ToString(k.GetValue("ProductName"));
                        string b = Convert.ToString(k.GetValue("CurrentBuild"));
                        if (!string.IsNullOrEmpty(p))
                            return p + (string.IsNullOrEmpty(b) ? "" : "  Build " + b);
                    }
                }
            }
            catch { }
            return Environment.OSVersion.VersionString;
        }

        private static string DetectOffice()
        {
            List<string> found = new List<string>();
            try
            {
                string[] roots = new string[]
                {
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                };
                foreach (string rp in roots)
                {
                    using (RegistryKey root = Registry.LocalMachine.OpenSubKey(rp))
                    {
                        if (root == null) continue;
                        foreach (string sub in root.GetSubKeyNames())
                        {
                            using (RegistryKey k = root.OpenSubKey(sub))
                            {
                                if (k == null) continue;
                                string name = Convert.ToString(k.GetValue("DisplayName"));
                                if (string.IsNullOrEmpty(name)) continue;
                                if (name.IndexOf("Microsoft Office", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    name.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    name.IndexOf("金山", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    name.IndexOf("LibreOffice", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    if (!found.Contains(name)) found.Add(name);
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            // 再按可执行文件确认（注册表有时列不全）
            string[] appPaths = new string[]
            {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\et.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\et.exe"
            };
            foreach (string ap in appPaths)
            {
                try
                {
                    using (RegistryKey k = Registry.LocalMachine.OpenSubKey(ap))
                    {
                        if (k == null) continue;
                        string p = Convert.ToString(k.GetValue(""));
                        if (!string.IsNullOrEmpty(p) && File.Exists(p))
                        {
                            string label = "Excel/WPS：" + p;
                            if (!found.Contains(label)) found.Add(label);
                        }
                    }
                }
                catch { }
            }

            if (found.Count == 0) return "未检测到 Excel / WPS（本工具生成 Excel 不需要安装 Office）";
            return Tx.Join(found.ToArray(), "；");
        }

        private static List<string> OdbcDrivers()
        {
            List<string> list = new List<string>();
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"))
                {
                    if (k != null) list.AddRange(k.GetValueNames());
                }
            }
            catch { }
            return list;
        }

        private static void CollectFiles(string dir, int depth, string[] dbExt, string[] cfgExt,
            List<string> dbFiles, List<string> cfgFiles, ref int budget)
        {
            if (budget <= 0 || string.IsNullOrEmpty(dir)) return;
            string[] files;
            try { files = Directory.GetFiles(dir); }
            catch { return; }
            budget -= files.Length;
            for (int i = 0; i < files.Length; i++)
            {
                string ext = Path.GetExtension(files[i]).ToLowerInvariant();
                if (Array.IndexOf(dbExt, ext) >= 0) dbFiles.Add(files[i]);
                else if (Array.IndexOf(cfgExt, ext) >= 0) cfgFiles.Add(files[i]);
            }
            if (depth <= 0) return;
            string[] subs;
            try { subs = Directory.GetDirectories(dir); }
            catch { return; }
            for (int i = 0; i < subs.Length; i++) CollectFiles(subs[i], depth - 1, dbExt, cfgExt, dbFiles, cfgFiles, ref budget);
        }

        private static List<string> Dedup(List<string> src)
        {
            List<string> list = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string s in src) if (seen.Add(s)) list.Add(s);
            return list;
        }

        public static string FmtSize(long bytes)
        {
            if (bytes >= 1024L * 1024 * 1024) return (bytes / 1024.0 / 1024 / 1024).ToString("0.0") + " GB";
            if (bytes >= 1024 * 1024) return (bytes / 1024.0 / 1024).ToString("0.0") + " MB";
            if (bytes >= 1024) return (bytes / 1024.0).ToString("0") + " KB";
            return bytes + " B";
        }
    }
}
