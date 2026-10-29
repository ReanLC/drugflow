using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Text;
using Microsoft.Win32;

namespace DrugFlow
{
    // =====================================================================
    // SQL Server
    // =====================================================================
    public class SqlServerProvider : IProvider
    {
        private readonly string _cs;
        private SqlConnection _cn;

        public SqlServerProvider(string connectionString) { _cs = connectionString; }
        public string Kind { get { return "sqlserver"; } }

        private SqlConnection Cn
        {
            get
            {
                if (_cn == null)
                {
                    _cn = new SqlConnection(_cs);
                    _cn.Open();
                }
                return _cn;
            }
        }

        public List<string> ListTables()
        {
            List<string> list = new List<string>();
            string sql = "SELECT TABLE_SCHEMA, TABLE_NAME FROM INFORMATION_SCHEMA.TABLES " +
                         "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_SCHEMA, TABLE_NAME";
            using (SqlCommand cmd = new SqlCommand(sql, Cn))
            using (SqlDataReader r = cmd.ExecuteReader())
            {
                while (r.Read())
                {
                    string schema = r.IsDBNull(0) ? "dbo" : r.GetString(0);
                    list.Add(schema + "." + r.GetString(1));
                }
            }
            return list;
        }

        public List<ColumnInfo> ListColumns(string table)
        {
            string schema, name;
            SplitTable(table, out schema, out name);
            List<ColumnInfo> list = new List<ColumnInfo>();
            string sql = "SELECT COLUMN_NAME, DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS " +
                         "WHERE TABLE_NAME = '" + name.Replace("'", "''") + "' " +
                         "AND TABLE_SCHEMA = '" + schema.Replace("'", "''") + "' ORDER BY ORDINAL_POSITION";
            using (SqlCommand cmd = new SqlCommand(sql, Cn))
            using (SqlDataReader r = cmd.ExecuteReader())
            {
                while (r.Read())
                {
                    list.Add(new ColumnInfo(r.GetString(0), r.IsDBNull(1) ? "" : r.GetString(1)));
                }
            }
            return list;
        }

        public DataTable Query(string sql)
        {
            DataTable dt = new DataTable();
            using (SqlDataAdapter da = new SqlDataAdapter(sql, Cn)) da.Fill(dt);
            return dt;
        }

        public string Quote(string name) { return "[" + name.Replace("]", "]]") + "]"; }

        public string QuoteTable(string name)
        {
            string schema, table;
            SplitTable(name, out schema, out table);
            return Quote(schema) + "." + Quote(table);
        }

        public string DateLiteral(DateTime d) { return DateLiteral(d, "auto"); }

        public string DateLiteral(DateTime d, string style)
        {
            if (style == "ymd8") return "'" + d.ToString("yyyyMMdd") + "'";
            if (style == "slash") return "'" + d.ToString("yyyy/MM/dd") + "'";
            if (style == "ymd") return "'" + d.ToString("yyyy-MM-dd") + "'";
            return "'" + d.ToString("yyyy-MM-dd HH:mm:ss") + "'";
        }

        public void Dispose() { if (_cn != null) { _cn.Dispose(); _cn = null; } }

        public static void SplitTable(string full, out string schema, out string table)
        {
            schema = "dbo";
            table = full;
            if (string.IsNullOrEmpty(full)) return;
            int i = full.LastIndexOf('.');
            if (i > 0)
            {
                schema = full.Substring(0, i).Trim('[', ']', ' ');
                table = full.Substring(i + 1).Trim('[', ']', ' ');
            }
            else
            {
                table = full.Trim('[', ']', ' ');
            }
        }
    }

    // =====================================================================
    // Access（.mdb / .accdb）
    // =====================================================================
    public class AccessProvider : IProvider
    {
        private readonly string _path;
        private readonly string _cs;
        private OleDbConnection _cn;

        public AccessProvider(string path)
        {
            _path = path;
            _cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Persist Security Info=False;";
        }

        public string Kind { get { return "access"; } }

        private OleDbConnection Cn
        {
            get
            {
                if (_cn == null)
                {
                    OleDbConnection c = null;
                    string[] providers = new string[]
                    {
                        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _path + ";Persist Security Info=False;",
                        "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + _path + ";Persist Security Info=False;",
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _path + ";Persist Security Info=False;"
                    };
                    Exception last = null;
                    for (int i = 0; i < providers.Length; i++)
                    {
                        try
                        {
                            OleDbConnection t = new OleDbConnection(providers[i]);
                            t.Open();
                            c = t;
                            break;
                        }
                        catch (Exception ex) { last = ex; }
                    }
                    if (c == null) throw new Exception("无法打开 Access 数据库（缺少 Access 数据库引擎）：" + (last == null ? "" : last.Message));
                    _cn = c;
                }
                return _cn;
            }
        }

        public List<string> ListTables()
        {
            List<string> list = new List<string>();
            DataTable schema = Cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                new object[] { null, null, null, "TABLE" });
            if (schema != null)
            {
                foreach (DataRow row in schema.Rows)
                {
                    string n = Convert.ToString(row["TABLE_NAME"]);
                    if (!string.IsNullOrEmpty(n)) list.Add(n);
                }
            }
            list.Sort(StringComparer.OrdinalIgnoreCase);
            return list;
        }

        public List<ColumnInfo> ListColumns(string table)
        {
            List<ColumnInfo> list = new List<ColumnInfo>();
            DataTable schema = Cn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                new object[] { null, null, table, null });
            if (schema != null)
            {
                foreach (DataRow row in schema.Rows)
                {
                    string n = Convert.ToString(row["COLUMN_NAME"]);
                    string t = "";
                    try { t = Convert.ToString(row["DATA_TYPE"]); } catch { }
                    list.Add(new ColumnInfo(n, t));
                }
            }
            return list;
        }

        public DataTable Query(string sql)
        {
            DataTable dt = new DataTable();
            using (OleDbDataAdapter da = new OleDbDataAdapter(sql, Cn)) da.Fill(dt);
            return dt;
        }

        public string Quote(string name) { return "[" + name.Replace("]", "]]") + "]"; }
        public string QuoteTable(string name) { return Quote(name); }
        public string DateLiteral(DateTime d) { return DateLiteral(d, "auto"); }

        public string DateLiteral(DateTime d, string style)
        {
            if (style == null || style == "auto") return "#" + d.ToString("yyyy-MM-dd HH:mm:ss") + "#";
            if (style == "ymd8") return "'" + d.ToString("yyyyMMdd") + "'";
            if (style == "slash") return "'" + d.ToString("yyyy/MM/dd") + "'";
            return "'" + d.ToString("yyyy-MM-dd") + "'";
        }
        public void Dispose() { if (_cn != null) { _cn.Dispose(); _cn = null; } }
    }

    // =====================================================================
    // 通用 ODBC（Firebird / MySQL / SQL Server / 其它带 ODBC 驱动的库）
    // =====================================================================
    public class OdbcProvider : IProvider
    {
        private readonly string _cs;
        private OdbcConnection _cn;

        public OdbcProvider(string connectionString) { _cs = connectionString; }
        public string Kind { get { return "odbc"; } }

        private OdbcConnection Cn
        {
            get
            {
                if (_cn == null) { _cn = new OdbcConnection(_cs); _cn.Open(); }
                return _cn;
            }
        }

        public List<string> ListTables()
        {
            List<string> list = new List<string>();
            DataTable schema = Cn.GetSchema("Tables");
            foreach (DataRow row in schema.Rows)
            {
                string type = Convert.ToString(row["TABLE_TYPE"]);
                if (type.IndexOf("TABLE", StringComparison.OrdinalIgnoreCase) < 0) continue;
                string schemaName = "";
                try { schemaName = Convert.ToString(row["TABLE_SCHEM"]); } catch { }
                string name = Convert.ToString(row["TABLE_NAME"]);
                if (string.IsNullOrEmpty(name)) continue;
                list.Add(string.IsNullOrEmpty(schemaName) ? name : schemaName + "." + name);
            }
            list.Sort(StringComparer.OrdinalIgnoreCase);
            return list;
        }

        public List<ColumnInfo> ListColumns(string table)
        {
            List<ColumnInfo> list = new List<ColumnInfo>();
            string schema = null, name = table;
            int i = table.LastIndexOf('.');
            if (i > 0) { schema = table.Substring(0, i); name = table.Substring(i + 1); }
            DataTable dt = Cn.GetSchema("Columns", new string[] { null, schema, name, null });
            foreach (DataRow row in dt.Rows)
            {
                string n = Convert.ToString(row["COLUMN_NAME"]);
                string t = "";
                try { t = Convert.ToString(row["TYPE_NAME"]); } catch { }
                list.Add(new ColumnInfo(n, t));
            }
            return list;
        }

        public DataTable Query(string sql)
        {
            DataTable dt = new DataTable();
            using (OdbcDataAdapter da = new OdbcDataAdapter(sql, Cn)) da.Fill(dt);
            return dt;
        }

        public string Quote(string name) { return "\"" + name.Replace("\"", "\"\"") + "\""; }

        public string QuoteTable(string name)
        {
            string[] parts = name.Split('.');
            for (int i = 0; i < parts.Length; i++) parts[i] = Quote(parts[i].Trim('"'));
            return string.Join(".", parts);
        }

        public string DateLiteral(DateTime d) { return DateLiteral(d, "auto"); }

        public string DateLiteral(DateTime d, string style)
        {
            if (style == "ymd8") return "'" + d.ToString("yyyyMMdd") + "'";
            if (style == "slash") return "'" + d.ToString("yyyy/MM/dd") + "'";
            if (style == "ymd") return "'" + d.ToString("yyyy-MM-dd") + "'";
            return "'" + d.ToString("yyyy-MM-dd HH:mm:ss") + "'";
        }
        public void Dispose() { if (_cn != null) { _cn.Dispose(); _cn = null; } }
    }

    // =====================================================================
    // 示例数据（没有数据库也能看到成品效果）
    // =====================================================================
    public class DemoProvider : IProvider
    {
        public string Kind { get { return "demo"; } }

        public List<string> ListTables() { return new List<string>(new string[] { "示例_销售流水" }); }

        public List<ColumnInfo> ListColumns(string table)
        {
            return new List<ColumnInfo>(new ColumnInfo[]
            {
                new ColumnInfo("销售日期","datetime"),
                new ColumnInfo("单据号","varchar"),
                new ColumnInfo("药品名称","varchar"),
                new ColumnInfo("规格","varchar"),
                new ColumnInfo("单位","varchar"),
                new ColumnInfo("数量","decimal"),
                new ColumnInfo("单价","decimal"),
                new ColumnInfo("金额","decimal"),
                new ColumnInfo("批号","varchar")
            });
        }

        public DataTable Query(string sql)
        {
            DateTime lo = DateTime.MinValue, hi = DateTime.MaxValue;
            bool filtered = false;
            System.Text.RegularExpressions.Match m =
                System.Text.RegularExpressions.Regex.Match(sql, @">=\s*'([^']+)'\s+AND[\s\S]*?<\s*'([^']+)'");
            if (m.Success)
            {
                DateTime a, b;
                if (DateTime.TryParse(m.Groups[1].Value, out a) && DateTime.TryParse(m.Groups[2].Value, out b))
                {
                    lo = a; hi = b; filtered = true;
                }
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("销售日期", typeof(DateTime));
            dt.Columns.Add("单据号", typeof(string));
            dt.Columns.Add("药品名称", typeof(string));
            dt.Columns.Add("规格", typeof(string));
            dt.Columns.Add("单位", typeof(string));
            dt.Columns.Add("数量", typeof(double));
            dt.Columns.Add("单价", typeof(double));
            dt.Columns.Add("金额", typeof(double));
            dt.Columns.Add("批号", typeof(string));

            string[,] drugs = new string[,]
            {
                { "阿莫西林胶囊", "0.25g*24粒", "盒", "12.50" },
                { "感冒灵颗粒", "10g*9袋", "盒", "9.80" },
                { "布洛芬缓释胶囊", "0.3g*20粒", "盒", "15.60" },
                { "维生素C片", "100mg*100片", "瓶", "4.50" },
                { "头孢克肟分散片", "50mg*12片", "盒", "23.00" },
                { "藿香正气水", "10ml*10支", "盒", "8.20" },
                { "复方甘草片", "100片", "瓶", "6.00" },
                { "生理盐水", "500ml", "瓶", "2.80" }
            };
            DateTime day = DateTime.Today.AddDays(-29);
            int seq = 0;
            for (int d = 0; d < 30; d++)
            {
                DateTime cur = day.AddDays(d);
                int bills = 2 + (d % 3);
                for (int b = 0; b < bills; b++)
                {
                    seq++;
                    DateTime t = cur.AddHours(8 + b * 3).AddMinutes((d * 7 + b * 13) % 60);
                    if (filtered && (t < lo || t >= hi)) continue;
                    int lines = 1 + ((d + b) % 3);
                    for (int l = 0; l < lines; l++)
                    {
                        int idx = (d * 3 + b * 5 + l) % 8;
                        double price = double.Parse(drugs[idx, 3], System.Globalization.CultureInfo.InvariantCulture);
                        int qty = 1 + ((d + b + l) % 4);
                        DataRow r = dt.NewRow();
                        r["销售日期"] = t;
                        r["单据号"] = "XS" + cur.ToString("yyyyMMdd") + "-" + (b + 1);
                        r["药品名称"] = drugs[idx, 0];
                        r["规格"] = drugs[idx, 1];
                        r["单位"] = drugs[idx, 2];
                        r["数量"] = (double)qty;
                        r["单价"] = price;
                        r["金额"] = Math.Round(price * qty, 2);
                        r["批号"] = "2025" + (10 + (idx % 3)) + "01";
                        dt.Rows.Add(r);
                    }
                }
            }
            return dt;
        }

        public string Quote(string name) { return "[" + name + "]"; }
        public string QuoteTable(string name) { return Quote(name); }
        public string DateLiteral(DateTime d) { return DateLiteral(d, "auto"); }

        public string DateLiteral(DateTime d, string style)
        {
            if (style == "ymd8") return "'" + d.ToString("yyyyMMdd") + "'";
            if (style == "slash") return "'" + d.ToString("yyyy/MM/dd") + "'";
            if (style == "ymd") return "'" + d.ToString("yyyy-MM-dd") + "'";
            return "'" + d.ToString("yyyy-MM-dd HH:mm:ss") + "'";
        }
        public void Dispose() { }
    }

    // =====================================================================
    // 工厂 + SQL Server 探测
    // =====================================================================
    public static class DbFactory
    {
        public static IProvider Create(SourceItem s)
        {
            if (s == null) throw new ArgumentNullException("s");
            if (s.Kind == "demo") return new DemoProvider();
            if (s.Kind == "sqlserver") return new SqlServerProvider(s.Connection);
            if (s.Kind == "access") return new AccessProvider(s.Connection);
            if (s.Kind == "sqlite") return new SqliteProvider(s.Connection);
            if (s.Kind == "odbc") return new OdbcProvider(s.Connection);
            throw new Exception("未知的数据源类型：" + s.Kind);
        }

        public static string SqlServerConnection(string server, string database)
        {
            return "Data Source=" + server + ";Initial Catalog=" + (string.IsNullOrEmpty(database) ? "master" : database) +
                   ";Integrated Security=SSPI;Connect Timeout=4;Application Name=DrugFlow;";
        }

        /// <summary>用系统自带的 SQL Server ODBC 驱动连接（老版本 SQL Server 的备用通道）</summary>
        public static string SqlServerOdbcConnection(string server, string database, string uid, string pwd)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Driver={SQL Server};Server=").Append(server).Append(';');
            if (!string.IsNullOrEmpty(database)) sb.Append("Database=").Append(database).Append(';');
            if (!string.IsNullOrEmpty(uid)) sb.Append("Uid=").Append(uid).Append(";Pwd=").Append(pwd).Append(';');
            else sb.Append("Trusted_Connection=yes;");
            return sb.ToString();
        }

        public static List<string> SqlServerDatabasesOdbc(string server)
        {
            List<string> list = new List<string>();
            string cs = SqlServerOdbcConnection(server, "master", null, null);
            using (OdbcConnection cn = new OdbcConnection(cs))
            {
                cn.Open();
                using (OdbcCommand cmd = new OdbcCommand(
                    "SELECT name FROM sys.databases WHERE database_id > 4 AND state = 0 ORDER BY name", cn))
                using (OdbcDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read()) list.Add(r.GetString(0));
                }
            }
            return list;
        }

        /// <summary>列出本机注册的 SQL Server 实例名（如 MSSQLSERVER、SQLEXPRESS）</summary>
        public static List<string> LocalSqlInstances()
        {
            List<string> list = new List<string>();
            RegistryKey[] roots = new RegistryKey[]
            {
                Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL"),
                Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL")
            };
            foreach (RegistryKey k in roots)
            {
                if (k == null) continue;
                try
                {
                    foreach (string name in k.GetValueNames())
                    {
                        if (!string.IsNullOrEmpty(name) && !list.Contains(name)) list.Add(name);
                    }
                }
                catch { }
                finally { k.Close(); }
            }
            return list;
        }

        public static List<string> SqlServerDatabases(string server)
        {
            List<string> list = new List<string>();
            using (SqlConnection cn = new SqlConnection(SqlServerConnection(server, "master")))
            {
                cn.Open();
                bool old = false;
                try
                {
                    using (SqlCommand cmd = new SqlCommand("SELECT name FROM sys.databases WHERE database_id > 4 AND state = 0 ORDER BY name", cn))
                    using (SqlDataReader r = cmd.ExecuteReader())
                    {
                        while (r.Read()) list.Add(r.GetString(0));
                    }
                }
                catch
                {
                    old = true;
                }
                if (old)
                {
                    // SQL Server 2000 没有 sys.databases
                    using (SqlCommand cmd = new SqlCommand("SELECT name FROM master..sysdatabases WHERE dbid > 4 ORDER BY name", cn))
                    using (SqlDataReader r = cmd.ExecuteReader())
                    {
                        while (r.Read()) list.Add(r.GetString(0));
                    }
                }
            }
            return list;
        }
    }
}
