using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;

namespace DrugFlow
{
    /// <summary>
    /// SQLite 数据源，不依赖任何第三方库：
    /// 优先用 Windows 自带的 winsqlite3.dll（Win10/11），
    /// 其次用软件目录里的 sqlite3.dll（老系统）。
    /// 注意：这里必须用 DllImport 直接声明，不能用 GetDelegateForFunctionPointer，
    /// 后者在 32 位 + WinForms 下会破坏栈，导致创建窗口时进程崩溃。
    /// </summary>
    public class SqliteProvider : IProvider
    {
        // ---------------- 原生声明 ----------------
        private static class W
        {
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_libversion();
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_open_v2(byte[] filename, out IntPtr db, int flags, IntPtr vfs);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_close(IntPtr db);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int nByte, out IntPtr stmt, out IntPtr tail);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_step(IntPtr stmt);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_finalize(IntPtr stmt);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_column_count(IntPtr stmt);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_column_name(IntPtr stmt, int i);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_column_type(IntPtr stmt, int i);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern long sqlite3_column_int64(IntPtr stmt, int i);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern double sqlite3_column_double(IntPtr stmt, int i);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_column_text(IntPtr stmt, int i);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_column_bytes(IntPtr stmt, int i);
            [DllImport("winsqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_errmsg(IntPtr db);
        }

        private static class S
        {
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_libversion();
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_open_v2(byte[] filename, out IntPtr db, int flags, IntPtr vfs);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_close(IntPtr db);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int nByte, out IntPtr stmt, out IntPtr tail);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_step(IntPtr stmt);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_finalize(IntPtr stmt);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_column_count(IntPtr stmt);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_column_name(IntPtr stmt, int i);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_column_type(IntPtr stmt, int i);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern long sqlite3_column_int64(IntPtr stmt, int i);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern double sqlite3_column_double(IntPtr stmt, int i);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_column_text(IntPtr stmt, int i);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern int sqlite3_column_bytes(IntPtr stmt, int i);
            [DllImport("sqlite3.dll", CallingConvention = CallingConvention.Cdecl)]
            internal static extern IntPtr sqlite3_errmsg(IntPtr db);
        }

        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool SetDllDirectory(string path);

        private static int _impl = -1;   // -1 未检测，0=winsqlite3，1=sqlite3.dll
        private static string _loadError = "";
        private static string _loadedFrom = "";
        private static List<string> _extraDirs = new List<string>();

        public static void AddSearchDir(string dir)
        {
            if (!string.IsNullOrEmpty(dir) && !_extraDirs.Contains(dir)) _extraDirs.Add(dir);
        }

        public static string LoadedFrom { get { return _loadedFrom; } }

        private static bool EnsureImpl()
        {
            if (_impl >= 0) return true;
            if (_impl == -2) return false;

            try
            {
                W.sqlite3_libversion();
                _impl = 0;
                _loadedFrom = "winsqlite3.dll";
                return true;
            }
            catch { }

            List<string> dirs = new List<string>();
            dirs.Add(AppDomain.CurrentDomain.BaseDirectory);
            for (int i = 0; i < _extraDirs.Count; i++) dirs.Add(_extraDirs[i]);

            for (int i = 0; i < dirs.Count; i++)
            {
                if (string.IsNullOrEmpty(dirs[i])) continue;
                try
                {
                    SetDllDirectory(dirs[i]);
                    S.sqlite3_libversion();
                    _impl = 1;
                    _loadedFrom = System.IO.Path.Combine(dirs[i], "sqlite3.dll");
                    return true;
                }
                catch { }
                finally { try { SetDllDirectory(null); } catch { } }
            }

            _impl = -2;
            _loadError = "没有找到 SQLite 运行库（Windows 10/11 自带 winsqlite3.dll；老系统需要软件目录里的 sqlite3.dll）";
            return false;
        }

        // ---------------- 分发 ----------------
        private static int Open(byte[] f, out IntPtr db, int flags)
        {
            if (_impl == 0) return W.sqlite3_open_v2(f, out db, flags, IntPtr.Zero);
            return S.sqlite3_open_v2(f, out db, flags, IntPtr.Zero);
        }
        private static int Close(IntPtr db) { return _impl == 0 ? W.sqlite3_close(db) : S.sqlite3_close(db); }
        private static int Prepare(IntPtr db, byte[] sql, out IntPtr stmt, out IntPtr tail)
        {
            return _impl == 0 ? W.sqlite3_prepare_v2(db, sql, -1, out stmt, out tail) : S.sqlite3_prepare_v2(db, sql, -1, out stmt, out tail);
        }
        private static int Step(IntPtr st) { return _impl == 0 ? W.sqlite3_step(st) : S.sqlite3_step(st); }
        private static int Finalize(IntPtr st) { return _impl == 0 ? W.sqlite3_finalize(st) : S.sqlite3_finalize(st); }
        private static int ColCount(IntPtr st) { return _impl == 0 ? W.sqlite3_column_count(st) : S.sqlite3_column_count(st); }
        private static IntPtr ColName(IntPtr st, int i) { return _impl == 0 ? W.sqlite3_column_name(st, i) : S.sqlite3_column_name(st, i); }
        private static int ColType(IntPtr st, int i) { return _impl == 0 ? W.sqlite3_column_type(st, i) : S.sqlite3_column_type(st, i); }
        private static long ColInt(IntPtr st, int i) { return _impl == 0 ? W.sqlite3_column_int64(st, i) : S.sqlite3_column_int64(st, i); }
        private static double ColDouble(IntPtr st, int i) { return _impl == 0 ? W.sqlite3_column_double(st, i) : S.sqlite3_column_double(st, i); }
        private static IntPtr ColText(IntPtr st, int i) { return _impl == 0 ? W.sqlite3_column_text(st, i) : S.sqlite3_column_text(st, i); }
        private static int ColBytes(IntPtr st, int i) { return _impl == 0 ? W.sqlite3_column_bytes(st, i) : S.sqlite3_column_bytes(st, i); }
        private static IntPtr ErrMsgPtr(IntPtr db) { return _impl == 0 ? W.sqlite3_errmsg(db) : S.sqlite3_errmsg(db); }

        private static string Utf8Z(IntPtr p, int length)
        {
            if (p == IntPtr.Zero) return "";
            if (length <= 0)
            {
                int n = 0;
                while (n < 1 << 20 && Marshal.ReadByte(p, n) != 0) n++;
                length = n;
            }
            if (length == 0) return "";
            byte[] b = new byte[length];
            Marshal.Copy(p, b, 0, length);
            return Encoding.UTF8.GetString(b);
        }

        private IntPtr _db;

        public SqliteProvider(string path)
        {
            try
            {
                string d = System.IO.Path.GetDirectoryName(path);
                AddSearchDir(d);
                if (!string.IsNullOrEmpty(d)) AddSearchDir(System.IO.Path.GetDirectoryName(d));
            }
            catch { }

            if (!EnsureImpl()) throw new Exception(_loadError);

            byte[] fn = Encoding.UTF8.GetBytes(path + "\0");
            IntPtr db;
            int rc = Open(fn, out db, 1); // READONLY
            if (rc != 0)
            {
                rc = Open(fn, out db, 2); // READWRITE
            }
            if (rc != 0) throw new Exception("打不开 SQLite 数据库（错误码 " + rc + "）：" + path);
            _db = db;
        }

        public string Kind { get { return "sqlite"; } }

        private string ErrMsg()
        {
            try { return Utf8Z(ErrMsgPtr(_db), 0); }
            catch { return "未知错误"; }
        }

        public List<string> ListTables()
        {
            List<string> list = new List<string>();
            DataTable dt = Query("SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name");
            foreach (DataRow r in dt.Rows)
            {
                string n = Convert.ToString(r[0]);
                if (!string.IsNullOrEmpty(n)) list.Add(n);
            }
            return list;
        }

        public List<ColumnInfo> ListColumns(string table)
        {
            List<ColumnInfo> list = new List<ColumnInfo>();
            DataTable dt = Query("PRAGMA table_info(" + Quote(table) + ")");
            foreach (DataRow r in dt.Rows)
            {
                string n = Convert.ToString(r["name"]);
                string t = r.Table.Columns.Contains("type") ? Convert.ToString(r["type"]) : "";
                list.Add(new ColumnInfo(n, t));
            }
            return list;
        }

        public DataTable Query(string sql)
        {
            DataTable dt = new DataTable();
            IntPtr stmt, tail;
            byte[] zsql = Encoding.UTF8.GetBytes(sql + "\0");
            int rc = Prepare(_db, zsql, out stmt, out tail);
            if (rc != 0) throw new Exception("SQL 错误：" + ErrMsg() + "\n" + sql);

            try
            {
                int n = ColCount(stmt);
                for (int i = 0; i < n; i++)
                {
                    string name = Utf8Z(ColName(stmt, i), 0);
                    if (string.IsNullOrEmpty(name)) name = "列" + (i + 1);
                    string uniq = name;
                    int k = 2;
                    while (dt.Columns.Contains(uniq)) { uniq = name + "_" + k; k++; }
                    dt.Columns.Add(uniq, typeof(object));
                }

                while ((rc = Step(stmt)) == 100)
                {
                    object[] vals = new object[n];
                    for (int i = 0; i < n; i++)
                    {
                        int t = ColType(stmt, i);
                        if (t == 1) vals[i] = ColInt(stmt, i);
                        else if (t == 2) vals[i] = ColDouble(stmt, i);
                        else if (t == 3) vals[i] = Utf8Z(ColText(stmt, i), ColBytes(stmt, i));
                        else if (t == 4) vals[i] = "(二进制)";
                        else vals[i] = DBNull.Value;
                    }
                    dt.Rows.Add(vals);
                }
                if (rc != 101) throw new Exception("读取失败：" + ErrMsg());
            }
            finally
            {
                try { Finalize(stmt); } catch { }
            }
            return dt;
        }

        public string Quote(string name) { return "\"" + name.Replace("\"", "\"\"") + "\""; }
        public string QuoteTable(string name) { return Quote(name); }
        public string DateLiteral(DateTime d) { return DateLiteral(d, "auto"); }

        public string DateLiteral(DateTime d, string style)
        {
            if (style == "ymd8") return "'" + d.ToString("yyyyMMdd") + "'";
            if (style == "slash") return "'" + d.ToString("yyyy/MM/dd") + "'";
            if (style == "ymd") return "'" + d.ToString("yyyy-MM-dd") + "'";
            return "'" + d.ToString("yyyy-MM-dd HH:mm:ss") + "'";
        }

        public void Dispose()
        {
            if (_db != IntPtr.Zero)
            {
                try { Close(_db); } catch { }
                _db = IntPtr.Zero;
            }
        }
    }
}
