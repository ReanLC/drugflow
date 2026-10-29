using System;
using System.Collections.Generic;
using System.Data;

namespace DrugFlow
{
    /// <summary>数据库列信息</summary>
    public class ColumnInfo
    {
        public string Name;
        public string Type;

        public ColumnInfo() { }
        public ColumnInfo(string name, string type) { Name = name; Type = type; }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(Type)) return Name;
            return Name + "  (" + Type + ")";
        }
    }

    /// <summary>字段映射：哪一列是日期、药品、数量……</summary>
    public class FieldMap
    {
        public string Date;
        public string Drug;
        public string Spec;
        public string Qty;
        public string Price;
        public string Amount;
        public string Unit;
        public string Bill;
    }

    /// <summary>明细表 + 主表的关联方式（日期在主表里时用）</summary>
    public class JoinSpec
    {
        public string MasterTable;
        public string DetailKey;
        public string MasterKey;
        public string DateColumn;   // 主表里的日期字段
        public const string DateAlias = "__销售日期";

        public bool IsValid
        {
            get
            {
                return !string.IsNullOrEmpty(MasterTable) && !string.IsNullOrEmpty(DetailKey)
                    && !string.IsNullOrEmpty(MasterKey) && !string.IsNullOrEmpty(DateColumn);
            }
        }
    }

    /// <summary>一个可连接的数据源</summary>
    public class SourceItem
    {
        public string Kind;        // sqlserver / access / odbc / demo / profile
        public string Title;       // 显示名称
        public string Connection;  // 连接字符串或文件路径
        public string Database;
        public string Path;
        public Profile Profile;    // Kind == "profile" 时使用

        public override string ToString() { return Title; }
    }

    /// <summary>扫描发现的一条线索</summary>
    public class Finding
    {
        public string Kind;
        public string Title;
        public string Detail;
        public string Path;
    }

    /// <summary>扫描结果</summary>
    public class ScanResult
    {
        public List<Finding> Findings = new List<Finding>();
        public List<SourceItem> Sources = new List<SourceItem>();
        public List<string> CandidateDirs = new List<string>();
        public List<string> ProcessDump = new List<string>();
        public List<string> Notes = new List<string>();
        public List<string> Schema = new List<string>();
        public DateTime StartedAt = DateTime.Now;
    }

    /// <summary>数据访问抽象</summary>
    public interface IProvider : IDisposable
    {
        string Kind { get; }
        List<string> ListTables();
        List<ColumnInfo> ListColumns(string table);
        DataTable Query(string sql);
        string Quote(string name);
        string QuoteTable(string name);
        string DateLiteral(DateTime d);
        /// <summary>style: auto / ymd / ymd8 / slash</summary>
        string DateLiteral(DateTime d, string style);
    }

    public static class Tx
    {
        public static string Nz(string s) { return s == null ? "" : s; }

        public static bool Has(string s, string sub)
        {
            return s != null && sub != null && s.IndexOf(sub, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool HasAny(string s, string[] subs)
        {
            if (s == null || subs == null) return false;
            for (int i = 0; i < subs.Length; i++)
            {
                if (Has(s, subs[i])) return true;
            }
            return false;
        }

        public static string Short(string s, int n)
        {
            if (string.IsNullOrEmpty(s)) return "";
            s = s.Replace("\r", " ").Replace("\n", " ").Trim();
            if (s.Length <= n) return s;
            return s.Substring(0, n) + "…";
        }

        public static string Join(string[] arr, string sep)
        {
            if (arr == null || arr.Length == 0) return "";
            return string.Join(sep, arr);
        }

        /// <summary>把连接串里的密码打码，报告里可以安全发送</summary>
        public static string MaskSecret(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return System.Text.RegularExpressions.Regex.Replace(s, @"(?i)\b(password|pwd)\s*=\s*[^;]*", "$1=***");
        }
    }
}
