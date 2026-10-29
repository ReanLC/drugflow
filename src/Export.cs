using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text;

namespace DrugFlow
{
    public class DetailRow
    {
        public DateTime? Time;
        public string Bill = "";
        public string Drug = "";
        public string Spec = "";
        public string Unit = "";
        public double Qty;
        public double Price;
        public double Amount;
    }

    public class ExportResult
    {
        public int RowCount;
        public int DrugCount;
        public double TotalAmount;
        public double TotalQty;
        public string Path;
        public DateTime From;
        public DateTime To;
    }

    public static class Exporter
    {
        public static string SuggestFileName(DateTime from, DateTime to)
        {
            return "药品销售流水_" + from.ToString("yyyyMMdd") + "-" + to.ToString("yyyyMMdd") + ".xlsx";
        }

        public static List<DetailRow> Map(DataTable dt, FieldMap map)
        {
            List<DetailRow> list = new List<DetailRow>(dt.Rows.Count);
            foreach (DataRow r in dt.Rows)
            {
                DetailRow d = new DetailRow();
                d.Time = GetDate(r, map.Date);
                d.Bill = GetStr(r, map.Bill);
                d.Drug = GetStr(r, map.Drug);
                d.Spec = GetStr(r, map.Spec);
                d.Unit = GetStr(r, map.Unit);
                d.Qty = GetNum(r, map.Qty);
                d.Price = GetNum(r, map.Price);
                d.Amount = GetNum(r, map.Amount);
                if (string.IsNullOrEmpty(map.Amount) || double.IsNaN(d.Amount))
                {
                    d.Amount = Math.Round(d.Qty * d.Price, 2);
                }
                if (string.IsNullOrEmpty(map.Qty)) d.Qty = 1;
                list.Add(d);
            }
            list.Sort(delegate(DetailRow a, DetailRow b)
            {
                DateTime ta = a.Time.HasValue ? a.Time.Value : DateTime.MinValue;
                DateTime tb = b.Time.HasValue ? b.Time.Value : DateTime.MinValue;
                int c = ta.CompareTo(tb);
                if (c != 0) return c;
                return string.Compare(a.Bill, b.Bill, StringComparison.OrdinalIgnoreCase);
            });
            return list;
        }

        public static ExportResult Export(DataTable dt, FieldMap map, DateTime from, DateTime to, string path)
        {
            List<DetailRow> rows = Map(dt, map);

            XlsxSheet detail = new XlsxSheet();
            detail.Name = "销售明细";
            detail.Headers = new string[] { "序号", "销售日期", "单据号", "药品名称", "规格", "单位", "数量", "单价", "金额" };
            detail.Widths = new double[] { 6, 19, 20, 32, 20, 8, 10, 10, 12 };
            detail.Formats = new string[] { "int", "datetime", "text", "text", "text", "text", "number", "money", "money" };
            detail.FreezeHeader = true;
            detail.AutoFilter = true;
            detail.Rows = new List<object[]>();
            double total = 0, totalQty = 0;
            for (int i = 0; i < rows.Count; i++)
            {
                DetailRow d = rows[i];
                total += d.Amount;
                totalQty += d.Qty;
                detail.Rows.Add(new object[]
                {
                    i + 1,
                    d.Time.HasValue ? (object)d.Time.Value : "",
                    d.Bill, d.Drug, d.Spec, d.Unit,
                    d.Qty, d.Price, d.Amount
                });
            }

            XlsxSheet sum = new XlsxSheet();
            sum.Name = "按药品汇总";
            sum.Headers = new string[] { "药品名称", "规格", "单位", "销售数量", "销售金额", "笔数" };
            sum.Widths = new double[] { 34, 20, 8, 12, 14, 8 };
            sum.Formats = new string[] { "text", "text", "text", "number", "money", "int" };
            sum.FreezeHeader = true;
            sum.AutoFilter = true;
            sum.BoldLastRow = true;
            sum.Rows = new List<object[]>();

            Dictionary<string, double[]> agg = new Dictionary<string, double[]>();
            Dictionary<string, string[]> keyText = new Dictionary<string, string[]>();
            for (int i = 0; i < rows.Count; i++)
            {
                DetailRow d = rows[i];
                string key = d.Drug + "\u0001" + d.Spec + "\u0001" + d.Unit;
                double[] v;
                if (!agg.TryGetValue(key, out v))
                {
                    v = new double[3];
                    agg[key] = v;
                    keyText[key] = new string[] { d.Drug, d.Spec, d.Unit };
                }
                v[0] += d.Qty;
                v[1] += d.Amount;
                v[2] += 1;
            }

            List<string> keys = new List<string>(agg.Keys);
            keys.Sort(delegate(string a, string b) { return agg[b][1].CompareTo(agg[a][1]); });
            int drugCount = 0;
            for (int i = 0; i < keys.Count; i++)
            {
                string[] t = keyText[keys[i]];
                double[] v = agg[keys[i]];
                sum.Rows.Add(new object[] { t[0], t[1], t[2], v[0], Math.Round(v[1], 2), (int)v[2] });
                drugCount++;
            }
            sum.Rows.Add(new object[] { "合计", "", "", totalQty, Math.Round(total, 2), rows.Count });

            List<XlsxSheet> sheets = new List<XlsxSheet>();
            sheets.Add(detail);
            sheets.Add(sum);
            XlsxWriter.Write(path, sheets);

            ExportResult res = new ExportResult();
            res.RowCount = rows.Count;
            res.DrugCount = drugCount;
            res.TotalAmount = Math.Round(total, 2);
            res.TotalQty = totalQty;
            res.Path = path;
            res.From = from;
            res.To = to;
            return res;
        }

        /// <summary>明细表 + 主表关联查询：日期、单据号通常在主表里</summary>
        public static string BuildSqlJoin(IProvider p, string detailTable, JoinSpec join,
            DateTime from, DateTime to, string dateStyle)
        {
            return BuildSqlJoin(p, detailTable, null, join, from, to, dateStyle);
        }

        /// <summary>
        /// 同上，但把明细表的列逐个列出来（不用 d.*）。
        /// Access 对“别名.*”支持不好，而且列全写出来可以避免重名列。
        /// </summary>
        public static string BuildSqlJoin(IProvider p, string detailTable, List<ColumnInfo> detailColumns,
            JoinSpec join, DateTime from, DateTime to, string dateStyle)
        {
            StringBuilder cols = new StringBuilder();
            if (detailColumns != null)
            {
                for (int i = 0; i < detailColumns.Count; i++)
                {
                    string name = detailColumns[i].Name;
                    if (string.IsNullOrEmpty(name)) continue;
                    if (string.Equals(name, JoinSpec.DateAlias, StringComparison.OrdinalIgnoreCase)) continue;
                    if (cols.Length > 0) cols.Append(", ");
                    cols.Append("d.").Append(p.Quote(name));
                }
            }
            if (cols.Length == 0) cols.Append("d.*");

            string sql = "SELECT " + cols.ToString() + ", m." + p.Quote(join.DateColumn) +
                         " AS " + p.Quote(JoinSpec.DateAlias) +
                         " FROM " + p.QuoteTable(detailTable) + " AS d" +
                         " LEFT JOIN " + p.QuoteTable(join.MasterTable) + " AS m" +
                         " ON d." + p.Quote(join.DetailKey) + " = m." + p.Quote(join.MasterKey) +
                         " WHERE m." + p.Quote(join.DateColumn) + " >= " + p.DateLiteral(from, dateStyle) +
                         " AND m." + p.Quote(join.DateColumn) + " < " + p.DateLiteral(to.AddDays(1), dateStyle) +
                         " ORDER BY m." + p.Quote(join.DateColumn);
            return sql;
        }

        public static string BuildSql(IProvider p, string table, FieldMap map, DateTime from, DateTime to)
        {
            return BuildSql(p, table, map, from, to, "auto");
        }

        public static string BuildSql(IProvider p, string table, FieldMap map, DateTime from, DateTime to, string dateStyle)
        {
            string sql = "SELECT * FROM " + p.QuoteTable(table);
            if (!string.IsNullOrEmpty(map.Date))
            {
                sql += " WHERE " + p.Quote(map.Date) + " >= " + p.DateLiteral(from, dateStyle) +
                       " AND " + p.Quote(map.Date) + " < " + p.DateLiteral(to.AddDays(1), dateStyle);
            }
            if (!string.IsNullOrEmpty(map.Date))
            {
                sql += " ORDER BY " + p.Quote(map.Date);
            }
            return sql;
        }

        // ---------------- 取值辅助 ----------------
        public static string GetStr(DataRow r, string col)
        {
            if (string.IsNullOrEmpty(col) || !r.Table.Columns.Contains(col)) return "";
            object o = r[col];
            if (o == null || o == DBNull.Value) return "";
            if (o is DateTime) return ((DateTime)o).ToString("yyyy-MM-dd HH:mm:ss");
            return Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
        }

        public static double GetNum(DataRow r, string col)
        {
            if (string.IsNullOrEmpty(col) || !r.Table.Columns.Contains(col)) return 0;
            object o = r[col];
            if (o == null || o == DBNull.Value) return 0;
            try
            {
                if (o is double) return (double)o;
                if (o is decimal) return (double)(decimal)o;
                if (o is int) return (int)o;
                if (o is long) return (long)o;
                if (o is float) return (float)o;
                string s = Convert.ToString(o, CultureInfo.InvariantCulture).Replace(",", "").Replace("￥", "").Replace("¥", "").Trim();
                double d;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d)) return d;
            }
            catch { }
            return 0;
        }

        public static DateTime? GetDate(DataRow r, string col)
        {
            if (string.IsNullOrEmpty(col) || !r.Table.Columns.Contains(col)) return null;
            object o = r[col];
            if (o == null || o == DBNull.Value) return null;
            if (o is DateTime) return (DateTime)o;
            string s = Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
            if (s.Length == 0) return null;
            DateTime d;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out d)) return d;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out d)) return d;
            return null;
        }
    }
}
