using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace DrugFlow
{
    /// <summary>
    /// 一个工作表的描述。Rows 中每个元素可以是 string / double / int / long /
    /// decimal / DateTime / null（也接受其它常见数值类型）。
    /// </summary>
    public class XlsxSheet
    {
        public string Name;
        public string[] Headers;
        public double[] Widths;
        public string[] Formats;
        public List<object[]> Rows = new List<object[]>();
        public bool FreezeHeader;
        public bool AutoFilter;
        public bool BoldLastRow;
    }

    /// <summary>
    /// 零依赖的 .xlsx 生成器：自己拼 ZIP（DeflateStream + CRC32），
    /// 字符串用 inlineStr，不需要共享字符串表。
    /// 仅使用 .NET Framework 自带程序集，兼容 C# 5 编译器。
    /// </summary>
    public static class XlsxWriter
    {
        private const string XmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
        private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string NsContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

        // cellXfs 中的下标，必须和 Styles() 中的顺序一致。
        private const int StyleGeneralText = 0;
        private const int StyleHeader = 1;
        private const int StyleText = 2;
        private const int StyleGeneralNumber = 3;
        private const int StyleInt = 4;
        private const int StyleMoney = 5;
        private const int StyleNumber = 6;
        private const int StyleDate = 7;
        private const int StyleDateTime = 8;
        private const int StyleBoldText = 9;
        private const int StyleBoldGeneralNumber = 10;
        private const int StyleBoldInt = 11;
        private const int StyleBoldMoney = 12;
        private const int StyleBoldNumber = 13;
        private const int StyleBoldDate = 14;
        private const int StyleBoldDateTime = 15;

        private const int MaxColumns = 16384;
        private const int MaxCellText = 32767;

        private static readonly DateTime ExcelEpoch = new DateTime(1899, 12, 30);
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        /// <summary>生成 xlsx 文件。path 的父目录不存在时会自动创建。</summary>
        public static void Write(string path, List<XlsxSheet> sheets)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("path 不能为空。", "path");
            if (sheets == null || sheets.Count == 0)
                throw new ArgumentException("至少需要一个工作表。", "sheets");

            string[] sheetNames = BuildSheetNames(sheets);

            string fullPath = Path.GetFullPath(path);
            string dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            try
            {
                using (FileStream fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    ZipWriter zip = new ZipWriter(fs);

                    zip.Add("[Content_Types].xml", Utf8(ContentTypes(sheets.Count)));
                    zip.Add("_rels/.rels", Utf8(RootRels()));
                    zip.Add("xl/workbook.xml", Utf8(Workbook(sheetNames)));
                    zip.Add("xl/_rels/workbook.xml.rels", Utf8(WorkbookRels(sheets.Count)));
                    zip.Add("xl/styles.xml", Utf8(Styles()));

                    for (int i = 0; i < sheets.Count; i++)
                    {
                        XlsxSheet sheet = sheets[i];
                        if (sheet == null) sheet = new XlsxSheet();
                        zip.Add("xl/worksheets/sheet" + Num(i + 1) + ".xml",
                                Utf8(Worksheet(sheet, i == 0)));
                    }

                    zip.Finish();
                }
            }
            catch
            {
                // 不留下只写了一半的坏文件。
                try
                {
                    if (File.Exists(fullPath)) File.Delete(fullPath);
                }
                catch
                {
                    // 删除失败时保留原始异常。
                }
                throw;
            }
        }

        // ------------------------------------------------------------------
        // 包结构
        // ------------------------------------------------------------------

        private static string ContentTypes(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlHeader);
            sb.Append("<Types xmlns=\"").Append(NsContentTypes).Append("\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 0; i < sheetCount; i++)
            {
                sb.Append("<Override PartName=\"/xl/worksheets/sheet").Append(Num(i + 1))
                  .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        private static string RootRels()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlHeader);
            sb.Append("<Relationships xmlns=\"").Append(NsPkgRel).Append("\">");
            sb.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string Workbook(string[] sheetNames)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlHeader);
            sb.Append("<workbook xmlns=\"").Append(NsMain).Append("\" xmlns:r=\"").Append(NsRel).Append("\">");
            sb.Append("<bookViews><workbookView activeTab=\"0\"/></bookViews>");
            sb.Append("<sheets>");
            for (int i = 0; i < sheetNames.Length; i++)
            {
                sb.Append("<sheet name=\"").Append(Esc(sheetNames[i]))
                  .Append("\" sheetId=\"").Append(Num(i + 1))
                  .Append("\" r:id=\"rId").Append(Num(i + 1)).Append("\"/>");
            }
            sb.Append("</sheets>");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        private static string WorkbookRels(int sheetCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlHeader);
            sb.Append("<Relationships xmlns=\"").Append(NsPkgRel).Append("\">");
            for (int i = 0; i < sheetCount; i++)
            {
                sb.Append("<Relationship Id=\"rId").Append(Num(i + 1))
                  .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet")
                  .Append(Num(i + 1)).Append(".xml\"/>");
            }
            sb.Append("<Relationship Id=\"rId").Append(Num(sheetCount + 1))
              .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private static string Styles()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(XmlHeader);
            sb.Append("<styleSheet xmlns=\"").Append(NsMain).Append("\">");

            sb.Append("<numFmts count=\"4\">");
            sb.Append("<numFmt numFmtId=\"164\" formatCode=\"#,##0.00\"/>");
            sb.Append("<numFmt numFmtId=\"165\" formatCode=\"#,##0.###\"/>");
            sb.Append("<numFmt numFmtId=\"166\" formatCode=\"yyyy-mm-dd\"/>");
            sb.Append("<numFmt numFmtId=\"167\" formatCode=\"yyyy-mm-dd hh:mm:ss\"/>");
            sb.Append("</numFmts>");

            sb.Append("<fonts count=\"2\">");
            sb.Append("<font><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><family val=\"2\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><color rgb=\"FF000000\"/><name val=\"Calibri\"/><family val=\"2\"/></font>");
            sb.Append("</fonts>");

            sb.Append("<fills count=\"3\">");
            sb.Append("<fill><patternFill patternType=\"none\"/></fill>");
            sb.Append("<fill><patternFill patternType=\"gray125\"/></fill>");
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFDCE6F1\"/><bgColor indexed=\"64\"/></patternFill></fill>");
            sb.Append("</fills>");

            sb.Append("<borders count=\"2\">");
            sb.Append("<border><left/><right/><top/><bottom/><diagonal/></border>");
            sb.Append("<border>");
            sb.Append("<left style=\"thin\"><color rgb=\"FFB0B0B0\"/></left>");
            sb.Append("<right style=\"thin\"><color rgb=\"FFB0B0B0\"/></right>");
            sb.Append("<top style=\"thin\"><color rgb=\"FFB0B0B0\"/></top>");
            sb.Append("<bottom style=\"thin\"><color rgb=\"FFB0B0B0\"/></bottom>");
            sb.Append("<diagonal/>");
            sb.Append("</border>");
            sb.Append("</borders>");

            sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            sb.Append("<cellXfs count=\"16\">");

            // 0: 普通文本（左对齐）
            AppendXf(sb, 0, 0, 0, 0, false, false, false, false, "left");
            // 1: 表头（加粗、浅蓝底、细边框、居中、不换行）
            AppendXf(sb, 0, 1, 2, 1, false, true, true, true, "center");
            // 2: 显式 text，numFmt=@
            AppendXf(sb, 49, 0, 0, 0, true, false, false, false, "left");
            // 3: 未指定格式的数字（General，右对齐）
            AppendXf(sb, 0, 0, 0, 0, false, false, false, false, "right");
            // 4: int，numFmt=0（内置 1）
            AppendXf(sb, 1, 0, 0, 0, true, false, false, false, "right");
            // 5: money
            AppendXf(sb, 164, 0, 0, 0, true, false, false, false, "right");
            // 6: number
            AppendXf(sb, 165, 0, 0, 0, true, false, false, false, "right");
            // 7: date
            AppendXf(sb, 166, 0, 0, 0, true, false, false, false, "center");
            // 8: datetime
            AppendXf(sb, 167, 0, 0, 0, true, false, false, false, "center");

            // 9-15: 合计行 / 最后一行加粗版本
            AppendXf(sb, 49, 1, 0, 0, true, true, false, false, "left");
            AppendXf(sb, 0, 1, 0, 0, false, true, false, false, "right");
            AppendXf(sb, 1, 1, 0, 0, true, true, false, false, "right");
            AppendXf(sb, 164, 1, 0, 0, true, true, false, false, "right");
            AppendXf(sb, 165, 1, 0, 0, true, true, false, false, "right");
            AppendXf(sb, 166, 1, 0, 0, true, true, false, false, "center");
            AppendXf(sb, 167, 1, 0, 0, true, true, false, false, "center");

            sb.Append("</cellXfs>");
            sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
            sb.Append("<dxfs count=\"0\"/>");
            sb.Append("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>");
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        private static void AppendXf(StringBuilder sb, int numFmtId, int fontId, int fillId,
                                     int borderId, bool applyNumberFormat, bool applyFont,
                                     bool applyFill, bool applyBorder, string horizontal)
        {
            sb.Append("<xf numFmtId=\"").Append(Num(numFmtId))
              .Append("\" fontId=\"").Append(Num(fontId))
              .Append("\" fillId=\"").Append(Num(fillId))
              .Append("\" borderId=\"").Append(Num(borderId))
              .Append("\" xfId=\"0\"");
            if (applyNumberFormat) sb.Append(" applyNumberFormat=\"1\"");
            if (applyFont) sb.Append(" applyFont=\"1\"");
            if (applyFill) sb.Append(" applyFill=\"1\"");
            if (applyBorder) sb.Append(" applyBorder=\"1\"");
            sb.Append(" applyAlignment=\"1\"><alignment horizontal=\"").Append(horizontal)
              .Append("\" vertical=\"center\" wrapText=\"0\"/></xf>");
        }

        // ------------------------------------------------------------------
        // 工作表
        // ------------------------------------------------------------------

        private static string Worksheet(XlsxSheet sheet, bool firstSheet)
        {
            int colCount = GetColumnCount(sheet);
            int lastCol = colCount > 0 ? colCount : 1;
            bool hasHeader = sheet.Headers != null && sheet.Headers.Length > 0;

            int dataCount = sheet.Rows == null ? 0 : sheet.Rows.Count;
            int lastRow = (hasHeader ? 1 : 0) + dataCount;
            if (lastRow < 1) lastRow = 1;

            StringBuilder sb = new StringBuilder(4096);
            sb.Append(XmlHeader);
            sb.Append("<worksheet xmlns=\"").Append(NsMain).Append("\" xmlns:r=\"").Append(NsRel).Append("\">");

            if (lastRow == 1 && lastCol == 1)
                sb.Append("<dimension ref=\"A1\"/>");
            else
                sb.Append("<dimension ref=\"A1:").Append(CellRef(lastCol, lastRow)).Append("\"/>");

            sb.Append("<sheetViews><sheetView");
            if (firstSheet) sb.Append(" tabSelected=\"1\"");
            sb.Append(" workbookViewId=\"0\">");
            if (sheet.FreezeHeader)
            {
                sb.Append("<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>");
                sb.Append("<selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/>");
            }
            else
            {
                sb.Append("<selection activeCell=\"A1\" sqref=\"A1\"/>");
            }
            sb.Append("</sheetView></sheetViews>");

            sb.Append("<sheetFormatPr defaultRowHeight=\"15\"/>");

            AppendCols(sb, sheet, colCount);

            sb.Append("<sheetData>");

            if (hasHeader)
            {
                sb.Append("<row r=\"1\" ht=\"22\" customHeight=\"1\">");
                for (int c = 0; c < colCount; c++)
                {
                    string text = (sheet.Headers != null && c < sheet.Headers.Length && sheet.Headers[c] != null)
                        ? sheet.Headers[c] : "";
                    AppendTextCell(sb, CellRef(c + 1, 1), StyleHeader, text);
                }
                sb.Append("</row>");
            }

            int rowNumber = hasHeader ? 2 : 1;
            for (int r = 0; r < dataCount; r++)
            {
                object[] row = sheet.Rows[r];
                bool bold = sheet.BoldLastRow && r == dataCount - 1;

                sb.Append("<row r=\"").Append(Num(rowNumber)).Append("\">");
                if (row != null)
                {
                    int cellCount = row.Length;
                    if (cellCount > MaxColumns) cellCount = MaxColumns;
                    for (int c = 0; c < cellCount; c++)
                    {
                        object value = row[c];
                        if (value == null || value is DBNull) continue;

                        string fmt = (sheet.Formats != null && c < sheet.Formats.Length)
                            ? sheet.Formats[c] : null;
                        int style = PickStyle(fmt, value, bold);
                        AppendValueCell(sb, CellRef(c + 1, rowNumber), style, value, fmt);
                    }
                }
                sb.Append("</row>");
                rowNumber++;
            }

            sb.Append("</sheetData>");

            if (sheet.AutoFilter && hasHeader)
            {
                sb.Append("<autoFilter ref=\"A1:").Append(CellRef(lastCol, lastRow)).Append("\"/>");
            }

            sb.Append("<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>");
            sb.Append("</worksheet>");
            return sb.ToString();
        }

        private static void AppendCols(StringBuilder sb, XlsxSheet sheet, int colCount)
        {
            if (sheet.Widths == null || sheet.Widths.Length == 0 || colCount <= 0) return;

            StringBuilder cols = new StringBuilder();
            int count = sheet.Widths.Length;
            if (count > colCount) count = colCount;
            for (int c = 0; c < count; c++)
            {
                double width = sheet.Widths[c];
                if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0) continue;
                if (width > 255.0) width = 255.0;
                cols.Append("<col min=\"").Append(Num(c + 1))
                    .Append("\" max=\"").Append(Num(c + 1))
                    .Append("\" width=\"").Append(WidthXml(width))
                    .Append("\" customWidth=\"1\"/>");
            }
            if (cols.Length > 0)
                sb.Append("<cols>").Append(cols.ToString()).Append("</cols>");
        }

        private static int GetColumnCount(XlsxSheet sheet)
        {
            int n = 0;
            if (sheet.Headers != null) n = Math.Max(n, sheet.Headers.Length);
            if (sheet.Widths != null) n = Math.Max(n, sheet.Widths.Length);
            if (sheet.Formats != null) n = Math.Max(n, sheet.Formats.Length);
            if (sheet.Rows != null)
            {
                for (int i = 0; i < sheet.Rows.Count; i++)
                {
                    object[] row = sheet.Rows[i];
                    if (row != null) n = Math.Max(n, row.Length);
                }
            }
            if (n > MaxColumns) n = MaxColumns;
            return n;
        }

        private static int PickStyle(string fmt, object value, bool bold)
        {
            string f = NormalizeFormat(fmt);

            if (f == "text") return bold ? StyleBoldText : StyleText;
            if (f == "int") return bold ? StyleBoldInt : StyleInt;
            if (f == "money") return bold ? StyleBoldMoney : StyleMoney;
            if (f == "number") return bold ? StyleBoldNumber : StyleNumber;
            if (f == "date") return bold ? StyleBoldDate : StyleDate;
            if (f == "datetime") return bold ? StyleBoldDateTime : StyleDateTime;

            // 未指定格式时按值本身判断；DateTime 自动用日期/日期时间格式，
            // 否则 Excel 会把它显示成 45306 这样的序列号。
            DateTime dt;
            if (value is DateTime)
            {
                dt = (DateTime)value;
                if (dt.TimeOfDay.Ticks == 0)
                    return bold ? StyleBoldDate : StyleDate;
                return bold ? StyleBoldDateTime : StyleDateTime;
            }

            if (IsNumber(value))
                return bold ? StyleBoldGeneralNumber : StyleGeneralNumber;

            return bold ? StyleBoldText : StyleGeneralText;
        }

        private static string NormalizeFormat(string fmt)
        {
            if (fmt == null) return null;
            return fmt.Trim().ToLowerInvariant();
        }

        // ------------------------------------------------------------------
        // 单元格
        // ------------------------------------------------------------------

        private static void AppendValueCell(StringBuilder sb, string cellRef, int style, object value, string fmt)
        {
            if (value == null || value is DBNull) return;

            string f = NormalizeFormat(fmt);

            if (value is string)
            {
                AppendTextCell(sb, cellRef, style, (string)value);
                return;
            }
            if (value is char)
            {
                AppendTextCell(sb, cellRef, style, value.ToString());
                return;
            }
            if (value is bool)
            {
                bool b = (bool)value;
                sb.Append("<c r=\"").Append(cellRef).Append("\" s=\"").Append(Num(style))
                  .Append("\" t=\"b\"><v>").Append(b ? "1" : "0").Append("</v></c>");
                return;
            }
            if (value is DateTime)
            {
                DateTime dt = (DateTime)value;
                if (f == "text")
                {
                    AppendTextCell(sb, cellRef, style, DateTimeText(dt));
                }
                else
                {
                    string serial = NumXml(DateSerial(dt));
                    AppendNumberCell(sb, cellRef, style, serial);
                }
                return;
            }
            if (IsNumber(value))
            {
                string number = NumberText(value);
                if (number != null)
                    AppendNumberCell(sb, cellRef, style, number);
                return;
            }

            // 兜底：未知类型按文本写入，绝不要写出 "null"。
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            AppendTextCell(sb, cellRef, style, text == null ? "" : text);
        }

        private static void AppendTextCell(StringBuilder sb, string cellRef, int style, string text)
        {
            string limited = LimitText(text);
            sb.Append("<c r=\"").Append(cellRef).Append("\" s=\"").Append(Num(style))
              .Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">")
              .Append(Esc(limited))
              .Append("</t></is></c>");
        }

        private static void AppendNumberCell(StringBuilder sb, string cellRef, int style, string number)
        {
            sb.Append("<c r=\"").Append(cellRef).Append("\" s=\"").Append(Num(style))
              .Append("\"><v>").Append(number).Append("</v></c>");
        }

        private static string NumberText(object value)
        {
            if (value is double)
            {
                double d = (double)value;
                if (double.IsNaN(d) || double.IsInfinity(d)) return null;
                return NumXml(d);
            }
            if (value is float)
            {
                float f = (float)value;
                if (float.IsNaN(f) || float.IsInfinity(f)) return null;
                return f.ToString("R", CultureInfo.InvariantCulture);
            }
            if (value is decimal)
                return ((decimal)value).ToString(CultureInfo.InvariantCulture);
            if (value is int)
                return ((int)value).ToString(CultureInfo.InvariantCulture);
            if (value is long)
                return ((long)value).ToString(CultureInfo.InvariantCulture);
            if (value is short)
                return ((short)value).ToString(CultureInfo.InvariantCulture);
            if (value is sbyte)
                return ((sbyte)value).ToString(CultureInfo.InvariantCulture);
            if (value is uint)
                return ((uint)value).ToString(CultureInfo.InvariantCulture);
            if (value is ulong)
                return ((ulong)value).ToString(CultureInfo.InvariantCulture);
            if (value is ushort)
                return ((ushort)value).ToString(CultureInfo.InvariantCulture);
            if (value is byte)
                return ((byte)value).ToString(CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static bool IsNumber(object value)
        {
            return value is double || value is float || value is decimal ||
                   value is int || value is long || value is short || value is sbyte ||
                   value is uint || value is ulong || value is ushort || value is byte;
        }

        private static double DateSerial(DateTime dt)
        {
            return (dt - ExcelEpoch).TotalDays;
        }

        private static string DateTimeText(DateTime dt)
        {
            if (dt.TimeOfDay.Ticks == 0)
                return dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static string NumXml(double d)
        {
            if (d == 0.0) return "0";
            if (Math.Abs(d) < 1e15 && d == Math.Truncate(d))
                return ((long)d).ToString(CultureInfo.InvariantCulture);
            return d.ToString("R", CultureInfo.InvariantCulture);
        }

        private static string WidthXml(double width)
        {
            return Math.Round(width, 2).ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static string LimitText(string text)
        {
            if (text == null) return "";
            if (text.Length <= MaxCellText) return text;
            return Truncate(text, MaxCellText);
        }

        // ------------------------------------------------------------------
        // 工作表名
        // ------------------------------------------------------------------

        private static string[] BuildSheetNames(List<XlsxSheet> sheets)
        {
            string[] names = new string[sheets.Count];
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < sheets.Count; i++)
            {
                string raw = sheets[i] == null ? null : sheets[i].Name;
                string name = CleanSheetName(raw);

                if (name.Length == 0)
                    name = "Sheet" + Num(i + 1);

                // Excel 把 History 作为保留名，避开它。
                if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase))
                    name = name + "_";

                name = Truncate(name, 31);
                // 截断后可能留下首尾空格或单引号，Excel 不允许。
                name = name.Trim().Trim('\'').Trim();
                if (name.Length == 0)
                    name = "Sheet" + Num(i + 1);

                if (used.Contains(name))
                {
                    string baseName = name;
                    int n = 2;
                    while (true)
                    {
                        string suffix = "(" + Num(n) + ")";
                        int keep = 31 - suffix.Length;
                        if (keep < 1) keep = 1;
                        string candidate = Truncate(baseName, keep) + suffix;
                        if (candidate.Length > 31)
                            candidate = Truncate(candidate, 31);
                        if (!used.Contains(candidate))
                        {
                            name = candidate;
                            break;
                        }
                        n++;
                    }
                }

                used.Add(name);
                names[i] = name;
            }
            return names;
        }

        private static string CleanSheetName(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";

            StringBuilder sb = new StringBuilder(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];
                if (c == ':' || c == '\\' || c == '/' || c == '?' ||
                    c == '*' || c == '[' || c == ']')
                    continue;
                if (c < 0x20 || c == 0x7F)
                    continue;
                sb.Append(c);
            }

            // 首尾空格和单引号在 Excel 工作表名里是非法的。
            string s = sb.ToString().Trim();
            s = s.Trim('\'').Trim();
            return s;
        }

        private static string Truncate(string s, int max)
        {
            if (string.IsNullOrEmpty(s) || max <= 0) return "";
            if (s.Length <= max) return s;

            int n = max;
            // 不要把代理对切成两半。
            if (n < s.Length && char.IsHighSurrogate(s[n - 1]) && char.IsLowSurrogate(s[n]))
                n--;
            if (n <= 0) return "";
            return s.Substring(0, n);
        }

        // ------------------------------------------------------------------
        // XML 辅助
        // ------------------------------------------------------------------

        private static string CellRef(int col, int row)
        {
            return ColumnName(col) + Num(row);
        }

        private static string ColumnName(int col)
        {
            StringBuilder sb = new StringBuilder();
            int value = col;
            while (value > 0)
            {
                int m = (value - 1) % 26;
                sb.Insert(0, (char)('A' + m));
                value = (value - 1) / 26;
            }
            if (sb.Length == 0) sb.Append('A');
            return sb.ToString();
        }

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";

            StringBuilder sb = new StringBuilder(s.Length + 16);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];

                if (char.IsHighSurrogate(c))
                {
                    if (i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
                    {
                        sb.Append(c);
                        sb.Append(s[i + 1]);
                        i++;
                    }
                    // 孤立的代理项直接丢弃，写进 XML 会让文件非法。
                    continue;
                }
                if (char.IsLowSurrogate(c))
                    continue;

                switch (c)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\'': sb.Append("&apos;"); break;
                    case '\r': sb.Append("&#13;"); break;
                    default:
                        // XML 1.0 合法字符；其它控制字符丢弃。
                        if (c == '\t' || c == '\n' ||
                            (c >= 0x20 && c <= 0xD7FF) ||
                            (c >= 0xE000 && c <= 0xFFFD))
                            sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        private static string Num(int value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static string Num(double value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        private static byte[] Utf8(string text)
        {
            return Utf8NoBom.GetBytes(text);
        }

        // ------------------------------------------------------------------
        // ZIP（raw deflate + 自实现 CRC32）
        // ------------------------------------------------------------------

        private sealed class ZipEntryInfo
        {
            public byte[] Name;
            public uint Crc;
            public int CompressedSize;
            public int UncompressedSize;
            public long Offset;
        }

        private sealed class ZipWriter
        {
            private readonly Stream _stream;
            private readonly BinaryWriter _writer;
            private readonly List<ZipEntryInfo> _entries = new List<ZipEntryInfo>();

            public ZipWriter(Stream stream)
            {
                _stream = stream;
                _writer = new BinaryWriter(stream);
            }

            public void Add(string name, byte[] data)
            {
                byte[] nameBytes = Utf8NoBom.GetBytes(name);
                byte[] compressed = Deflate(data);
                uint crc = Crc32(data);
                long offset = _stream.Position;

                // local file header
                _writer.Write(0x04034b50u);
                _writer.Write((ushort)20);       // version needed (2.0 = deflate)
                _writer.Write((ushort)0);        // flags
                _writer.Write((ushort)8);        // method = deflate
                _writer.Write((ushort)0);        // mod time 00:00:00
                _writer.Write((ushort)0x0021);   // mod date 1980-01-01
                _writer.Write(crc);
                _writer.Write((uint)compressed.Length);
                _writer.Write((uint)data.Length);
                _writer.Write((ushort)nameBytes.Length);
                _writer.Write((ushort)0);        // extra length
                _writer.Write(nameBytes);
                _writer.Write(compressed);

                ZipEntryInfo info = new ZipEntryInfo();
                info.Name = nameBytes;
                info.Crc = crc;
                info.CompressedSize = compressed.Length;
                info.UncompressedSize = data.Length;
                info.Offset = offset;
                _entries.Add(info);
            }

            public void Finish()
            {
                long cdOffset = _stream.Position;

                for (int i = 0; i < _entries.Count; i++)
                {
                    ZipEntryInfo e = _entries[i];

                    // central directory header
                    _writer.Write(0x02014b50u);
                    _writer.Write((ushort)20);       // version made by
                    _writer.Write((ushort)20);       // version needed
                    _writer.Write((ushort)0);        // flags
                    _writer.Write((ushort)8);        // method
                    _writer.Write((ushort)0);        // mod time
                    _writer.Write((ushort)0x0021);   // mod date
                    _writer.Write(e.Crc);
                    _writer.Write((uint)e.CompressedSize);
                    _writer.Write((uint)e.UncompressedSize);
                    _writer.Write((ushort)e.Name.Length);
                    _writer.Write((ushort)0);        // extra length
                    _writer.Write((ushort)0);        // comment length
                    _writer.Write((ushort)0);        // disk number start
                    _writer.Write((ushort)0);        // internal attrs
                    _writer.Write((uint)0);          // external attrs
                    _writer.Write((uint)e.Offset);
                    _writer.Write(e.Name);
                }

                long cdSize = _stream.Position - cdOffset;

                // end of central directory
                _writer.Write(0x06054b50u);
                _writer.Write((ushort)0);            // this disk
                _writer.Write((ushort)0);            // disk with central dir
                _writer.Write((ushort)_entries.Count);
                _writer.Write((ushort)_entries.Count);
                _writer.Write((uint)cdSize);
                _writer.Write((uint)cdOffset);
                _writer.Write((ushort)0);            // comment length
                _writer.Flush();
            }

            private static byte[] Deflate(byte[] data)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    // 用 2 参数构造函数，.NET 2.0/4.0 都有；
                    // DeflateStream 默认会关闭 MemoryStream，关闭后 ToArray() 仍可用。
                    using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Compress))
                    {
                        ds.Write(data, 0, data.Length);
                    }
                    return ms.ToArray();
                }
            }

            private static uint Crc32(byte[] data)
            {
                uint crc = 0xFFFFFFFFu;
                for (int i = 0; i < data.Length; i++)
                {
                    crc = CrcTable[(int)((crc ^ data[i]) & 0xFF)] ^ (crc >> 8);
                }
                return crc ^ 0xFFFFFFFFu;
            }
        }

        private static readonly uint[] CrcTable = CreateCrcTable();

        private static uint[] CreateCrcTable()
        {
            uint[] table = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int k = 0; k < 8; k++)
                {
                    if ((c & 1) != 0)
                        c = 0xEDB88320u ^ (c >> 1);
                    else
                        c = c >> 1;
                }
                table[i] = c;
            }
            return table;
        }
    }
}
