using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace DrugFlow
{
    public class MainForm : Form
    {
        private ComboBox cbSource, cbTable, cbDate, cbDrug, cbSpec, cbQty, cbPrice, cbAmount, cbUnit, cbBill, cbDateFmt;
        private DateTimePicker dtFrom, dtTo;
        private TextBox txtSql;
        private DataGridView grid;
        private ToolStripStatusLabel status;
        private Label lblSourceInfo, lblExportInfo;
        private Button btnScan, btnManual, btnReport, btnLoad, btnPreview, btnExport, btnJoin;
        private JoinSpec _join;

        private ScanResult _scan;
        private IProvider _provider;
        private SourceItem _source;
        private List<ColumnInfo> _columns = new List<ColumnInfo>();
        private bool _busy;
        private bool _filling;
        private bool _joinAsked;
        private string _scanHint = "";
        private string _preferredTable;
        private List<Profile> _profiles = new List<Profile>();

        public MainForm()
        {
            BuildUi();
            this.Load += MainForm_Load;
        }

        private void BuildUi()
        {
            this.Text = AppInfo.Name + " v" + AppInfo.Version;
            this.ClientSize = new Size(1000, 742);
            this.MinimumSize = new Size(1020, 780);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft YaHei UI", 9F);
            this.BackColor = Color.FromArgb(248, 249, 251);

            // ---------------- 数据源 ----------------
            GroupBox gb1 = new GroupBox();
            gb1.Text = " 第一步：找到卖药软件的数据 ";
            gb1.SetBounds(12, 10, 976, 88);
            gb1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(gb1);

            gb1.Controls.Add(Lab("数据源", 12, 26, 56));
            cbSource = Cbo(70, 22, 530);
            cbSource.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cbSource.SelectedIndexChanged += cbSource_SelectedIndexChanged;
            gb1.Controls.Add(cbSource);

            btnScan = new Button();
            btnScan.Text = "① 自动扫描";
            btnScan.SetBounds(608, 20, 100, 28);
            btnScan.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnScan.Click += btnScan_Click;
            gb1.Controls.Add(btnScan);

            btnManual = new Button();
            btnManual.Text = "手动连接";
            btnManual.SetBounds(714, 20, 92, 28);
            btnManual.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnManual.Click += btnManual_Click;
            gb1.Controls.Add(btnManual);

            btnReport = new Button();
            btnReport.Text = "生成诊断报告";
            btnReport.SetBounds(812, 20, 150, 28);
            btnReport.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnReport.Click += btnReport_Click;
            gb1.Controls.Add(btnReport);

            lblSourceInfo = Lab("尚未连接。点“自动扫描”，程序会自己找本机的卖药软件数据库。", 70, 54, 890);
            lblSourceInfo.ForeColor = Color.FromArgb(90, 98, 110);
            gb1.Controls.Add(lblSourceInfo);

            // ---------------- 表与字段 ----------------
            GroupBox gb2 = new GroupBox();
            gb2.Text = " 第二步：选择数据表和字段（一般不用改，程序会自动猜） ";
            gb2.SetBounds(12, 104, 976, 170);
            gb2.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(gb2);

            gb2.Controls.Add(Lab("数据表", 12, 26, 56));
            cbTable = Cbo(70, 22, 330);
            cbTable.SelectedIndexChanged += cbTable_SelectedIndexChanged;
            gb2.Controls.Add(cbTable);

            btnLoad = new Button();
            btnLoad.Text = "刷新字段";
            btnLoad.SetBounds(410, 21, 88, 26);
            btnLoad.Click += btnLoad_Click;
            gb2.Controls.Add(btnLoad);

            btnJoin = new Button();
            btnJoin.Text = "关联主表…";
            btnJoin.SetBounds(504, 21, 96, 26);
            btnJoin.Click += btnJoin_Click;
            gb2.Controls.Add(btnJoin);
            ToolTip tip2 = new ToolTip();
            tip2.SetToolTip(btnJoin, "日期不在明细表里时用：把销售明细表和销售单头（主表）关联起来");

            gb2.Controls.Add(Lab("日期字段", 12, 60, 64));
            cbDate = Cbo(80, 56, 190); gb2.Controls.Add(cbDate);

            gb2.Controls.Add(Lab("药品名称", 286, 60, 64));
            cbDrug = Cbo(354, 56, 190); gb2.Controls.Add(cbDrug);

            gb2.Controls.Add(Lab("规格", 560, 60, 36));
            cbSpec = Cbo(598, 56, 190); gb2.Controls.Add(cbSpec);

            gb2.Controls.Add(Lab("数量", 12, 96, 64));
            cbQty = Cbo(80, 92, 190); gb2.Controls.Add(cbQty);

            gb2.Controls.Add(Lab("单价", 286, 96, 64));
            cbPrice = Cbo(354, 92, 190); gb2.Controls.Add(cbPrice);

            gb2.Controls.Add(Lab("金额", 560, 96, 36));
            cbAmount = Cbo(598, 92, 190); gb2.Controls.Add(cbAmount);

            gb2.Controls.Add(Lab("单位", 12, 132, 64));
            cbUnit = Cbo(80, 128, 190); gb2.Controls.Add(cbUnit);

            gb2.Controls.Add(Lab("单据号", 286, 132, 64));
            cbBill = Cbo(354, 128, 190); gb2.Controls.Add(cbBill);

            Label hint = Lab("数量/单价/金额对不上时，可以在这里手工指定对应字段", 560, 132, 400);
            hint.ForeColor = Color.Gray;
            gb2.Controls.Add(hint);

            // ---------------- 日期 ----------------
            GroupBox gb3 = new GroupBox();
            gb3.Text = " 第三步：选择日期范围 ";
            gb3.SetBounds(12, 280, 544, 94);
            this.Controls.Add(gb3);

            gb3.Controls.Add(Lab("从", 12, 32, 24));
            dtFrom = new DateTimePicker();
            dtFrom.Format = DateTimePickerFormat.Custom;
            dtFrom.CustomFormat = "yyyy-MM-dd";
            dtFrom.SetBounds(38, 28, 140, 24);
            dtFrom.ValueChanged += dt_ValueChanged;
            gb3.Controls.Add(dtFrom);

            gb3.Controls.Add(Lab("到", 190, 32, 24));
            dtTo = new DateTimePicker();
            dtTo.Format = DateTimePickerFormat.Custom;
            dtTo.CustomFormat = "yyyy-MM-dd";
            dtTo.SetBounds(216, 28, 140, 24);
            dtTo.ValueChanged += dt_ValueChanged;
            gb3.Controls.Add(dtTo);

            string[] quick = new string[] { "今天", "昨天", "本月", "上月", "今年" };
            for (int i = 0; i < quick.Length; i++)
            {
                Button b = new Button();
                b.Text = quick[i];
                b.SetBounds(12 + i * 62, 58, 58, 26);
                b.Tag = quick[i];
                b.Click += quick_Click;
                gb3.Controls.Add(b);
            }

            gb3.Controls.Add(Lab("日期格式", 366, 32, 62));
            cbDateFmt = new ComboBox();
            cbDateFmt.SetBounds(430, 28, 100, 24);
            cbDateFmt.DropDownStyle = ComboBoxStyle.DropDownList;
            cbDateFmt.DropDownWidth = 180;
            cbDateFmt.Items.Add("自动");
            cbDateFmt.Items.Add("yyyy-MM-dd");
            cbDateFmt.Items.Add("yyyyMMdd");
            cbDateFmt.Items.Add("yyyy/MM/dd");
            cbDateFmt.SelectedIndex = 0;
            cbDateFmt.SelectedIndexChanged += delegate { RebuildSql(); };
            gb3.Controls.Add(cbDateFmt);

            ToolTip tip = new ToolTip();
            tip.SetToolTip(cbDateFmt, "默认“自动”。如果预览是空的、或者日期范围对不上，\n说明软件把日期存成了文本，换一种格式再试。");
            tip.SetToolTip(cbDate, "销售日期所在的字段");

            // ---------------- 导出 ----------------
            GroupBox gb4 = new GroupBox();
            gb4.Text = " 第四步：预览并导出 Excel ";
            gb4.SetBounds(568, 280, 420, 94);
            gb4.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(gb4);

            btnPreview = new Button();
            btnPreview.Text = "预览数据";
            btnPreview.SetBounds(12, 20, 110, 28);
            btnPreview.Click += btnPreview_Click;
            gb4.Controls.Add(btnPreview);

            btnExport = new Button();
            btnExport.Text = "导出 Excel";
            btnExport.SetBounds(130, 16, 160, 36);
            btnExport.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            btnExport.BackColor = Color.FromArgb(31, 119, 214);
            btnExport.ForeColor = Color.White;
            btnExport.FlatStyle = FlatStyle.Flat;
            btnExport.FlatAppearance.BorderSize = 0;
            btnExport.Click += btnExport_Click;
            gb4.Controls.Add(btnExport);

            lblExportInfo = Lab("导出后生成两个工作表：销售明细 + 按药品汇总。", 12, 58, 396);
            lblExportInfo.ForeColor = Color.FromArgb(90, 98, 110);
            gb4.Controls.Add(lblExportInfo);

            // ---------------- SQL ----------------
            Label labSql = Lab("查询 SQL（高级：可以手工修改，然后点“预览数据”或“导出 Excel”）", 12, 382, 700);
            labSql.ForeColor = Color.FromArgb(90, 98, 110);
            this.Controls.Add(labSql);

            txtSql = new TextBox();
            txtSql.SetBounds(12, 404, 976, 58);
            txtSql.Multiline = true;
            txtSql.ScrollBars = ScrollBars.Vertical;
            txtSql.Font = new Font("Consolas", 9F);
            txtSql.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(txtSql);

            // ---------------- 表格 ----------------
            grid = new DataGridView();
            grid.SetBounds(12, 470, 976, 236);
            grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.RowHeadersVisible = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.FixedSingle;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(246, 249, 253);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            this.Controls.Add(grid);

            // ---------------- 状态栏 ----------------
            StatusStrip ss = new StatusStrip();
            status = new ToolStripStatusLabel("就绪");
            status.Spring = true;
            status.TextAlign = ContentAlignment.MiddleLeft;
            ss.Items.Add(status);

            ToolStripButton btnImport = new ToolStripButton("导入方案文件…");
            btnImport.ToolTipText = "导入别人做好的适配方案（.ini），不用重新配置";
            btnImport.Click += btnImport_Click;
            ss.Items.Add(btnImport);

            this.Controls.Add(ss);
        }

        private static Label Lab(string text, int x, int y, int w)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(x, y, w, 20);
            l.TextAlign = ContentAlignment.MiddleLeft;
            return l;
        }

        private static ComboBox Cbo(int x, int y, int w)
        {
            ComboBox c = new ComboBox();
            c.SetBounds(x, y, w, 24);
            c.DropDownStyle = ComboBoxStyle.DropDownList;
            c.DropDownWidth = Math.Max(w, 320);
            return c;
        }

        // ================= 启动 =================
        private void MainForm_Load(object sender, EventArgs e)
        {
            DateTime today = DateTime.Today;
            dtFrom.Value = new DateTime(today.Year, today.Month, 1);
            dtTo.Value = today;

            _profiles = Profile.LoadAll();
            RebuildSourceList();
            ConnectSource();
            SetStatus("准备就绪，正在自动扫描本机卖药软件…");
            StartScan();
        }

        // ================= 扫描 =================
        private void btnScan_Click(object sender, EventArgs e) { StartScan(); }

        private void StartScan()
        {
            if (_busy) return;
            SetBusy(true);
            SetStatus("正在扫描…");
            Thread t = new Thread(delegate()
            {
                ScanResult r = null;
                string err = null;
                try { r = Scanner.Scan(delegate(string msg) { SetStatus(msg); }); }
                catch (Exception ex) { err = ex.Message; }
                try { this.BeginInvoke(new Action(delegate() { OnScanDone(r, err); })); }
                catch { }
            });
            t.IsBackground = true;
            t.Start();
        }

        private void OnScanDone(ScanResult r, string err)
        {
            SetBusy(false);
            if (r == null)
            {
                SetStatus("扫描失败：" + err);
                MessageBox.Show(this, "扫描时出错：\n" + err, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _scan = r;
            _profiles = Profile.LoadAll();
            RebuildSourceList();
            ConnectSource();

            int soft = 0;
            foreach (Finding f in r.Findings) if (f.Kind == "运行中的软件" || f.Kind == "已安装程序") soft++;
            lblSourceInfo.Text = "扫描完成：发现 " + soft + " 个疑似医药软件线索，" + r.Sources.Count + " 个可用数据源。" +
                                 "请在上面的下拉框里选择卖药软件的数据源；选不到就点“生成诊断报告”发给我。";
            _scanHint = "（已扫描到 " + r.Sources.Count + " 个数据源）";
            SetStatus("扫描完成，共 " + r.Findings.Count + " 条线索。");
        }

        /// <summary>数据源下拉框 = 示例 + 方案 + 扫描结果</summary>
        private void RebuildSourceList()
        {
            _filling = true;
            cbSource.Items.Clear();

            SourceItem demo = null;
            List<SourceItem> others = new List<SourceItem>();
            if (_scan != null)
            {
                foreach (SourceItem s in _scan.Sources)
                {
                    if (s.Kind == "demo") demo = s;
                    else others.Add(s);
                }
            }
            if (demo == null)
            {
                demo = new SourceItem();
                demo.Kind = "demo";
                demo.Title = "示例数据（先看看导出效果）";
            }
            cbSource.Items.Add(demo);

            for (int i = 0; i < _profiles.Count; i++)
            {
                SourceItem ps = new SourceItem();
                ps.Kind = "profile";
                ps.Title = "方案：" + _profiles[i].Name;
                ps.Connection = _profiles[i].Connection;
                ps.Profile = _profiles[i];
                cbSource.Items.Add(ps);
            }

            for (int i = 0; i < others.Count; i++) cbSource.Items.Add(others[i]);

            if (cbSource.Items.Count > 0) cbSource.SelectedIndex = 0;
            _filling = false;
        }

        // ================= 连接 =================
        private void cbSource_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_filling) return;
            ConnectSource();
        }

        private void ConnectSource()
        {
            if (_busy) return;
            SourceItem s = cbSource.SelectedItem as SourceItem;
            if (s == null) return;

            CloseProvider();
            _source = s;
            _columns.Clear();
            _joinAsked = false;

            if (s.Kind == "profile" && s.Profile != null)
            {
                ApplyProfile(s);
                return;
            }
            cbTable.Items.Clear();
            ResetMappingCombos();
            txtSql.Text = "";
            grid.DataSource = null;

            if (s.Kind == "demo")
            {
                _provider = new DemoProvider();
                LoadTables();
                lblSourceInfo.Text = "当前数据源：示例数据（演示用，不是真实数据）　选真实数据请用上面的下拉框。";
                return;
            }

            Cursor = Cursors.WaitCursor;
            SetStatus("正在连接：" + s.Title + " …");
            try
            {
                _provider = DbFactory.Create(s);
                LoadTables();
                lblSourceInfo.Text = "当前数据源：" + s.Title + "  " + _scanHint;
                SetStatus("已连接：" + s.Title);
            }
            catch (Exception ex)
            {
                CloseProvider();
                SetStatus("连接失败：" + ex.Message);
                MessageBox.Show(this, "连接失败：\n" + ex.Message + "\n\n可以试试：\n1) 点“手动连接”换一种方式；\n2) 先启动卖药软件再扫描；\n3) 点“生成诊断报告”发给我。",
                    "连接失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Cursor = Cursors.Default; }
        }

        private void LoadTables()
        {
            if (_provider == null) return;
            Cursor = Cursors.WaitCursor;
            try
            {
                List<string> tables = _provider.ListTables();
                _filling = true;
                cbTable.Items.Clear();
                foreach (string t in tables) cbTable.Items.Add(t);
                _filling = false;
                if (tables.Count == 0)
                {
                    SetStatus("这个数据源里没有表。");
                    return;
                }
                int best = 0, bestScore = int.MinValue;
                if (!string.IsNullOrEmpty(_preferredTable))
                {
                    for (int i = 0; i < tables.Count; i++)
                    {
                        if (string.Equals(tables[i], _preferredTable, StringComparison.OrdinalIgnoreCase))
                        {
                            best = i;
                            bestScore = int.MaxValue;
                            break;
                        }
                    }
                }
                if (bestScore != int.MaxValue)
                {
                    for (int i = 0; i < tables.Count; i++)
                    {
                        int sc = Scanner.TableScore(tables[i]);
                        if (sc > bestScore) { bestScore = sc; best = i; }
                    }
                }
                _preferredTable = null;
                cbTable.SelectedIndex = best;
                LoadColumns();

                if (string.IsNullOrEmpty(SelectedName(cbDate)) && _join == null && cbTable.Items.Count > 1 && !_joinAsked)
                {
                    _joinAsked = true;
                    Cursor = Cursors.Default;
                    DialogResult dr = MessageBox.Show(this,
                        "这张表里没有日期字段。\n日期很可能记在另一张表（销售单头 / 主表）里。\n\n现在设置“关联主表”吗？",
                        "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes) btnJoin_Click(null, null);
                }
            }
            catch (Exception ex)
            {
                SetStatus("读取表失败：" + ex.Message);
            }
            finally { Cursor = Cursors.Default; }
        }

        private void cbTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_filling) return;
            LoadColumns();
        }

        private void btnLoad_Click(object sender, EventArgs e) { LoadColumns(); }

        private void LoadColumns()
        {
            if (_provider == null || cbTable.SelectedItem == null) return;
            _join = null;
            Cursor = Cursors.WaitCursor;
            try
            {
                _columns = _provider.ListColumns(Convert.ToString(cbTable.SelectedItem));
                FillMappingCombos();
                AutoMap();
                RebuildSql();
                if (string.IsNullOrEmpty(SelectedName(cbDate)) && _join == null)
                {
                    SetStatus("提示：这张表里没有日期字段。如果日期在另一张表（销售单头）里，请点“关联主表…”。");
                }
                else
                {
                    SetStatus("已读取 " + _columns.Count + " 个字段。");
                }
            }
            catch (Exception ex)
            {
                SetStatus("读取字段失败：" + ex.Message);
            }
            finally { Cursor = Cursors.Default; }
        }

        private void ResetMappingCombos()
        {
            ComboBox[] all = new ComboBox[] { cbDate, cbDrug, cbSpec, cbQty, cbPrice, cbAmount, cbUnit, cbBill };
            foreach (ComboBox c in all)
            {
                c.Items.Clear();
                c.Items.Add("（不使用）");
                c.SelectedIndex = 0;
            }
        }

        private void FillMappingCombos()
        {
            ComboBox[] all = new ComboBox[] { cbDate, cbDrug, cbSpec, cbQty, cbPrice, cbAmount, cbUnit, cbBill };
            _filling = true;
            foreach (ComboBox c in all)
            {
                c.Items.Clear();
                c.Items.Add("（不使用）");
                foreach (ColumnInfo ci in _columns) c.Items.Add(ci.Name);
                c.SelectedIndex = 0;
            }
            _filling = false;
        }

        private void AutoMap()
        {
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _filling = true;
            cbDate.SelectedIndex = Pick(cbDate, used, new string[] { "日期", "销售日期", "业务日期", "单据日期", "开单日期", "收费日期", "销售时间", "时间" },
                new string[] { "日期" }, new string[] { "日期", "时间", "date", "time" }, "datetime");
            cbDrug.SelectedIndex = Pick(cbDrug, used, new string[] { "药品名称", "药品", "品名", "商品名称", "名称", "药名" },
                new string[] { "药品", "品名" }, new string[] { "药品", "品名", "药名", "drug", "item", "product", "goods" }, null);
            cbSpec.SelectedIndex = Pick(cbSpec, used, new string[] { "规格" }, new string[] { "规格" }, new string[] { "规格", "spec", "包装" }, null);
            cbQty.SelectedIndex = Pick(cbQty, used, new string[] { "数量", "销售数量" }, new string[] { "数量" }, new string[] { "数量", "qty", "quantity" }, "numeric");
            cbPrice.SelectedIndex = Pick(cbPrice, used, new string[] { "单价", "零售价", "售价" }, new string[] { "单价" }, new string[] { "单价", "price", "售价" }, "numeric");
            cbAmount.SelectedIndex = Pick(cbAmount, used, new string[] { "金额", "销售金额", "合计", "小计", "应收金额", "实收金额" },
                new string[] { "金额" }, new string[] { "金额", "amount", "total", "小计" }, "numeric");
            cbUnit.SelectedIndex = Pick(cbUnit, used, new string[] { "单位" }, new string[] { "单位" }, new string[] { "单位", "unit" }, null);
            cbBill.SelectedIndex = Pick(cbBill, used, new string[] { "单据号", "单号", "销售单号", "票号", "流水号" },
                new string[] { "单号", "单据" }, new string[] { "单号", "单据", "bill", "票号", "no" }, null);
            _filling = false;
        }

        private int Pick(ComboBox box, HashSet<string> used, string[] exact, string[] starts, string[] contains, string typeHint)
        {
            int bestIdx = 0, bestScore = 0;
            for (int i = 0; i < _columns.Count; i++)
            {
                ColumnInfo ci = _columns[i];
                if (ci.Name == null || used.Contains(ci.Name)) continue;
                string n = ci.Name.ToLowerInvariant();
                int sc = 0;
                for (int k = 0; k < exact.Length; k++)
                {
                    if (string.Equals(ci.Name, exact[k], StringComparison.OrdinalIgnoreCase)) { sc = Math.Max(sc, 100); }
                }
                for (int k = 0; k < starts.Length; k++)
                {
                    if (n.StartsWith(starts[k].ToLowerInvariant())) sc = Math.Max(sc, 80);
                }
                for (int k = 0; k < contains.Length; k++)
                {
                    if (n.IndexOf(contains[k].ToLowerInvariant(), StringComparison.Ordinal) >= 0) sc = Math.Max(sc, 60);
                }
                if (sc == 0) continue;
                if (!string.IsNullOrEmpty(typeHint))
                {
                    string t = (ci.Type == null ? "" : ci.Type.ToLowerInvariant());
                    if (typeHint == "datetime" && (t.IndexOf("date") >= 0 || t.IndexOf("time") >= 0 || t.IndexOf("7") == 0)) sc += 15;
                    if (typeHint == "numeric" && (t.IndexOf("int") >= 0 || t.IndexOf("decimal") >= 0 || t.IndexOf("numeric") >= 0 || t.IndexOf("money") >= 0 || t.IndexOf("float") >= 0 || t.IndexOf("double") >= 0)) sc += 15;
                }
                if (sc > bestScore) { bestScore = sc; bestIdx = i + 1; }
            }
            if (bestScore == 0) return 0;
            used.Add(_columns[bestIdx - 1].Name);
            return bestIdx;
        }

        // ================= 方案 =================
        private void ApplyProfile(SourceItem s)
        {
            Profile p = s.Profile;
            Cursor = Cursors.WaitCursor;
            _filling = true;
            cbTable.Items.Clear();
            ResetMappingCombos();
            txtSql.Text = "";
            grid.DataSource = null;
            int style = 0;
            if (p.DateStyle == "ymd") style = 1;
            else if (p.DateStyle == "ymd8") style = 2;
            else if (p.DateStyle == "slash") style = 3;
            cbDateFmt.SelectedIndex = style;
            _filling = false;

            try
            {
                SourceItem real = new SourceItem();
                real.Kind = p.Kind;
                real.Connection = p.Connection;
                real.Title = p.Name;
                _provider = DbFactory.Create(real);
                _preferredTable = p.Table;
                _joinAsked = true;
                LoadTables();
                ApplyMapping(p);
                _join = p.ToJoin();
                RebuildSql();
                lblSourceInfo.Text = "当前方案：" + p.Name + "  " + _scanHint;
                SetStatus("已套用方案：" + p.Name + "，选好日期就能导出。");
            }
            catch (Exception ex)
            {
                CloseProvider();
                SetStatus("方案连接失败：" + ex.Message);
                MessageBox.Show(this, "这个方案连接不上：\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Cursor = Cursors.Default; }
        }

        private void ApplyMapping(Profile p)
        {
            SetCombo(cbDate, p.Date);
            SetCombo(cbDrug, p.Drug);
            SetCombo(cbSpec, p.Spec);
            SetCombo(cbQty, p.Qty);
            SetCombo(cbPrice, p.Price);
            SetCombo(cbAmount, p.Amount);
            SetCombo(cbUnit, p.Unit);
            SetCombo(cbBill, p.Bill);
        }

        private static void SetCombo(ComboBox c, string name)
        {
            if (c == null || string.IsNullOrEmpty(name)) return;
            int i = c.Items.IndexOf(name);
            if (i >= 0) c.SelectedIndex = i;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "选择方案文件";
            ofd.Filter = "方案文件 (*.ini)|*.ini|所有文件 (*.*)|*.*";
            if (ofd.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                Profile p = Profile.Load(ofd.FileName);
                if (string.IsNullOrEmpty(p.Kind) || string.IsNullOrEmpty(p.Connection))
                {
                    MessageBox.Show(this, "这个文件不是有效的方案文件（缺少 kind 或 connection）。", "提示");
                    return;
                }
                if (!System.IO.Directory.Exists(Profile.UserDir)) System.IO.Directory.CreateDirectory(Profile.UserDir);
                string safe = p.Name;
                char[] bad = System.IO.Path.GetInvalidFileNameChars();
                for (int i = 0; i < bad.Length; i++) safe = safe.Replace(bad[i], '_');
                p.Save(System.IO.Path.Combine(Profile.UserDir, safe + ".ini"));

                _profiles = Profile.LoadAll();
                RebuildSourceList();
                for (int i = 0; i < cbSource.Items.Count; i++)
                {
                    SourceItem si = cbSource.Items[i] as SourceItem;
                    if (si != null && si.Profile != null && si.Profile.Name == p.Name)
                    {
                        _filling = true;
                        cbSource.SelectedIndex = i;
                        _filling = false;
                        ConnectSource();
                        break;
                    }
                }
                MessageBox.Show(this, "方案已导入并选中：" + p.Name + "\n以后打开程序，在“数据源”里选它就行。", "导入成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "导入失败：" + ex.Message, "提示");
            }
        }

        private void SaveLastProfile()
        {
            try
            {
                if (_source == null) return;
                Profile p = new Profile();
                if (_source.Profile != null)
                {
                    p.Kind = _source.Profile.Kind;
                    p.Connection = _source.Profile.Connection;
                }
                else
                {
                    p.Kind = _source.Kind;
                    p.Connection = _source.Connection;
                }
                if (p.Kind == "demo" || string.IsNullOrEmpty(p.Kind) || string.IsNullOrEmpty(p.Connection)) return;

                p.Table = Convert.ToString(cbTable.Text);
                FieldMap m = CurrentMap();
                p.Date = (_join != null && m.Date == JoinSpec.DateAlias) ? _join.DateColumn : m.Date;
                p.Drug = m.Drug; p.Spec = m.Spec; p.Qty = m.Qty;
                p.Price = m.Price; p.Amount = m.Amount; p.Unit = m.Unit; p.Bill = m.Bill;
                p.DateStyle = DateStyle();
                p.FromJoin(_join);
                Profile.SaveLast(p);
                _profiles = Profile.LoadAll();
            }
            catch { }
        }

        // ================= 日期 / SQL =================
        private void dt_ValueChanged(object sender, EventArgs e)
        {
            if (_filling) return;
            RebuildSql();
        }

        private void quick_Click(object sender, EventArgs e)
        {
            Button b = sender as Button;
            string tag = b == null ? "" : Convert.ToString(b.Tag);
            DateTime today = DateTime.Today;
            DateTime from = dtFrom.Value.Date, to = dtTo.Value.Date;
            if (tag == "今天") { from = today; to = today; }
            else if (tag == "昨天") { from = today.AddDays(-1); to = today.AddDays(-1); }
            else if (tag == "本月") { from = new DateTime(today.Year, today.Month, 1); to = today; }
            else if (tag == "上月")
            {
                DateTime first = new DateTime(today.Year, today.Month, 1);
                from = first.AddMonths(-1);
                to = first.AddDays(-1);
            }
            else if (tag == "今年") { from = new DateTime(today.Year, 1, 1); to = today; }
            dtFrom.Value = from;
            dtTo.Value = to;
            RebuildSql();
        }

        private string DateStyle()
        {
            if (cbDateFmt == null) return "auto";
            switch (cbDateFmt.SelectedIndex)
            {
                case 1: return "ymd";
                case 2: return "ymd8";
                case 3: return "slash";
                default: return "auto";
            }
        }

        private FieldMap CurrentMap()
        {
            FieldMap m = new FieldMap();
            m.Date = SelectedName(cbDate);
            m.Drug = SelectedName(cbDrug);
            m.Spec = SelectedName(cbSpec);
            m.Qty = SelectedName(cbQty);
            m.Price = SelectedName(cbPrice);
            m.Amount = SelectedName(cbAmount);
            m.Unit = SelectedName(cbUnit);
            m.Bill = SelectedName(cbBill);
            return m;
        }

        private static string SelectedName(ComboBox c)
        {
            if (c.SelectedIndex <= 0) return "";
            return Convert.ToString(c.SelectedItem);
        }

        private void RebuildSql()
        {
            if (_provider == null || cbTable.SelectedItem == null) { txtSql.Text = ""; return; }
            if (dtTo.Value.Date < dtFrom.Value.Date) dtTo.Value = dtFrom.Value;
            try
            {
                string table = Convert.ToString(cbTable.SelectedItem);
                if (_join != null && _join.IsValid)
                {
                    txtSql.Text = Exporter.BuildSqlJoin(_provider, table, _columns, _join, dtFrom.Value.Date, dtTo.Value.Date, DateStyle());
                    if (!cbDate.Items.Contains(JoinSpec.DateAlias)) cbDate.Items.Add(JoinSpec.DateAlias);
                    cbDate.SelectedItem = JoinSpec.DateAlias;
                }
                else
                {
                    txtSql.Text = Exporter.BuildSql(_provider, table, CurrentMap(), dtFrom.Value.Date, dtTo.Value.Date, DateStyle());
                }
            }
            catch (Exception ex) { txtSql.Text = "-- 生成 SQL 失败：" + ex.Message; }
        }

        private void btnJoin_Click(object sender, EventArgs e)
        {
            if (_provider == null || cbTable.SelectedItem == null)
            {
                MessageBox.Show(this, "请先连接数据源、选好明细表。", "提示");
                return;
            }
            List<string> tables = new List<string>();
            foreach (object o in cbTable.Items) tables.Add(Convert.ToString(o));
            using (JoinForm f = new JoinForm(_provider, Convert.ToString(cbTable.SelectedItem), tables, _columns))
            {
                if (f.ShowDialog(this) != DialogResult.OK || f.Result == null) return;
                _join = f.Result;
                RebuildSql();
                SetStatus("已关联主表 " + _join.MasterTable + "，按主表的 " + _join.DateColumn + " 筛选日期。");
                lblExportInfo.Text = "已关联主表，日期取自 " + _join.MasterTable + "." + _join.DateColumn;
            }
        }

        // ================= 预览 =================
        private void btnPreview_Click(object sender, EventArgs e)
        {
            if (!EnsureReady()) return;
            Cursor = Cursors.WaitCursor;
            try
            {
                DataTable dt = _provider.Query(txtSql.Text);
                DataTable shown = CopyTop(dt, 500);
                grid.DataSource = shown;
                SetStatus("预览 " + shown.Rows.Count + " 行" + (dt.Rows.Count > shown.Rows.Count ? "（共 " + dt.Rows.Count + " 行，仅显示前 500 行）" : "") + "。");
                lblExportInfo.Text = "预览正常？点“导出 Excel”生成文件。";
            }
            catch (Exception ex)
            {
                SetStatus("查询失败：" + ex.Message);
                MessageBox.Show(this, "查询失败：\n" + ex.Message + "\n\n可以检查上面的 SQL，或把这段错误连同诊断报告一起发我。",
                    "查询失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Cursor = Cursors.Default; }
        }

        private static DataTable CopyTop(DataTable dt, int n)
        {
            if (dt.Rows.Count <= n) return dt;
            DataTable copy = dt.Clone();
            for (int i = 0; i < n; i++) copy.ImportRow(dt.Rows[i]);
            return copy;
        }

        // ================= 导出 =================
        private void btnExport_Click(object sender, EventArgs e)
        {
            if (!EnsureReady()) return;

            FieldMap map = CurrentMap();
            if (string.IsNullOrEmpty(map.Drug))
            {
                MessageBox.Show(this, "还没有指定“药品名称”字段。\n请在中间的下拉框里选一下。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (string.IsNullOrEmpty(map.Date))
            {
                DialogResult warn = MessageBox.Show(this,
                    "没有选择“日期字段”，程序会把表里所有记录都导出，而不只是所选日期范围。\n\n" +
                    "如果日期在另一张表（销售单头）里，请先点“关联主表…”。\n\n" +
                    "仍要继续导出吗？",
                    "注意", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (warn != DialogResult.Yes) return;
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "保存 Excel 文件";
            sfd.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
            sfd.FileName = Exporter.SuggestFileName(dtFrom.Value.Date, dtTo.Value.Date);
            try { sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); } catch { }
            if (sfd.ShowDialog(this) != DialogResult.OK) return;

            Cursor = Cursors.WaitCursor;
            SetStatus("正在导出…");
            try
            {
                DateTime from = dtFrom.Value.Date;
                DateTime to = dtTo.Value.Date;
                DataTable dt = _provider.Query(txtSql.Text);
                if (dt.Rows.Count == 0)
                {
                    Cursor = Cursors.Default;
                    SetStatus("没有数据。");
                    MessageBox.Show(this, "这个日期范围内没有查到记录。\n请确认日期，或点“预览数据”看看 SQL 查出来是什么。", "没有数据",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                ExportResult res = Exporter.Export(dt, map, from, to, sfd.FileName);
                SaveLastProfile();
                grid.DataSource = CopyTop(dt, 500);
                SetStatus("导出成功：" + res.Path);
                lblExportInfo.Text = "已导出 " + res.RowCount + " 行，" + res.DrugCount + " 种药品，合计 " + res.TotalAmount.ToString("0.00") + " 元。";
                DialogResult dr = MessageBox.Show(this,
                    "导出成功！\n\n" +
                    "明细行数：" + res.RowCount + "\n" +
                    "药品品种：" + res.DrugCount + "\n" +
                    "合计金额：" + res.TotalAmount.ToString("0.00") + " 元\n\n" +
                    "文件位置：\n" + res.Path + "\n\n要现在打开这个文件吗？",
                    "导出成功", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes)
                {
                    try { Process.Start(res.Path); }
                    catch (Exception ex) { MessageBox.Show(this, "打不开文件：" + ex.Message + "\n（文件已经生成，可以到文件夹里双击打开）", "提示"); }
                }
            }
            catch (Exception ex)
            {
                SetStatus("导出失败：" + ex.Message);
                MessageBox.Show(this, "导出失败：\n" + ex.Message, "导出失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { Cursor = Cursors.Default; }
        }

        private bool EnsureReady()
        {
            if (_provider == null)
            {
                MessageBox.Show(this, "还没有连接数据源。\n请先点“自动扫描”，或在“数据源”里选一个。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (string.IsNullOrEmpty(txtSql.Text.Trim()))
            {
                MessageBox.Show(this, "还没有生成查询语句。\n请先选择数据表和字段。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        // ================= 诊断报告 =================
        private void btnReport_Click(object sender, EventArgs e)
        {
            // 诊断报告要尽量全：这里总是做一次深度扫描（会读进程列表）
            Cursor = Cursors.WaitCursor;
            SetStatus("正在收集完整信息（会读取进程列表，稍慢）…");
            try { _scan = Scanner.Scan(delegate(string msg) { SetStatus(msg); }); }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                MessageBox.Show(this, "生成报告失败：" + ex.Message, "提示");
                return;
            }
            finally { Cursor = Cursors.Default; }
            try
            {
                string path = Report.Save(_scan);
                try { Clipboard.SetText(path); } catch { }
                SetStatus("诊断报告已保存：" + path);
                DialogResult dr = MessageBox.Show(this,
                    "诊断报告已生成（路径已复制到剪贴板）：\n\n" + path + "\n\n" +
                    "请把这个 txt 文件发给我，我就能判断你家用的是哪种数据库、该怎么取数。\n" +
                    "报告里没有具体药品和顾客信息，不用担心泄露。\n\n要现在打开它所在的文件夹吗？",
                    "诊断报告", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (dr == DialogResult.Yes)
                {
                    try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "保存报告失败：" + ex.Message, "提示");
            }
        }

        // ================= 手动连接 =================
        private void btnManual_Click(object sender, EventArgs e)
        {
            using (ManualForm f = new ManualForm())
            {
                if (f.ShowDialog(this) != DialogResult.OK) return;
                SourceItem s = f.Result;
                if (s == null) return;
                _filling = true;
                cbSource.Items.Add(s);
                cbSource.SelectedIndex = cbSource.Items.Count - 1;
                _filling = false;
                ConnectSource();
            }
        }

        // ================= 通用 =================
        private void SetBusy(bool busy)
        {
            _busy = busy;
            btnScan.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        private void SetStatus(string msg)
        {
            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new Action(delegate() { status.Text = msg; })); } catch { }
                return;
            }
            status.Text = msg;
        }

        private void CloseProvider()
        {
            if (_provider != null)
            {
                try { _provider.Dispose(); } catch { }
                _provider = null;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            CloseProvider();
            base.OnFormClosing(e);
        }
    }
}
