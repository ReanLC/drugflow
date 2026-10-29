using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace DrugFlow
{
    /// <summary>当日期不在明细表里时，把销售明细表和销售主表关联起来</summary>
    public class JoinForm : Form
    {
        private IProvider _provider;
        private string _detailTable;
        private List<string> _tables;
        private List<ColumnInfo> _detailColumns;

        private ComboBox cbMaster, cbDetailKey, cbMasterKey, cbMasterDate;
        private Label lblDetail;
        public JoinSpec Result;

        public JoinForm(IProvider provider, string detailTable, List<string> tables, List<ColumnInfo> detailColumns)
        {
            _provider = provider;
            _detailTable = detailTable;
            _tables = tables;
            _detailColumns = detailColumns;

            this.Text = "关联主表";
            this.ClientSize = new Size(640, 330);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei UI", 9F);

            this.Controls.Add(Lab("明细表", 16, 18, 70));
            lblDetail = Lab(detailTable, 90, 18, 530);
            lblDetail.ForeColor = Color.FromArgb(31, 119, 214);
            this.Controls.Add(lblDetail);

            this.Controls.Add(Lab("主表", 16, 54, 70));
            cbMaster = new ComboBox();
            cbMaster.SetBounds(90, 50, 530, 24);
            cbMaster.DropDownStyle = ComboBoxStyle.DropDownList;
            cbMaster.SelectedIndexChanged += cbMaster_SelectedIndexChanged;
            this.Controls.Add(cbMaster);

            this.Controls.Add(Lab("明细表字段", 16, 90, 80));
            cbDetailKey = new ComboBox();
            cbDetailKey.SetBounds(100, 86, 190, 24);
            cbDetailKey.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Controls.Add(cbDetailKey);

            this.Controls.Add(Lab("主表字段", 306, 90, 70));
            cbMasterKey = new ComboBox();
            cbMasterKey.SetBounds(380, 86, 240, 24);
            cbMasterKey.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Controls.Add(cbMasterKey);

            this.Controls.Add(Lab("主表日期字段", 16, 126, 90));
            cbMasterDate = new ComboBox();
            cbMasterDate.SetBounds(110, 122, 220, 24);
            cbMasterDate.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Controls.Add(cbMasterDate);

            Label hint = new Label();
            hint.SetBounds(16, 160, 610, 90);
            hint.ForeColor = Color.FromArgb(90, 98, 110);
            hint.Text = "什么时候用这个：\r\n" +
                "销售明细表里只有药品、数量、金额，日期和单号在另一张“销售单头/主表”里。\r\n" +
                "两张表靠“单据号”之类的字段对应。选好之后，按日期筛选就走主表的日期。";
            this.Controls.Add(hint);

            Button ok = new Button();
            ok.Text = "确定关联";
            ok.SetBounds(400, 270, 100, 32);
            ok.Click += btnOk_Click;
            this.Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.SetBounds(510, 270, 100, 32);
            cancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(cancel);

            LoadTables();
        }

        private static Label Lab(string text, int x, int y, int w)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(x, y, w, 20);
            l.TextAlign = ContentAlignment.MiddleLeft;
            return l;
        }

        private void LoadTables()
        {
            for (int i = 0; i < _tables.Count; i++)
            {
                if (string.Equals(_tables[i], _detailTable, StringComparison.OrdinalIgnoreCase)) continue;
                cbMaster.Items.Add(_tables[i]);
            }

            for (int i = 0; i < _detailColumns.Count; i++) cbDetailKey.Items.Add(_detailColumns[i].Name);

            // 猜明细表的关联字段
            cbDetailKey.SelectedIndex = Guess(cbDetailKey, new string[] { "单据号", "单号", "销售单号", "销售单号", "billno", "bill_no", "bill", "orderno", "order_no", "no" });

            // 猜主表：名字像“主表/单头/销售”的优先
            int best = 0, bestScore = int.MinValue;
            for (int i = 0; i < cbMaster.Items.Count; i++)
            {
                string n = Convert.ToString(cbMaster.Items[i]).ToLowerInvariant();
                int sc = 0;
                if (n.IndexOf("主表") >= 0 || n.IndexOf("单头") >= 0 || n.IndexOf("master") >= 0 || n.IndexOf("head") >= 0) sc += 50;
                if (n.IndexOf("销售") >= 0 || n.IndexOf("sale") >= 0) sc += 30;
                if (n.IndexOf("明细") >= 0 || n.IndexOf("detail") >= 0) sc -= 60;
                if (n.IndexOf("入库") >= 0 || n.IndexOf("进货") >= 0) sc -= 40;
                if (sc > bestScore) { bestScore = sc; best = i; }
            }
            if (cbMaster.Items.Count > 0) cbMaster.SelectedIndex = best;
        }

        private void cbMaster_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbMaster.SelectedIndex < 0) return;
            string master = Convert.ToString(cbMaster.SelectedItem);
            List<ColumnInfo> cols;
            try { cols = _provider.ListColumns(master); }
            catch (Exception ex)
            {
                MessageBox.Show(this, "读取主表字段失败：" + ex.Message, "提示");
                return;
            }

            string detailKey = cbDetailKey.SelectedIndex >= 0 ? Convert.ToString(cbDetailKey.SelectedItem) : "";

            cbMasterKey.Items.Clear();
            cbMasterDate.Items.Clear();
            for (int i = 0; i < cols.Count; i++)
            {
                cbMasterKey.Items.Add(cols[i].Name);
                cbMasterDate.Items.Add(cols[i].Name);
            }

            int k = -1;
            for (int i = 0; i < cbMasterKey.Items.Count; i++)
            {
                if (string.Equals(Convert.ToString(cbMasterKey.Items[i]), detailKey, StringComparison.OrdinalIgnoreCase)) { k = i; break; }
            }
            if (k < 0) k = Guess(cbMasterKey, new string[] { "单据号", "单号", "billno", "bill", "no", "id" });
            if (cbMasterKey.Items.Count > 0) cbMasterKey.SelectedIndex = Math.Max(0, k);

            int d = Guess(cbMasterDate, new string[] { "销售日期", "日期", "业务日期", "单据日期", "开单日期", "时间", "date", "time" });
            if (cbMasterDate.Items.Count > 0) cbMasterDate.SelectedIndex = Math.Max(0, d);
        }

        private int Guess(ComboBox box, string[] keys)
        {
            for (int k = 0; k < keys.Length; k++)
            {
                for (int i = 0; i < box.Items.Count; i++)
                {
                    if (string.Equals(Convert.ToString(box.Items[i]), keys[k], StringComparison.OrdinalIgnoreCase)) return i;
                }
            }
            for (int k = 0; k < keys.Length; k++)
            {
                for (int i = 0; i < box.Items.Count; i++)
                {
                    string n = Convert.ToString(box.Items[i]).ToLowerInvariant();
                    if (n.IndexOf(keys[k].ToLowerInvariant()) >= 0) return i;
                }
            }
            return box.Items.Count > 0 ? 0 : -1;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (cbMaster.SelectedIndex < 0)
            {
                MessageBox.Show(this, "请选择主表。", "提示");
                return;
            }
            if (cbDetailKey.SelectedIndex < 0 || cbMasterKey.SelectedIndex < 0 || cbMasterDate.SelectedIndex < 0)
            {
                MessageBox.Show(this, "请把三个字段都选上。", "提示");
                return;
            }
            JoinSpec j = new JoinSpec();
            j.MasterTable = Convert.ToString(cbMaster.SelectedItem);
            j.DetailKey = Convert.ToString(cbDetailKey.SelectedItem);
            j.MasterKey = Convert.ToString(cbMasterKey.SelectedItem);
            j.DateColumn = Convert.ToString(cbMasterDate.SelectedItem);
            Result = j;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
