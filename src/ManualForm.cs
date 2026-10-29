using System;
using System.Drawing;
using System.Windows.Forms;

namespace DrugFlow
{
    /// <summary>手动指定数据源（自动扫描找不到时的备选）</summary>
    public class ManualForm : Form
    {
        private ComboBox cbType;
        private TextBox txtConn;
        private Label lblHint;
        private Button btnBrowse, btnTest, btnOk, btnCancel;
        public SourceItem Result;

        public ManualForm()
        {
            this.Text = "手动连接数据库";
            this.ClientSize = new Size(620, 300);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei UI", 9F);

            this.Controls.Add(Lab("类型", 16, 20, 40));
            cbType = new ComboBox();
            cbType.SetBounds(60, 16, 540, 24);
            cbType.DropDownStyle = ComboBoxStyle.DropDownList;
            cbType.Items.Add("SQL Server 数据库（最常见）");
            cbType.Items.Add("Access 数据库文件（.mdb / .accdb）");
            cbType.Items.Add("SQLite 数据库文件（.db / .sqlite）");
            cbType.Items.Add("其它 ODBC 数据源（Firebird / MySQL 等）");
            cbType.SelectedIndex = 0;
            cbType.SelectedIndexChanged += cbType_SelectedIndexChanged;
            this.Controls.Add(cbType);

            this.Controls.Add(Lab("连接", 16, 58, 40));
            txtConn = new TextBox();
            txtConn.SetBounds(60, 54, 430, 24);
            this.Controls.Add(txtConn);

            btnBrowse = new Button();
            btnBrowse.Text = "浏览…";
            btnBrowse.SetBounds(496, 53, 104, 26);
            btnBrowse.Click += btnBrowse_Click;
            this.Controls.Add(btnBrowse);

            lblHint = new Label();
            lblHint.SetBounds(16, 88, 584, 130);
            lblHint.ForeColor = Color.FromArgb(90, 98, 110);
            this.Controls.Add(lblHint);

            btnTest = new Button();
            btnTest.Text = "测试连接";
            btnTest.SetBounds(16, 238, 100, 30);
            btnTest.Click += btnTest_Click;
            this.Controls.Add(btnTest);

            btnOk = new Button();
            btnOk.Text = "确定使用";
            btnOk.SetBounds(392, 238, 100, 30);
            btnOk.Click += btnOk_Click;
            this.Controls.Add(btnOk);

            btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.SetBounds(500, 238, 100, 30);
            btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(btnCancel);

            cbType_SelectedIndexChanged(null, null);
        }

        private static Label Lab(string text, int x, int y, int w)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(x, y, w, 20);
            l.TextAlign = ContentAlignment.MiddleLeft;
            return l;
        }

        private void cbType_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool file = cbType.SelectedIndex == 1 || cbType.SelectedIndex == 2;
            btnBrowse.Visible = file;
            txtConn.Width = file ? 430 : 540;
            if (cbType.SelectedIndex == 1)
            {
                lblHint.Text = "选择一个 .mdb 或 .accdb 文件。\n通常就在卖药软件的安装目录里。";
                txtConn.Text = "";
            }
            else if (cbType.SelectedIndex == 2)
            {
                lblHint.Text = "选择一个 .db / .sqlite 文件。\nWindows 10 自带 SQLite 运行库，可以直接读；\nWin7 需要软件目录里带 sqlite3.dll。";
                txtConn.Text = "";
            }
            else if (cbType.SelectedIndex == 0)
            {
                lblHint.Text = "填写 SQL Server 连接字符串，例如：\n" +
                    "Data Source=.\\SQLEXPRESS;Initial Catalog=药房数据库;Integrated Security=SSPI\n\n" +
                    "如果数据库要求用户名密码，用：\n" +
                    "Data Source=192.168.1.10;Initial Catalog=xxx;User ID=sa;Password=***\n\n" +
                    "不确定的话，先点“自动扫描”，或者直接生成诊断报告发我。";
                if (txtConn.Text.Length == 0)
                    txtConn.Text = "Data Source=.\\SQLEXPRESS;Integrated Security=SSPI";
            }
            else
            {
                lblHint.Text = "填写 ODBC 连接字符串，例如 Firebird：\n" +
                    "Driver=Firebird/InterBase(r) driver;Dbname=C:\\data\\drug.fdb;Uid=SYSDBA;Pwd=masterkey\n\n" +
                    "注意：本机必须已安装对应的 ODBC 驱动（32 位），否则连接不上。";
                if (txtConn.Text.IndexOf("Data Source", StringComparison.OrdinalIgnoreCase) >= 0)
                    txtConn.Text = "";
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (cbType.SelectedIndex == 2)
            {
                ofd.Title = "选择 SQLite 数据库文件";
                ofd.Filter = "SQLite 数据库 (*.db;*.sqlite;*.db3)|*.db;*.sqlite;*.db3|所有文件 (*.*)|*.*";
            }
            else
            {
                ofd.Title = "选择 Access 数据库文件";
                ofd.Filter = "Access 数据库 (*.mdb;*.accdb)|*.mdb;*.accdb|所有文件 (*.*)|*.*";
            }
            if (ofd.ShowDialog(this) == DialogResult.OK) txtConn.Text = ofd.FileName;
        }

        private string BuildConnection()
        {
            string t = txtConn.Text.Trim();
            if (t.Length == 0) return "";
            if (cbType.SelectedIndex == 1)
            {
                return t;
            }
            return t;
        }

        private SourceItem MakeSource()
        {
            string conn = BuildConnection();
            if (conn.Length == 0) return null;
            SourceItem s = new SourceItem();
            if (cbType.SelectedIndex == 0) s.Kind = "sqlserver";
            else if (cbType.SelectedIndex == 1) s.Kind = "access";
            else if (cbType.SelectedIndex == 2) s.Kind = "sqlite";
            else s.Kind = "odbc";
            s.Connection = conn;
            s.Path = conn;
            s.Title = "手动：" + (s.Kind == "sqlserver" ? Tx.Short(conn, 70) : System.IO.Path.GetFileName(conn));
            return s;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            SourceItem s = MakeSource();
            if (s == null)
            {
                MessageBox.Show(this, "请先填写连接信息。", "提示");
                return;
            }
            Cursor = Cursors.WaitCursor;
            try
            {
                using (IProvider p = DbFactory.Create(s))
                {
                    System.Collections.Generic.List<string> tables = p.ListTables();
                    MessageBox.Show(this, "连接成功！\n共找到 " + tables.Count + " 张表。\n\n" + Tx.Short(Tx.Join(tables.ToArray(), "，"), 300),
                        "测试成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "连接失败：\n" + ex.Message, "失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Cursor = Cursors.Default; }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            SourceItem s = MakeSource();
            if (s == null)
            {
                MessageBox.Show(this, "请先填写连接信息。", "提示");
                return;
            }
            Result = s;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
