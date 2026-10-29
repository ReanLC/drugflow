using System;
using System.Threading;
using System.Windows.Forms;

namespace DrugFlow
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.ThreadException += delegate(object sender, ThreadExceptionEventArgs e)
            {
                ShowError(e.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            {
                Exception ex = e.ExceptionObject as Exception;
                if (ex != null) ShowError(ex);
            };

            Application.Run(new MainForm());
        }

        private static void ShowError(Exception ex)
        {
            try
            {
                MessageBox.Show("程序遇到一个问题：\n\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch { }
        }
    }
}
