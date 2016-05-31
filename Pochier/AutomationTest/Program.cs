using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading;

namespace AutomationTest
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Mutex MX;

            try
            {
                MX = new Mutex(false, "Pochier" + Application.ProductVersion);
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message);
                return;
            }

            //二重か
            if (!MX.WaitOne(1))
            {
                MessageBox.Show("二重起動です。");
                return;
            }

            Application.Run(new Pochier());
        }
    }
}
