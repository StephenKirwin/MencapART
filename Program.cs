using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MencapART
{
    static class DocumentCarrier
    {
        public static int value = 3;
        public static string documentValue;
        public static bool reloadPage = false;
        public static bool pullPage = false;
    }

    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Thread thread = new Thread(runBrowser);
            thread.Start();
            Application.Run(new Form1());
        }

        static void runBrowser()
        {
            Application.Run(new Form2());
        }
    }
}
