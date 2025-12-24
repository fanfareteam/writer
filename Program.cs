using System;
using System.Windows.Forms;

namespace ProjectWriter
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            // Pass command line arguments (if any) to Form1
            Application.Run(new Form1(args));
        }
    }
}