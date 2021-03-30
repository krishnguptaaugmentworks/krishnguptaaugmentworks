using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace AgingReport
{
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
            AddOnInfo AOI = new AddOnInfo();
            AOI.StartAddOn();
            Application.Run();
        }
    }
}
