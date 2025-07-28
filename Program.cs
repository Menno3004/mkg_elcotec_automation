using Mkg_Elcotec_Automation.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mkg_Elcotec_Automation
{
    static class Program
    {
        [STAThread]
        static void Main()
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var form = new Elcotec();
            Application.Run(form);
        }
    }
}
