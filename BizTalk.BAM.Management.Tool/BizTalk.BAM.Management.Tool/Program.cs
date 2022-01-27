using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BizTalk.BAM.Management.Tool
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
            BAMMgmtUtlForm frm = new BAMMgmtUtlForm();
            Application.Run(frm);
        }
    }
}