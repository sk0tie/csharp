using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;

namespace WindowsFormsApplication1
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonBackup_Click(object sender, EventArgs e)
        {
            try
            {
                ConnectionOptions connOptions = new ConnectionOptions();
                connOptions.Username = "MAIN\\swilcox";
                connOptions.Password = "S67x41a!";
                connOptions.Impersonation = ImpersonationLevel.Impersonate;
                connOptions.EnablePrivileges = true;
                ManagementScope manScope = new ManagementScope(String.Format(@"\\{0}\ROOT\CIMV2", "sp-sea-01"), connOptions);
                manScope.Connect();
                ObjectGetOptions objectGetOptions = new ObjectGetOptions();
                ManagementPath managementPath = new ManagementPath("Win32_Process");
                ManagementClass processClass = new ManagementClass
                    (manScope, managementPath, objectGetOptions);
                ManagementBaseObject inParams = processClass.GetMethodParameters("Create");
                inParams["CommandLine"] = @"\\" + "sp-sea-01" + "\\c$\\ntbackup.bat";
                ManagementBaseObject outParams = processClass.InvokeMethod("Create", inParams, null);
                Console.WriteLine("Creation of the process returned: " + outParams["returnValue"]);
                Console.WriteLine("Process ID: " + outParams["processId"]);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
