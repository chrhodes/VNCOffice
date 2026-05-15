using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelHlp = VNC.AddinHelper.Excel;
using VNC;

// TODO
// Do something clever so only have to walk the worksheet one time looking for matching rows.

namespace SupportTools_Excel.User_Interface.Task_Panes
{
    public partial class TaskPane_ActiveDirectory : UserControl
    {
        #region Constructors and Load

 
        public TaskPane_ActiveDirectory()
        {
            InitializeComponent();
            LoadControlContents();
        }
        
        private void TaskPane_ActiveDirectory_Load(object sender, EventArgs e)
        {
            wucAD_Picker1.ControlChanged += adPicker1_ControlChanged;
        }

        #endregion

        #region Constants

         #endregion

        #region Fields and Properties


        #endregion

        #region Event Handlers

        private void adPicker1_ControlChanged()
        {
            teADDomainName.Text = wucAD_Picker1.Name;
            teDefaultNamingContext.Text = wucAD_Picker1.DefaultNamingContext;
            teDNSHostName.Text = wucAD_Picker1.DNSHostName;
        }

        private void btnAddUser_Click(object sender, EventArgs e)
        {
            //VNC.AD.Helper
        }

        private void btnFineName_Click(object sender, EventArgs e)
        {
            teOutput.Text = VNC.ActiveDirectory.Helper.FindName(teUserName.Text);
        }

        private void btnGetAllGroups_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.GetAllADDomainGroups();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetAllUsers_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.GetAllADDomainUsers();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetAllUsersPath_Click(object sender, EventArgs e)
        {
            string domainPath = textEdit1.Text;

            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.GetAllADDomainUsers(domainPath);

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetDomainControllers_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.EnumerateDomainControllers();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        private void btnGetDomains_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.EnumerateDomains();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }


        private void btnGetGlobalCatalogs_Click(object sender, EventArgs e)
        {
            System.Collections.ArrayList al = VNC.ActiveDirectory.Helper.EnumerateGlobalCatalogs();

            StringBuilder sb = new StringBuilder();

            foreach (var item in al)
            {
                sb.AppendLine(item.ToString());
            }

            MessageBox.Show(sb.ToString());
        }

        #endregion

        #region Main Function Routines


        private void foo()
        {
            
        }

        #endregion

        #region Private Methods
        
        private void LoadControlContents()
        {
            wucAD_Picker1.PopulateControlFromFile(Common.cCONFIG_FILE);
        }

        #endregion
  
    }
}
