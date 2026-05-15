namespace SupportTools_Excel.User_Interface.Task_Panes
{
    partial class TaskPane_ActiveDirectory
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.teUserName = new DevExpress.XtraEditors.TextEdit();
            this.teLastname = new DevExpress.XtraEditors.TextEdit();
            this.teOutput = new DevExpress.XtraEditors.TextEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.btnFindName = new DevExpress.XtraEditors.SimpleButton();
            this.btnGetDomains = new DevExpress.XtraEditors.SimpleButton();
            this.btnAddUser = new DevExpress.XtraEditors.SimpleButton();
            this.btnGetAllUsers = new DevExpress.XtraEditors.SimpleButton();
            this.btnGetAllGroups = new DevExpress.XtraEditors.SimpleButton();
            this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
            this.teDefaultNamingContext = new DevExpress.XtraEditors.TextEdit();
            this.teDNSHostName = new DevExpress.XtraEditors.TextEdit();
            this.teADDomainName = new DevExpress.XtraEditors.TextEdit();
            this.btnGetDomainControllers = new DevExpress.XtraEditors.SimpleButton();
            this.btnGetGlobalCatalogs = new DevExpress.XtraEditors.SimpleButton();
            this.btnGetAllUsersPath = new DevExpress.XtraEditors.SimpleButton();
            this.btnGetAllGroupsPath = new DevExpress.XtraEditors.SimpleButton();
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.wucAD_Picker1 = new VNC.AD.User_Interface.User_Controls.wucAD_Picker();
            ((System.ComponentModel.ISupportInitialize)(this.teUserName.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.teLastname.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.teOutput.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.teDefaultNamingContext.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.teDNSHostName.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.teADDomainName.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // teUserName
            // 
            this.teUserName.Location = new System.Drawing.Point(115, 336);
            this.teUserName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.teUserName.Name = "teUserName";
            this.teUserName.Size = new System.Drawing.Size(133, 22);
            this.teUserName.TabIndex = 2;
            // 
            // teLastname
            // 
            this.teLastname.Location = new System.Drawing.Point(115, 368);
            this.teLastname.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.teLastname.Name = "teLastname";
            this.teLastname.Size = new System.Drawing.Size(133, 22);
            this.teLastname.TabIndex = 3;
            // 
            // teOutput
            // 
            this.teOutput.Location = new System.Drawing.Point(115, 436);
            this.teOutput.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.teOutput.Name = "teOutput";
            this.teOutput.Size = new System.Drawing.Size(133, 22);
            this.teOutput.TabIndex = 4;
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(25, 340);
            this.labelControl1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(58, 16);
            this.labelControl1.TabIndex = 5;
            this.labelControl1.Text = "Username";
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(25, 372);
            this.labelControl2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(55, 16);
            this.labelControl2.TabIndex = 6;
            this.labelControl2.Text = "Lastname";
            // 
            // btnFindName
            // 
            this.btnFindName.Location = new System.Drawing.Point(115, 400);
            this.btnFindName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnFindName.Name = "btnFindName";
            this.btnFindName.Size = new System.Drawing.Size(100, 28);
            this.btnFindName.TabIndex = 7;
            this.btnFindName.Text = "FindName";
            this.btnFindName.Click += new System.EventHandler(this.btnFineName_Click);
            // 
            // btnGetDomains
            // 
            this.btnGetDomains.Location = new System.Drawing.Point(5, 574);
            this.btnGetDomains.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetDomains.Name = "btnGetDomains";
            this.btnGetDomains.Size = new System.Drawing.Size(187, 28);
            this.btnGetDomains.TabIndex = 8;
            this.btnGetDomains.Text = "Get Domains";
            this.btnGetDomains.Click += new System.EventHandler(this.btnGetDomains_Click);
            // 
            // btnAddUser
            // 
            this.btnAddUser.Location = new System.Drawing.Point(5, 528);
            this.btnAddUser.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAddUser.Name = "btnAddUser";
            this.btnAddUser.Size = new System.Drawing.Size(100, 28);
            this.btnAddUser.TabIndex = 10;
            this.btnAddUser.Text = "AddUser";
            this.btnAddUser.Click += new System.EventHandler(this.btnAddUser_Click);
            // 
            // btnGetAllUsers
            // 
            this.btnGetAllUsers.Location = new System.Drawing.Point(7, 222);
            this.btnGetAllUsers.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetAllUsers.Name = "btnGetAllUsers";
            this.btnGetAllUsers.Size = new System.Drawing.Size(100, 28);
            this.btnGetAllUsers.TabIndex = 11;
            this.btnGetAllUsers.Text = "Get All Users";
            this.btnGetAllUsers.Click += new System.EventHandler(this.btnGetAllUsers_Click);
            // 
            // btnGetAllGroups
            // 
            this.btnGetAllGroups.Location = new System.Drawing.Point(115, 219);
            this.btnGetAllGroups.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetAllGroups.Name = "btnGetAllGroups";
            this.btnGetAllGroups.Size = new System.Drawing.Size(100, 28);
            this.btnGetAllGroups.TabIndex = 12;
            this.btnGetAllGroups.Text = "Get All Groups";
            this.btnGetAllGroups.Click += new System.EventHandler(this.btnGetAllGroups_Click);
            // 
            // textEdit1
            // 
            this.textEdit1.Location = new System.Drawing.Point(5, 257);
            this.textEdit1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textEdit1.Name = "textEdit1";
            this.textEdit1.Size = new System.Drawing.Size(208, 22);
            this.textEdit1.TabIndex = 13;
            // 
            // teDefaultNamingContext
            // 
            this.teDefaultNamingContext.Location = new System.Drawing.Point(53, 187);
            this.teDefaultNamingContext.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.teDefaultNamingContext.Name = "teDefaultNamingContext";
            this.teDefaultNamingContext.Size = new System.Drawing.Size(236, 22);
            this.teDefaultNamingContext.TabIndex = 14;
            // 
            // teDNSHostName
            // 
            this.teDNSHostName.Location = new System.Drawing.Point(53, 155);
            this.teDNSHostName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.teDNSHostName.Name = "teDNSHostName";
            this.teDNSHostName.Size = new System.Drawing.Size(236, 22);
            this.teDNSHostName.TabIndex = 15;
            // 
            // teADDomainName
            // 
            this.teADDomainName.Location = new System.Drawing.Point(53, 123);
            this.teADDomainName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.teADDomainName.Name = "teADDomainName";
            this.teADDomainName.Size = new System.Drawing.Size(236, 22);
            this.teADDomainName.TabIndex = 16;
            // 
            // btnGetDomainControllers
            // 
            this.btnGetDomainControllers.Location = new System.Drawing.Point(5, 609);
            this.btnGetDomainControllers.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetDomainControllers.Name = "btnGetDomainControllers";
            this.btnGetDomainControllers.Size = new System.Drawing.Size(187, 28);
            this.btnGetDomainControllers.TabIndex = 17;
            this.btnGetDomainControllers.Text = "Get Domain Controllers";
            this.btnGetDomainControllers.Click += new System.EventHandler(this.btnGetDomainControllers_Click);
            // 
            // btnGetGlobalCatalogs
            // 
            this.btnGetGlobalCatalogs.Location = new System.Drawing.Point(5, 645);
            this.btnGetGlobalCatalogs.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetGlobalCatalogs.Name = "btnGetGlobalCatalogs";
            this.btnGetGlobalCatalogs.Size = new System.Drawing.Size(187, 28);
            this.btnGetGlobalCatalogs.TabIndex = 18;
            this.btnGetGlobalCatalogs.Text = "Get Global Catalogs";
            this.btnGetGlobalCatalogs.Click += new System.EventHandler(this.btnGetGlobalCatalogs_Click);
            // 
            // btnGetAllUsersPath
            // 
            this.btnGetAllUsersPath.Location = new System.Drawing.Point(7, 289);
            this.btnGetAllUsersPath.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetAllUsersPath.Name = "btnGetAllUsersPath";
            this.btnGetAllUsersPath.Size = new System.Drawing.Size(100, 28);
            this.btnGetAllUsersPath.TabIndex = 19;
            this.btnGetAllUsersPath.Text = "Get All Users";
            this.btnGetAllUsersPath.Click += new System.EventHandler(this.btnGetAllUsersPath_Click);
            // 
            // btnGetAllGroupsPath
            // 
            this.btnGetAllGroupsPath.Location = new System.Drawing.Point(115, 289);
            this.btnGetAllGroupsPath.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetAllGroupsPath.Name = "btnGetAllGroupsPath";
            this.btnGetAllGroupsPath.Size = new System.Drawing.Size(100, 28);
            this.btnGetAllGroupsPath.TabIndex = 20;
            this.btnGetAllGroupsPath.Text = "Get All Groups";
            // 
            // elementHost1
            // 
            this.elementHost1.Location = new System.Drawing.Point(0, 4);
            this.elementHost1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(289, 112);
            this.elementHost1.TabIndex = 21;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.wucAD_Picker1;
            // 
            // TaskPane_ActiveDirectory
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.elementHost1);
            this.Controls.Add(this.btnGetAllGroupsPath);
            this.Controls.Add(this.btnGetAllUsersPath);
            this.Controls.Add(this.btnGetGlobalCatalogs);
            this.Controls.Add(this.btnGetDomainControllers);
            this.Controls.Add(this.teADDomainName);
            this.Controls.Add(this.teDNSHostName);
            this.Controls.Add(this.teDefaultNamingContext);
            this.Controls.Add(this.textEdit1);
            this.Controls.Add(this.btnGetAllGroups);
            this.Controls.Add(this.btnGetAllUsers);
            this.Controls.Add(this.btnAddUser);
            this.Controls.Add(this.btnGetDomains);
            this.Controls.Add(this.btnFindName);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.teOutput);
            this.Controls.Add(this.teLastname);
            this.Controls.Add(this.teUserName);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MinimumSize = new System.Drawing.Size(293, 0);
            this.Name = "TaskPane_ActiveDirectory";
            this.Size = new System.Drawing.Size(293, 741);
            this.Load += new System.EventHandler(this.TaskPane_ActiveDirectory_Load);
            ((System.ComponentModel.ISupportInitialize)(this.teUserName.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.teLastname.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.teOutput.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.teDefaultNamingContext.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.teDNSHostName.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.teADDomainName.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private DevExpress.XtraEditors.TextEdit teUserName;
        private DevExpress.XtraEditors.TextEdit teLastname;
        private DevExpress.XtraEditors.TextEdit teOutput;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.SimpleButton btnFindName;
        private DevExpress.XtraEditors.SimpleButton btnGetDomains;
        private DevExpress.XtraEditors.SimpleButton btnAddUser;
        private DevExpress.XtraEditors.SimpleButton btnGetAllUsers;
        private DevExpress.XtraEditors.SimpleButton btnGetAllGroups;
        private DevExpress.XtraEditors.TextEdit textEdit1;
        private DevExpress.XtraEditors.TextEdit teDefaultNamingContext;
        private DevExpress.XtraEditors.TextEdit teDNSHostName;
        private DevExpress.XtraEditors.TextEdit teADDomainName;
        private DevExpress.XtraEditors.SimpleButton btnGetDomainControllers;
        private DevExpress.XtraEditors.SimpleButton btnGetGlobalCatalogs;
        private DevExpress.XtraEditors.SimpleButton btnGetAllUsersPath;
        private DevExpress.XtraEditors.SimpleButton btnGetAllGroupsPath;
        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private VNC.AD.User_Interface.User_Controls.wucAD_Picker wucAD_Picker1;
    }
}
