namespace SupportTools_Excel.User_Interface.Task_Panes
{
    partial class TaskPane_LogParser
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
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
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
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.wucTaskPane_LogParser1 = new SupportTools_Excel.User_Interface.User_Controls.wucTaskPane_LogParser();
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(400, 900);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.wucTaskPane_LogParser1;
            // 
            // TaskPane_LogParser
            // 
            this.Controls.Add(this.elementHost1);
            this.MinimumSize = new System.Drawing.Size(400, 900);
            this.Name = "TaskPane_LogParser";
            this.Size = new System.Drawing.Size(400, 900);
            this.ResumeLayout(false);
        }

        #endregion

        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private User_Controls.wucTaskPane_LogParser wucTaskPane_LogParser1;
    }
}
