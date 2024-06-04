namespace VNC.VSTOAddIn.Presentation
{
    partial class frmDebugWindow
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
            if(disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnClearOutput = new System.Windows.Forms.Button();
            this.gbDebugOptions = new System.Windows.Forms.GroupBox();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.gbDebugOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClearOutput
            // 
            this.btnClearOutput.Location = new System.Drawing.Point(18, 18);
            this.btnClearOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnClearOutput.Name = "btnClearOutput";
            this.btnClearOutput.Size = new System.Drawing.Size(267, 35);
            this.btnClearOutput.TabIndex = 7;
            this.btnClearOutput.Text = "Clear Output";
            this.btnClearOutput.UseVisualStyleBackColor = true;
            this.btnClearOutput.Click += new System.EventHandler(this.btnClearOutput_Click);
            // 
            // gbDebugOptions
            // 
            this.gbDebugOptions.Location = new System.Drawing.Point(18, 82);
            this.gbDebugOptions.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.gbDebugOptions.Name = "gbDebugOptions";
            this.gbDebugOptions.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.gbDebugOptions.Size = new System.Drawing.Size(267, 200);
            this.gbDebugOptions.TabIndex = 8;
            this.gbDebugOptions.TabStop = false;
            this.gbDebugOptions.Text = "Debug Options";
            // 
            // txtOutput
            // 
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOutput.Location = new System.Drawing.Point(294, 18);
            this.txtOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtOutput.Size = new System.Drawing.Size(862, 826);
            this.txtOutput.TabIndex = 6;
            // 
            // frmDebugWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1176, 865);
            this.Controls.Add(this.btnClearOutput);
            this.Controls.Add(this.gbDebugOptions);
            this.Controls.Add(this.txtOutput);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "frmDebugWindow";
            this.Text = "frmDebugWindow";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmDebugWindow_FormClosed);
            this.gbDebugOptions.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Button btnClearOutput;
        internal System.Windows.Forms.GroupBox gbDebugOptions;
        internal System.Windows.Forms.TextBox txtOutput;
    }
}