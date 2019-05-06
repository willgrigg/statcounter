namespace StatCounter
{
    partial class frmMain
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.btnLoadFile = new System.Windows.Forms.Button();
            this.tbFile1 = new System.Windows.Forms.TextBox();
            this.lblStats = new System.Windows.Forms.Label();
            this.tbFile2 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.Location = new System.Drawing.Point(18, 96);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.Size = new System.Drawing.Size(93, 41);
            this.btnLoadFile.TabIndex = 0;
            this.btnLoadFile.Text = "Load File";
            this.btnLoadFile.UseVisualStyleBackColor = true;
            this.btnLoadFile.Click += new System.EventHandler(this.BtnOpen_Click);
            // 
            // tbFile1
            // 
            this.tbFile1.Location = new System.Drawing.Point(18, 34);
            this.tbFile1.Name = "tbFile1";
            this.tbFile1.ReadOnly = true;
            this.tbFile1.Size = new System.Drawing.Size(681, 22);
            this.tbFile1.TabIndex = 1;
            // 
            // lblStats
            // 
            this.lblStats.AutoSize = true;
            this.lblStats.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStats.Location = new System.Drawing.Point(12, 141);
            this.lblStats.Name = "lblStats";
            this.lblStats.Size = new System.Drawing.Size(0, 32);
            this.lblStats.TabIndex = 2;
            // 
            // tbFile2
            // 
            this.tbFile2.Location = new System.Drawing.Point(18, 64);
            this.tbFile2.Name = "tbFile2";
            this.tbFile2.ReadOnly = true;
            this.tbFile2.Size = new System.Drawing.Size(681, 22);
            this.tbFile2.TabIndex = 3;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(711, 450);
            this.Controls.Add(this.tbFile2);
            this.Controls.Add(this.lblStats);
            this.Controls.Add(this.tbFile1);
            this.Controls.Add(this.btnLoadFile);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.Text = "Spreadsheet Counter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnLoadFile;
        private System.Windows.Forms.TextBox tbFile1;
        private System.Windows.Forms.Label lblStats;
        private System.Windows.Forms.TextBox tbFile2;
    }
}

