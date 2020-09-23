namespace ExcelAddImage
{
    partial class Main2
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
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.status = new System.Windows.Forms.Label();
            this.outputExcelFileName = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.openExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Cursor = System.Windows.Forms.Cursors.Default;
            this.progressBar.Location = new System.Drawing.Point(22, 32);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(472, 23);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 8;
            // 
            // status
            // 
            this.status.AutoSize = true;
            this.status.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.status.Location = new System.Drawing.Point(18, 13);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(51, 16);
            this.status.TabIndex = 7;
            this.status.Text = "Status";
            // 
            // outputExcelFileName
            // 
            this.outputExcelFileName.AutoSize = true;
            this.outputExcelFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.outputExcelFileName.Location = new System.Drawing.Point(144, 72);
            this.outputExcelFileName.Name = "outputExcelFileName";
            this.outputExcelFileName.Size = new System.Drawing.Size(155, 16);
            this.outputExcelFileName.TabIndex = 22;
            this.outputExcelFileName.Text = "outputExcelFileName";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.label6.Location = new System.Drawing.Point(19, 72);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 16);
            this.label6.TabIndex = 21;
            this.label6.Text = "Save file name:";
            // 
            // openExcel
            // 
            this.openExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openExcel.Location = new System.Drawing.Point(147, 106);
            this.openExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.openExcel.Name = "openExcel";
            this.openExcel.Size = new System.Drawing.Size(180, 38);
            this.openExcel.TabIndex = 23;
            this.openExcel.Text = "Open Excel";
            this.openExcel.UseVisualStyleBackColor = true;
            this.openExcel.Click += new System.EventHandler(this.openExcel_Click);
            // 
            // Main2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(517, 160);
            this.Controls.Add(this.openExcel);
            this.Controls.Add(this.outputExcelFileName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.status);
            this.Name = "Main2";
            this.Text = "Excel add image by command line";
            this.Load += new System.EventHandler(this.Main2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label status;
        private System.Windows.Forms.Label outputExcelFileName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button openExcel;
    }
}