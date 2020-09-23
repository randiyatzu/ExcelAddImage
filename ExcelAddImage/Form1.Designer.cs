namespace ExcelAddImage
{
    partial class Main
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
            this.browseFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.fromFilePath = new System.Windows.Forms.Label();
            this.toExcel = new System.Windows.Forms.Button();
            this.status = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.trackBarImageHeight = new System.Windows.Forms.TrackBar();
            this.label3 = new System.Windows.Forms.Label();
            this.picExample = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.outputExcelFileName = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.openExcel = new System.Windows.Forms.Button();
            this.cbPathColumn = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.rbExt1 = new System.Windows.Forms.RadioButton();
            this.rbExt2 = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label1ImageDirectory = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.cbImageColumn = new System.Windows.Forms.ComboBox();
            this.chKBoxInsColumn = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.trackBarImageHeight)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picExample)).BeginInit();
            this.SuspendLayout();
            // 
            // browseFile
            // 
            this.browseFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.browseFile.Location = new System.Drawing.Point(34, 42);
            this.browseFile.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.browseFile.Name = "browseFile";
            this.browseFile.Size = new System.Drawing.Size(180, 50);
            this.browseFile.TabIndex = 0;
            this.browseFile.Text = "Open excel";
            this.browseFile.UseVisualStyleBackColor = true;
            this.browseFile.Click += new System.EventHandler(this.browseFile_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(33, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "File name:";
            // 
            // fromFilePath
            // 
            this.fromFilePath.AutoSize = true;
            this.fromFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fromFilePath.Location = new System.Drawing.Point(119, 101);
            this.fromFilePath.Name = "fromFilePath";
            this.fromFilePath.Size = new System.Drawing.Size(95, 16);
            this.fromFilePath.TabIndex = 2;
            this.fromFilePath.Text = "fromFilePath";
            // 
            // toExcel
            // 
            this.toExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toExcel.Location = new System.Drawing.Point(34, 432);
            this.toExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.toExcel.Name = "toExcel";
            this.toExcel.Size = new System.Drawing.Size(180, 50);
            this.toExcel.TabIndex = 3;
            this.toExcel.Text = "Add image";
            this.toExcel.UseVisualStyleBackColor = true;
            this.toExcel.Click += new System.EventHandler(this.toExcel_Click);
            // 
            // status
            // 
            this.status.AutoSize = true;
            this.status.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.status.Location = new System.Drawing.Point(249, 431);
            this.status.Name = "status";
            this.status.Size = new System.Drawing.Size(51, 16);
            this.status.TabIndex = 4;
            this.status.Text = "Status";
            // 
            // progressBar
            // 
            this.progressBar.Cursor = System.Windows.Forms.Cursors.Default;
            this.progressBar.Location = new System.Drawing.Point(253, 450);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(309, 23);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progressBar.TabIndex = 6;
            // 
            // trackBarImageHeight
            // 
            this.trackBarImageHeight.Location = new System.Drawing.Point(351, 172);
            this.trackBarImageHeight.Maximum = -1;
            this.trackBarImageHeight.Minimum = -100;
            this.trackBarImageHeight.Name = "trackBarImageHeight";
            this.trackBarImageHeight.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.trackBarImageHeight.Size = new System.Drawing.Size(45, 157);
            this.trackBarImageHeight.TabIndex = 7;
            this.trackBarImageHeight.TabStop = false;
            this.trackBarImageHeight.TickStyle = System.Windows.Forms.TickStyle.None;
            this.trackBarImageHeight.Value = -100;
            this.trackBarImageHeight.ValueChanged += new System.EventHandler(this.trackBarImageHeight_ValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.label3.Location = new System.Drawing.Point(430, 165);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 16);
            this.label3.TabIndex = 8;
            this.label3.Text = "Image height";
            // 
            // picExample
            // 
            this.picExample.Location = new System.Drawing.Point(380, 185);
            this.picExample.Name = "picExample";
            this.picExample.Size = new System.Drawing.Size(192, 131);
            this.picExample.TabIndex = 11;
            this.picExample.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 3);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(190, 22);
            this.label4.TabIndex = 17;
            this.label4.Text = "Step 1 Choose File:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(12, 389);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(160, 22);
            this.label5.TabIndex = 18;
            this.label5.Text = "Step 3: To Excel";
            // 
            // outputExcelFileName
            // 
            this.outputExcelFileName.AutoSize = true;
            this.outputExcelFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.outputExcelFileName.Location = new System.Drawing.Point(158, 545);
            this.outputExcelFileName.Name = "outputExcelFileName";
            this.outputExcelFileName.Size = new System.Drawing.Size(155, 16);
            this.outputExcelFileName.TabIndex = 20;
            this.outputExcelFileName.Text = "outputExcelFileName";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.label6.Location = new System.Drawing.Point(33, 545);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 16);
            this.label6.TabIndex = 19;
            this.label6.Text = "Save file name:";
            // 
            // openExcel
            // 
            this.openExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.openExcel.Location = new System.Drawing.Point(36, 565);
            this.openExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.openExcel.Name = "openExcel";
            this.openExcel.Size = new System.Drawing.Size(180, 50);
            this.openExcel.TabIndex = 4;
            this.openExcel.Text = "Open Excel";
            this.openExcel.UseVisualStyleBackColor = true;
            this.openExcel.Click += new System.EventHandler(this.openExcel_Click);
            // 
            // cbPathColumn
            // 
            this.cbPathColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPathColumn.FormattingEnabled = true;
            this.cbPathColumn.Location = new System.Drawing.Point(155, 174);
            this.cbPathColumn.Name = "cbPathColumn";
            this.cbPathColumn.Size = new System.Drawing.Size(158, 23);
            this.cbPathColumn.TabIndex = 2;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(12, 140);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(145, 22);
            this.label7.TabIndex = 23;
            this.label7.Text = "Step 2: Setting";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(33, 177);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(84, 16);
            this.label8.TabIndex = 24;
            this.label8.Text = "Path column:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Arial", 14.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(12, 514);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(160, 22);
            this.label9.TabIndex = 25;
            this.label9.Text = "Step 4: Finished";
            // 
            // rbExt1
            // 
            this.rbExt1.AutoSize = true;
            this.rbExt1.Checked = true;
            this.rbExt1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.rbExt1.Location = new System.Drawing.Point(155, 287);
            this.rbExt1.Name = "rbExt1";
            this.rbExt1.Size = new System.Drawing.Size(52, 20);
            this.rbExt1.TabIndex = 26;
            this.rbExt1.TabStop = true;
            this.rbExt1.Text = "PCX";
            this.rbExt1.UseVisualStyleBackColor = true;
            // 
            // rbExt2
            // 
            this.rbExt2.AutoSize = true;
            this.rbExt2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.rbExt2.Location = new System.Drawing.Point(155, 312);
            this.rbExt2.Name = "rbExt2";
            this.rbExt2.Size = new System.Drawing.Size(52, 20);
            this.rbExt2.TabIndex = 27;
            this.rbExt2.Text = "JPG";
            this.rbExt2.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(33, 287);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 16);
            this.label2.TabIndex = 28;
            this.label2.Text = "Image priority:";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(33, 352);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(104, 16);
            this.label10.TabIndex = 29;
            this.label10.Text = "Image directory:";
            // 
            // label1ImageDirectory
            // 
            this.label1ImageDirectory.AutoSize = true;
            this.label1ImageDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1ImageDirectory.Location = new System.Drawing.Point(158, 352);
            this.label1ImageDirectory.Name = "label1ImageDirectory";
            this.label1ImageDirectory.Size = new System.Drawing.Size(101, 16);
            this.label1ImageDirectory.TabIndex = 30;
            this.label1ImageDirectory.Text = "Image directory";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(33, 216);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(95, 16);
            this.label11.TabIndex = 31;
            this.label11.Text = "Image column:";
            // 
            // cbImageColumn
            // 
            this.cbImageColumn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbImageColumn.FormattingEnabled = true;
            this.cbImageColumn.Location = new System.Drawing.Point(161, 209);
            this.cbImageColumn.Name = "cbImageColumn";
            this.cbImageColumn.Size = new System.Drawing.Size(158, 23);
            this.cbImageColumn.TabIndex = 32;
            // 
            // chKBoxInsColumn
            // 
            this.chKBoxInsColumn.AutoSize = true;
            this.chKBoxInsColumn.Checked = true;
            this.chKBoxInsColumn.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chKBoxInsColumn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F);
            this.chKBoxInsColumn.Location = new System.Drawing.Point(161, 238);
            this.chKBoxInsColumn.Name = "chKBoxInsColumn";
            this.chKBoxInsColumn.Size = new System.Drawing.Size(105, 20);
            this.chKBoxInsColumn.TabIndex = 33;
            this.chKBoxInsColumn.Text = "Insert column";
            this.chKBoxInsColumn.UseVisualStyleBackColor = true;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(597, 628);
            this.Controls.Add(this.chKBoxInsColumn);
            this.Controls.Add(this.cbImageColumn);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label1ImageDirectory);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.rbExt2);
            this.Controls.Add(this.rbExt1);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cbPathColumn);
            this.Controls.Add(this.openExcel);
            this.Controls.Add(this.outputExcelFileName);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.picExample);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.trackBarImageHeight);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.status);
            this.Controls.Add(this.toExcel);
            this.Controls.Add(this.fromFilePath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.browseFile);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Main";
            this.Text = "Excel add image";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.trackBarImageHeight)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picExample)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button browseFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label fromFilePath;
        private System.Windows.Forms.Button toExcel;
        private System.Windows.Forms.Label status;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox picExample;
        private System.Windows.Forms.TrackBar trackBarImageHeight;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label outputExcelFileName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button openExcel;
        private System.Windows.Forms.ComboBox cbPathColumn;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.RadioButton rbExt1;
        private System.Windows.Forms.RadioButton rbExt2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label1ImageDirectory;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox cbImageColumn;
        private System.Windows.Forms.CheckBox chKBoxInsColumn;
    }
}

