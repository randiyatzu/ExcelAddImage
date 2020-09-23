using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using IniParser;
using IniParser.Model;
using System.Diagnostics;
using OfficeOpenXml;
using System.Collections.Generic;
using System.ComponentModel;

namespace ExcelAddImage
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        // 拖曳進入Form event
        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        // 拖曳放入檔案event
        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                FileInfo2 f = new FileInfo2(file);
                // 取第1個檔案
                if (f.isExcel())
                {
                    bwGenExcelArgs.sourceFile = f.getFullName();
                    fromFilePath.Text = bwGenExcelArgs.sourceFile;

                    initFormWhenOpenExcel();
                    setComboboxItem(bwGenExcelArgs.sourceFile);

                    break;
                }
            }
        }

        // BackgroundWorker參數class
        private BwGenExcelArgs bwGenExcelArgs = new BwGenExcelArgs();

        // 產生Excel BackgroundWorker
        private BackgroundWorker bw;

        // 取圖片副檔名順序
        private String imagePriority;

        // 範例圖片
        private string exampleImagePath;

        // 儲存範例圖片的高、寬
        private int picExampleHeight;
        private int picExampleWidth;

        // Button按下的時間
        private double buttonClickMilliseconds = 0;

        private void Form1_Load(object sender, EventArgs e)
        {
            string version = System.Windows.Forms.Application.ProductVersion;
            this.Text = String.Format("Excel add image (version {0})", version);

            fromFilePath.Text = "";
            outputExcelFileName.Text = "";

            // 程式執行路徑
            string executingDirectory = Fun.getExecutingDirectory();

            // 2018.7.31 增加
            // Nancy 反應在WinXP上縮放圖片會出現"Parameter is not valid"錯誤,所以增加完整路徑指定範例圖片
            exampleImagePath = String.Format(@"{0}\{1}", executingDirectory, "example.jpg");

            // 讀取ini
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile(String.Format(@"{0}\{1}", executingDirectory, "Config.ini"));
            bwGenExcelArgs.imageDirectory = data["ExcelAddImage"]["ImageDirectory"];
            imagePriority = data["ExcelAddImage"]["ImagePriority"];

            // 檢查圖片路徑存在否
            if (!Directory.Exists(bwGenExcelArgs.imageDirectory))
                Fun.showMessageBox(String.Format("\"{0}\" does not exist.", bwGenExcelArgs.imageDirectory), "Error");

            label1ImageDirectory.Text = bwGenExcelArgs.imageDirectory;

            // 儲存範例圖片的高、寬
            picExampleHeight = picExample.Height;
            picExampleWidth = picExample.Width;

            ShowExampleImage(exampleImagePath, picExampleWidth, picExampleHeight);

            // 取圖片副檔名順序
            if (imagePriority.Equals("1"))
                rbExt1.Checked = true;
            else
                rbExt2.Checked = true;
        }

        // 1.選擇Excel檔案
        private void browseFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            // Set filter options and filter index.
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.FilterIndex = 1;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                bwGenExcelArgs.sourceFile = openFileDialog1.FileName;
                fromFilePath.Text = bwGenExcelArgs.sourceFile;

                initFormWhenOpenExcel();
                setComboboxItem(bwGenExcelArgs.sourceFile);
            }
        }

        // 開啟檔案時候部份物件設定為預設值
        private void initFormWhenOpenExcel()
        {
            SetText("Status");
            SetOutputExcelFileName("");
            SetProgressBar(true, 100, 0);
        }

        // 設定Combobox item
        private void setComboboxItem(string file)
        {
            // 將有資料的欄位名稱新增到Combobox供使用者選擇圖片路徑是哪一個欄位

            // EPPLUS只可以用在 *.xlsx
            FileInfo newFile = new FileInfo(file);
            ExcelPackage pck = new ExcelPackage(newFile);
            var worksheet = pck.Workbook.Worksheets[1];

            cbPathColumn.Items.Clear();
            cbImageColumn.Items.Clear();
            for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
            {
                ComboboxItem item = new ComboboxItem();
                item.Text = String.Format("{0} - {1}",
                    Fun.GetExcelColumnNameFromNumber(col), worksheet.Cells[1, col].Value);
                item.Value = col;
                cbPathColumn.Items.Add(item);
                cbImageColumn.Items.Add(item);
            }
            cbPathColumn.SelectedIndex = 0;
            cbImageColumn.SelectedIndex = 0;
            pck.Dispose();
        }

        // Set output file label
        private void SetOutputExcelFileName(string filename)
        {
            outputExcelFileName.Text = filename;
        }

        // Set status label
        private void SetText(string text)
        {
            this.status.Text = text;
        }

        // Set Button Status
        private void SetButtonStatus(bool b)
        {
            this.controlButton(b);
        }

        // Set ProgressBar
        private void SetProgressBar(bool begin, int maximum, int value)
        {
            if (begin)
            {
                this.progressBar.Maximum = maximum;
                this.progressBar.Value = value;
            }
            else
            {
                this.progressBar.Value = value;
            }
        }

        // Button enable or disable
        private void controlButton(bool enable)
        {
            browseFile.Enabled = enable;
            toExcel.Enabled = enable;
            openExcel.Enabled = enable;
            trackBarImageHeight.Enabled = enable;
            cbPathColumn.Enabled = enable;
            rbExt1.Enabled = enable;
            rbExt2.Enabled = enable;
            cbImageColumn.Enabled = enable;
            chKBoxInsColumn.Enabled = enable;
        }

         // 2.加入鞋圖
        private void toExcel_Click(object sender, EventArgs e)
        {
            // 2秒內避免重複按出錯
            DateTime localDateTime = DateTime.Now;
            double now = localDateTime.ToUniversalTime().Subtract(
                new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
            if (now - buttonClickMilliseconds < 2000)
            {
                return;
            }
            else
            {
                buttonClickMilliseconds = now;
            }

            if (bwGenExcelArgs.sourceFile == null || (!(File.Exists(bwGenExcelArgs.sourceFile))))
            {
                Fun.showMessageBox("Please choose a excel file", "Caution");
                return;
            }

            bwGenExcelArgs.imageHeight = Fun.imageHeightRange(Math.Abs(trackBarImageHeight.Value));

            // 取圖片副檔名順序存入LIST
            if (rbExt1.Checked)
                bwGenExcelArgs.imagePriorityList = Fun.getExtPriorityList(1);
            else
                bwGenExcelArgs.imagePriorityList = Fun.getExtPriorityList(2);

            // 取Excel圖片路徑的column
            bwGenExcelArgs.excelImagePathColumnIdx = Int32.Parse(
                (cbPathColumn.SelectedItem as ComboboxItem).Value.ToString()
            );

            // Excel新增圖片的column
            bwGenExcelArgs.excelAddImageColumnIdx = Int32.Parse(
                (cbImageColumn.SelectedItem as ComboboxItem).Value.ToString()
            );

            // Excel新增圖片的column是否要insert一欄
            bwGenExcelArgs.excelColumnInsert = chKBoxInsColumn.Checked;

            // 線程產生Excel
            BwGenExcel bw_DoWork = new BwGenExcel(bwGenExcelArgs);
            bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork.DoWork);
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            // disable button
            this.SetButtonStatus(false);

            bw.RunWorkerAsync();
        }

        // BackgroundWorker 更新ui
        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            BwGenExcelReport state = e.UserState as BwGenExcelReport;

            if (state.reportType == 1)
            {
                SetProgressBar(state.begin, state.maximum, state.value);
                SetText(state.msg);
            }
            else
            {
                Fun.showMessageBox(state.msg, "Error");
            }
        }

        // BackgroundWorker 執行完成
        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Result != null)
            {

                ShowSaveFileDialogRe showSaveFileDialogRe = Fun.ShowSaveFileDialog(e.Result.ToString(), null);

                // 存檔按鈕
                if (showSaveFileDialogRe.dialogResult == DialogResult.OK)
                {
                    string finishedFileName = showSaveFileDialogRe.msg;

                    SetOutputExcelFileName(finishedFileName);
                    // 如果是覆蓋原檔案,必需要重新讀取ComboboxItem
                    if (finishedFileName != null)
                        if (finishedFileName.Equals(bwGenExcelArgs.sourceFile))
                        {
                            setComboboxItem(finishedFileName);
                        }
                }
                else if (showSaveFileDialogRe.dialogResult == DialogResult.Abort)
                {
                    Fun.showMessageBox(
                        string.Format("{0}",
                        showSaveFileDialogRe.msg), "SaveAs error");
                }
            }
            this.SetButtonStatus(true);
        }

        // 顯示範例圖片
        private Bitmap exampleImage;
        public void ShowExampleImage(String fileToDisplay, int xSize, int ySize)
        {
            // Sets up an image object to be displayed.
            if (exampleImage != null)
            {
                exampleImage.Dispose();
            }

            // Stretches the image to fit the pictureBox.
            picExample.SizeMode = PictureBoxSizeMode.Zoom;

            exampleImage = new Bitmap(fileToDisplay);
            picExample.ClientSize = new Size(xSize, ySize);

            picExample.Image = (Image)exampleImage;
        }

        // 動態縮放圖片
        private void trackBarImageHeight_ValueChanged(object sender, EventArgs e)
        {
            // 目前刻度
            double val = Math.Abs(((System.Windows.Forms.TrackBar)(sender)).Value);

            val = Fun.imageHeightRange(Convert.ToInt32(val));

            // 縮放比例
            double scale = val / 100;
            int clientXSize = Convert.ToInt32(picExampleWidth * scale);
            int clientYSize = Convert.ToInt32(picExampleHeight * scale);

            // 範例圖片
            ShowExampleImage(exampleImagePath, clientXSize, clientYSize);
        }

        // open saveAS Excel
        private void openExcel_Click(object sender, EventArgs e)
        {
            if (File.Exists(outputExcelFileName.Text))
                System.Diagnostics.Process.Start(outputExcelFileName.Text);
        }
    }
}
