using IniParser;
using IniParser.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddImage
{
    public partial class Main2 : Form
    {
        // command line參數
        private string[] args;

        // BackgroundWorker參數class
        private BwGenExcelArgs bwGenExcelArgs;

        // 產生Excel BackgroundWorker
        private BackgroundWorker bw;

        // 鞋圖資料夾
        private String imageDirectory;

        // 取圖片副檔名順序
        private String imagePriority;

        // 開啟的Excel檔
        private FileInfo2 sourceFile;

        // 另存新檔的檔案物件
        private FileInfo2 saveAsFile;

        // 是否要顯示SaveAs DialogBox
        private bool showSaveAsDialogBox;

        // Excel新增圖片的column是否要insert一欄
        private bool excelColumnInsert;

        public Main2(string[] Args)
        {
            InitializeComponent();
            this.args = Args;
        }

        private void Main2_Load(object sender, EventArgs e)
        {
            string version = System.Windows.Forms.Application.ProductVersion;
            this.Text = String.Format("Excel add image by command line (version {0})", version);

            outputExcelFileName.Text = "";
            openExcel.Visible = false;

            if (!(args.Length == 8))
            {
                string msg = @"
EXAMPLE:
    ExcelAddImage.exe ""D:\te\source_excel"" ""D:\te\output_excel"" Y B C Y 1 80

    parameter 1 : Source excel file (excluded .xlsx)
    parameter 2 : Destination excel file (excluded .xlsx)
    parameter 3 : Y or N, Show the SaveAs Dialog Box
    parameter 4 : Image path column from source excel file
    parameter 5 : Add image column from source excel file
    parameter 6 : Y or N, Insert a blank column to add image
    parameter 7 : 1 or 2 (PCX or JPG), Get image priority
    parameter 8 : Image range of height from 30 to 100
";
                Fun.showMessageBox(msg, "Error : Incorrect number of parameter");
                Application.Exit();
            }

            // 檢查圖片高度參數範圍,超過範圍跳出訊息並關閉程式
            if (Int32.Parse(args[7]) != Fun.imageHeightRange(Int32.Parse(args[7])))
            {
                Fun.showMessageBox(
                    String.Format("Check image range of height from 30 to 100"), "Error");
                Application.Exit();
            }

            sourceFile = new FileInfo2(string.Format("{0}.xlsx", args[0]));
            saveAsFile = new FileInfo2(string.Format("{0}.xlsx", args[1]));

            if (!sourceFile.isFileExists())
            {
                Fun.showMessageBox(
                    String.Format("\"{0}\" source excel does not exist.",
                        sourceFile.getFullName()), "SaveAs error");
                Application.Exit();
            }

            // 程式執行路徑
            string executingDirectory = Fun.getExecutingDirectory();

            // 讀取ini
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile(String.Format(@"{0}\{1}", executingDirectory, "Config.ini"));
            imageDirectory = data["ExcelAddImage"]["ImageDirectory"];
            imagePriority = data["ExcelAddImage"]["ImagePriority"];

            // 檢查圖片路徑存在否
            if (!Directory.Exists(imageDirectory))
                Fun.showMessageBox(String.Format("\"{0}\" does not exist.", imageDirectory), "Error");

            // 顯示SaveAs Dialog Box
            if (args[2].ToUpper().Equals("Y"))
                showSaveAsDialogBox = true;
            else
                showSaveAsDialogBox = false;

            // Excel新增圖片的column是否要insert一欄
            if (args[5].ToUpper().Equals("Y"))
                excelColumnInsert = true;
            else
                excelColumnInsert = false;

            // BwGenExcel參數Object
            bwGenExcelArgs = new BwGenExcelArgs
            {
                sourceFile = sourceFile.getFullName(),
                outputFile = saveAsFile.getFullName(),
                excelImagePathColumnIdx = Fun.GetNumberFromExcelColumn(args[3]),
                imageHeight = Int32.Parse(args[7]),
                imagePriorityList = Fun.getExtPriorityList(Int32.Parse(args[6])),
                imageDirectory = imageDirectory,
                excelAddImageColumnIdx = Fun.GetNumberFromExcelColumn(args[4]),
                excelColumnInsert = excelColumnInsert
            };

            // 線程產生Excel
            BwGenExcel bw_DoWork = new BwGenExcel(bwGenExcelArgs);
            bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(bw_DoWork.DoWork);
            bw.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
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
                string finishedFileName = null;
                try
                {
                    if (showSaveAsDialogBox)
                    {
                        RunShowSaveFileDialogRe(e.Result.ToString());
                    }
                    else
                    {
                        if (saveAsFile.isDirectoryExists())
                        {
                            finishedFileName = bwGenExcelArgs.outputFile;
                            File.Copy(e.Result.ToString(), finishedFileName, true);
                            SetOutputExcelFileName(finishedFileName);
                            openExcel.Visible = true;
                        }
                        else
                        {
                            // 指定目錄不存在要跳出MessageBox
                            Fun.showMessageBox(
                                String.Format("\"{0}\" destination directory does not exist.",
                                    saveAsFile.getDirectoryName()), "SaveAs error");
                            RunShowSaveFileDialogRe(e.Result.ToString());
                        }
                    }
                }
                catch (IOException ex)
                {
                    Fun.showMessageBox(
                        string.Format("{0}",
                        ex.Message), "SaveAs error");
                }

            }
        }

        // 存檔視窗
        private void RunShowSaveFileDialogRe(string tempFilename) {
            ShowSaveFileDialogRe showSaveFileDialogRe = Fun.ShowSaveFileDialog(tempFilename, saveAsFile);
            // 存檔按鈕
            if (showSaveFileDialogRe.dialogResult == DialogResult.OK)
            {
                string finishedFileName = showSaveFileDialogRe.msg;

                SetOutputExcelFileName(finishedFileName);
                openExcel.Visible = true;
            }
            else if (showSaveFileDialogRe.dialogResult == DialogResult.Abort)
            {
                Fun.showMessageBox(
                    string.Format("{0}",
                    showSaveFileDialogRe.msg), "SaveAs error");
            }
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

        // Set status label
        private void SetText(string text)
        {
            this.status.Text = text;
        }

        // Set output file label
        private void SetOutputExcelFileName(string filename)
        {
            outputExcelFileName.Text = filename;
        }

        private void openExcel_Click(object sender, EventArgs e)
        {
            // open saveAS Excel
            if (File.Exists(outputExcelFileName.Text))
                System.Diagnostics.Process.Start(outputExcelFileName.Text);
        }
    }
}
