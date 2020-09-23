using ImageMagick;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ExcelAddImage
{
    class BwGenExcel
    {
        // Magick.NET轉檔、output Excel暫存目錄
        private string workTempDirectory;

        BwGenExcelArgs bwGenExcelArgs;

        private String jpgTempFile = "ExcelAddImageTemp{0}.jpg";

        public BwGenExcel(BwGenExcelArgs pBwGenExcelArgs)
        {
            this.bwGenExcelArgs = pBwGenExcelArgs;

            workTempDirectory = Path.GetTempPath();

            // C:\Users\Administrator\AppData\Local\Temp\ExcelAddImage\
            workTempDirectory = Path.Combine(Path.GetTempPath(), @"ExcelAddImage\");

            //Debug.WriteLine("workTempDirectory: " + workTempDirectory);

            Directory.CreateDirectory(workTempDirectory);

            DirectoryInfo di = new DirectoryInfo(workTempDirectory);
        }

        // BackgroundWorker run
        public void DoWork(object sender, DoWorkEventArgs e)
        {
            var bw = sender as BackgroundWorker;

            int imagePathColumnIdx;
            int addImageColumnIdx;

            try
            {
                // 產生暫存Excel
                string outExcelFileName = @"ExcelAddImageTemp.xlsx";
                string outExcelFilePath = String.Format("{0}{1}", workTempDirectory, outExcelFileName);

                string addExcelImage;

                // EPPLUS只可以用在 *.xlsx
                FileInfo newFile = new FileInfo(this.bwGenExcelArgs.sourceFile);
                ExcelPackage pck = new ExcelPackage(newFile);
                var worksheet = pck.Workbook.Worksheets[1];

                imagePathColumnIdx = this.bwGenExcelArgs.excelImagePathColumnIdx;
                addImageColumnIdx = this.bwGenExcelArgs.excelAddImageColumnIdx;

                // 2018.8.6 增加可插入在欄位的任何位置
                // 勾選放置圖片的欄位前插入1欄位
                if (this.bwGenExcelArgs.excelColumnInsert)
                {
                    worksheet.InsertColumn(addImageColumnIdx, 1);

                    // 放置圖片的欄位 <= path column --> path column + 1
                    // 放置圖片的欄位 > path column --> path column 不變
                    if (addImageColumnIdx <= imagePathColumnIdx)
                        imagePathColumnIdx++;
                }

                //放圖片的欄寬
                worksheet.Column(addImageColumnIdx).Width = 30;

                // 資料筆數
                var rowCount = worksheet.Dimension.End.Row;

                // 進度
                bw.ReportProgress(0,
                    new BwGenExcelReport
                    {
                        reportType = 1,
                        begin = true,
                        maximum = rowCount,
                        value = 0,
                        msg = "Working ..."
                    }
                );

                String msgText;
                for (int i = 1; i <= rowCount; i++)
                {
                    // 抓取鞋型名稱
                    String shoeid;
                    if (worksheet.Cells[i, imagePathColumnIdx].Value != null)
                        shoeid = worksheet.Cells[i, imagePathColumnIdx].Value.ToString();
                    else
                        shoeid = "null";

                    msgText = String.Format("Row: {0}, Image: {1}", i, shoeid);

                    // 進度
                    bw.ReportProgress(0,
                        new BwGenExcelReport
                        {
                            reportType = 1,
                            begin = false,
                            maximum = 0,
                            value = i,
                            msg = msgText
                        }
                    );

                    // 設定列高 points = pixels * 72 / 96 (0.75) 加上0.03增加高度
                    worksheet.Row(i).Height = this.bwGenExcelArgs.imageHeight * 0.78;

                    // 圖片名稱路徑
                    //String picFullPath = analyzeImagePath(shoeid);
                    FileInfo2 picFullPath = new FileInfo2(analyzeImagePath(shoeid));

                    // 檢查有無圖檔
                    if (File.Exists(picFullPath.getFullName()))
                    {
                        addExcelImage = rotationTempFile(i);

                        try
                        {
                            // 如果是jpg檔不需要轉檔直接copy到暫存目錄
                            if (picFullPath.isJPG())
                            {
                                // 2018.8.1 nancy反應winxp出現"insufficient image data in file "G:\PROD\XXXX.jpg"錯誤
                                // 原因為jpg檔轉jpg檔
                                File.Copy(picFullPath.getFullName(), addExcelImage, true);
                            }
                            else
                            {
                                // Read first frame of pcx image
                                using (MagickImage image = new MagickImage(picFullPath.getFullName()))
                                {
                                    // Save frame as jpg
                                    image.Write(addExcelImage);
                                }
                            }

                            // 抓取圖片長寬
                            Image img = Image.FromFile(addExcelImage);

                            // 插入圖片
                            var picture = worksheet.Drawings.AddPicture(Generate15UniqueDigits(), img);

                            picture.SetPosition(i - 1, 2, addImageColumnIdx - 1, 2);

                            // 縮放比例
                            double h = img.Height;
                            double p = (this.bwGenExcelArgs.imageHeight / h) * 100;
                            picture.SetSize(Convert.ToInt32(p));

                            img.Dispose();
                            img = null;

                        }
                        catch (Exception ex)
                        {
                            // 回報錯誤
                            bw.ReportProgress(0,
                                new BwGenExcelReport
                                {
                                    reportType = 2,
                                    begin = false,
                                    maximum = 0,
                                    value = i,
                                    msg = string.Format("{0}: {1}", "Error 1", ex.Message)
                                }
                            );
                        }
                    }
                    else
                    {
                        if (worksheet.Cells[i, addImageColumnIdx].Value == null)
                            worksheet.Cells[i, addImageColumnIdx].Value = "Image not found";
                    }

                }
                pck.SaveAs(new FileInfo(outExcelFilePath));

                // Control UI (Status label)
                msgText = String.Format("Finished..");

                // 進度
                bw.ReportProgress(0,
                    new BwGenExcelReport
                    {
                        reportType = 1,
                        begin = false,
                        maximum = 0,
                        value = rowCount,
                        msg = msgText
                    }
                );

                pck.Dispose();

                e.Result = outExcelFilePath;
            }
            catch (Exception ex)
            {
                // 回報錯誤
                bw.ReportProgress(0,
                    new BwGenExcelReport
                    {
                        reportType = 2,
                        begin = false,
                        maximum = 0,
                        value = 0,
                        msg = String.Format("{0}: {1}", "Error 2", ex.Message)
                    }
                );
            }
            finally
            {
                GC.Collect();
            }
        }

        // 解析圖片欄位內容
        private string analyzeImagePath(string val)
        {
            // 欄位內容可以為以下格式資料
            // 狀況1: G:\PROD\NG69.PCX
            // 狀況2: G:\PROD\NG69
            // 狀況3: NG69.PCX
            // 狀況4: NG69

            // 回傳值
            string returnFile = val;

            // 磁碟機代號
            string pattern1 = @"^[a-zA-Z]:\\";

            // 副檔名
            string pattern2 = @".jpg$|.pcx$";

            bool pattern1IsMatch = Regex.IsMatch(val, pattern1, RegexOptions.IgnoreCase);
            bool pattern2IsMatch = Regex.IsMatch(val, pattern2, RegexOptions.IgnoreCase);

            // 沒有輸入路徑的補上完整路徑
            if (!pattern1IsMatch)
                returnFile = string.Format(@"{0}\{1}", this.bwGenExcelArgs.imageDirectory, val);

            // 有輸入副檔名的直接回傳
            if (pattern2IsMatch)
                return returnFile;

            // 沒有輸入副檔名,依取檔順序檢查
            foreach (string v in this.bwGenExcelArgs.imagePriorityList)
            {
                string addExt = string.Format(@"{0}.{1}", returnFile, v);
                if (File.Exists(addExt))
                    return addExt;
            }

            return val;
        }

        // Temp file檔名,使用10個檔案循環,避免Lock
        private String rotationTempFile(int idx)
        {
            return workTempDirectory + String.Format(jpgTempFile, idx % 10);
        }

        // 產生UniqueDigits
        static object locker = new object();
        static string Generate15UniqueDigits()
        {
            lock (locker)
            {
                Thread.Sleep(2);
                return DateTime.Now.ToString("yyyyMMddHHmmssfff");
            }
        }
    }
}
