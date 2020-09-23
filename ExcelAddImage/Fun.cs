using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelAddImage
{
    class Fun
    {
        // show MessageBox
        public static void showMessageBox(string pMessage, string pCaption)
        {
            string message = pMessage;
            string caption = pCaption;
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result;

            result = MessageBox.Show(message, caption, buttons);
        }

        // 取圖片副檔名順序存入LIST
        public static List<string> getExtPriorityList(int p)
        {
            List<string> imagePriorityList = new List<string>();
            if (p == 1)
            {
                imagePriorityList.Add("pcx");
                imagePriorityList.Add("jpg");
            }
            else
            {
                imagePriorityList.Add("jpg");
                imagePriorityList.Add("pcx");
            }
            return imagePriorityList;
        }

        // 將欄位index轉換Excel column name ex: 1 -> "A"、2 -> "B"
        public static string GetExcelColumnNameFromNumber(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }

        // 將欄位Excel column name轉換index ex: "A" -> 1、"B" -> 2
        public static int GetNumberFromExcelColumn(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }

        // 圖片高度限制範圍30 ~ 100
        public static int imageHeightRange(int val)
        {
            if (val < 30)
                return 30;

            if (val > 100)
                return 100;

            return val;
        }

        // Set ShowSaveFileDialog
        public static ShowSaveFileDialogRe ShowSaveFileDialog(string tempFilename, FileInfo2 defaultFilename)
        {
            try
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel Files|*.xlsx";
                saveFileDialog1.FilterIndex = 1;
                saveFileDialog1.RestoreDirectory = true;
                saveFileDialog1.OverwritePrompt = true;

                // 預設檔案名稱
                if (defaultFilename != null)
                {
                    saveFileDialog1.FileName = defaultFilename.getGetFileNameWithoutExtension();
                    saveFileDialog1.InitialDirectory = defaultFilename.getDirectoryName();
                    saveFileDialog1.RestoreDirectory = true;
                }

                DialogResult dr = saveFileDialog1.ShowDialog();
                if (dr == DialogResult.OK && saveFileDialog1.FileName != "")
                {
                    File.Copy(tempFilename, saveFileDialog1.FileName, true);
                    // 回傳資訊
                    return new ShowSaveFileDialogRe
                        {
                            dialogResult = dr,
                            msg = saveFileDialog1.FileName
                        };
                }
                // 回傳資訊
                return new ShowSaveFileDialogRe
                {
                    dialogResult = dr,
                    msg = saveFileDialog1.FileName
                };
            }
            catch (IOException ex)
            {
                // 回傳資訊
                return new ShowSaveFileDialogRe
                {
                    dialogResult = DialogResult.Abort,
                    msg = ex.Message
                };
            }
        }

        // ExcelAddImage.exe程式執行路徑
        public static string getExecutingDirectory()
        {
            string executingPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            return Path.GetDirectoryName(executingPath);
        }

    }
}
