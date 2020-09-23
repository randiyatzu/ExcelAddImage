using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelAddImage
{
    class FileInfo2
    {
        // 取檔資訊
        private FileInfo filename;

        public FileInfo2(string pFilename)
        {
            filename = new FileInfo(pFilename);
        }

        // 取完整路徑+檔名
        public string getFullName()
        {
            return filename.FullName;
        }

        // 取目錄名稱
        public string getDirectoryName()
        {
            return filename.DirectoryName;
        }

        // 取檔案名稱+副檔名
        public string getName()
        {
            return filename.Name;
        }

        // 取檔案名稱不含副檔名
        public string getGetFileNameWithoutExtension()
        {
            return Path.GetFileNameWithoutExtension(getName());
        }

        // 回傳檔案的目錄存在否
        public bool isDirectoryExists()
        {
            if (Directory.Exists(getDirectoryName()))
                return true;
            return false;
        }

        // 回傳檔案的檔案存在否
        public bool isFileExists()
        {
            if (File.Exists(getFullName()))
                return true;
            return false;
        }

        // 是否為jpg
        public bool isJPG()
        {
            string pattern = ".jpg$";
            return Regex.IsMatch(getName(), pattern, RegexOptions.IgnoreCase);
        }

        // 是否為xlsx檔
        public bool isExcel()
        {
            string pattern = ".xlsx$";
            return Regex.IsMatch(getName(), pattern, RegexOptions.IgnoreCase);
        }
    }
}
