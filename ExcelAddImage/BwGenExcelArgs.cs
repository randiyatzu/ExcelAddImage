using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddImage
{
    class BwGenExcelArgs
    {
        // 原始Excel檔
        public string sourceFile { get; set; }

        // 產生Excel檔
        public string outputFile { get; set; }

        // Excel放圖片路徑的column
        public int excelImagePathColumnIdx { get; set; }

        // Excel新增圖片的column
        public int excelAddImageColumnIdx { get; set; }

        // Excel新增圖片的column是否要insert一欄
        public bool excelColumnInsert { get; set; }

        // 圖片設定的高度
        public int imageHeight { get; set; }

        // 取圖片副檔名順序
        public List<string> imagePriorityList { get; set; }

        // 預設圖片目錄
        public string imageDirectory { get; set; }
    }
}
