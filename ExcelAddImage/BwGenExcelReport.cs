using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAddImage
{
    class BwGenExcelReport
    {
        // 1:ProgressBar、2:Error
        public int reportType { get; set; }

        // ProgressBar使用
        public bool begin { get; set; }
        public int maximum { get; set; }
        public int value { get; set; }

        // ProgressBar、Error 使用
        public string msg { get; set; }
    }
}
