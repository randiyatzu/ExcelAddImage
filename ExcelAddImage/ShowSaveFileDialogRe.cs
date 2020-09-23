using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddImage
{
    // ShowSaveFileDialog回傳資訊
    class ShowSaveFileDialogRe
    {
        // Dialog動作
        public DialogResult dialogResult { get; set; }

        // 成功: 檔案
        // 失敗: 訊息
        public string msg { get; set; }
    }
}
