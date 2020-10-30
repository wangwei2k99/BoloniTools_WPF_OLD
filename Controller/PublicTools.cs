using Microsoft.Win32;
using System;
using BoloniTools.Controller;
namespace BoloniTools.Func
{
    public static class PublicTools
    {
        /// <param 文件列表="FileNames"></param>
        private static string[] fileNames;
        public static string[] FileNames
        {
            get
            {
                if (fileNames == null || fileNames[0] == "")
                {
                    Notice.NoticeFunc("请先选择文件再进行其他操作！");
                    return null;
                }
                else
                {
                    return fileNames;
                }
            }
            set { fileNames = value; }
        }
        #region 文件选择框
        public static string[] SelectExcelFile()
        {
            OpenFileDialog dlg = new OpenFileDialog()
            {
                Multiselect = true,
                RestoreDirectory = true,
                Filter = "Excel文件|*.xls*"
            };
            if (dlg.ShowDialog() == true)
            {
                return dlg.FileNames;
            }
            return null;
        }
        #endregion
        #region 保存文件
        public static string SaveFile()
        {
            SaveFileDialog dlg = new SaveFileDialog()
            {
                Filter = "Excel 文件(*.xlsx)|*.xlsx|Excel 文件(*.xls)|*.xls",
                FilterIndex = 0,
                RestoreDirectory = true, //保存对话框是否记忆上次打开的目录
                                         //saveFileDialog.CreatePrompt = true;
                Title = "导出Excel文件到"
            };
            DateTime now = DateTime.Now;
            dlg.FileName = "汇总清单" + now.Year.ToString().PadLeft(2) + now.Month.ToString().PadLeft(2, '0') + now.Day.ToString().PadLeft(2, '0') + "-" + now.Hour.ToString().PadLeft(2, '0') + now.Minute.ToString().PadLeft(2, '0') + now.Second.ToString().PadLeft(2, '0');
            return dlg.FileName;
        }
        #endregion
    }
}