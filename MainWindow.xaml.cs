using BoloniTools.Func;
using BoloniTools.View;
using MaterialDesignThemes.Wpf;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Path = System.IO.Path;
using BoloniTools.Controller;
namespace BoloniTools
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Notice.NoticeEvent += Notice_NoticeEvent;
        }

        private void Notice_NoticeEvent(object sender, string str)
        {
            SnackbarMessage message = new SnackbarMessage()
            {
                Content = str,
            };
            message.Content = str;
            this.snackbar1.Message = message;
            this.snackbar1.IsActive = true;
            DateTime current = DateTime.Now;
            while (current.AddMilliseconds(2000) > DateTime.Now)
            {
                System.Windows.Forms.Application.DoEvents();
            }
            this.snackbar1.IsActive = false;
        }

        #region 业务逻辑
        private void SelectFile(object sender, RoutedEventArgs e)//选择文件
        {
            PublicTools.FileNames = PublicTools.SelectExcelFile();
            DataTable Files = new DataTable();
            Files.Columns.Add("已选择的文件列表", typeof(string));
            if (!(PublicTools.FileNames is null))
            {
                foreach (var i in PublicTools.FileNames)
                {
                    Files.Rows.Add(Path.GetFileName(i));
                }
            }
            Grd1.ItemsSource = Files.AsDataView();
        }
        private void Summary(object sender, RoutedEventArgs e)//汇总
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            DataTable dt_gthz = SummaryCabinet.InputData(PublicTools.FileNames);
            if (dt_gthz == null || dt_gthz.Rows.Count == 0)
            {
                return;
            }
            else
            {
                Notice.NoticeFunc("汇总成功！已在当前目录生成文件！");
            }
            Grd1.ItemsSource = dt_gthz.AsDataView();

        }
        private void CabinetFlowCard(object sender, RoutedEventArgs e)
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            new FlowCard().InputFlowCardData(PublicTools.FileNames);
            if (StaticVariable.FlowCardDataTable == null || StaticVariable.FlowCardDataTable.Rows.Count == 0)
            {
                return;
            }
            var datatemp = StaticVariable.FlowCardDataTable.AsEnumerable().Where(p => !p["Class"].Equals("柜体")).ToList();
            datatemp.ForEach(p => StaticVariable.FlowCardDataTable.Rows.Remove(p));
            //DataRow[] dataRows = StaticVariable.FlowCardDataTable.Select("Class <> '柜体'");
            //foreach (DataRow dataRow in dataRows)
            //{
            //    dataRow.BeginEdit();
            //    dataRow.Delete();
            //    dataRow.EndEdit();
            //}
            //StaticVariable.FlowCardDataTable.AcceptChanges();
            var dtemp = new FlowCard().DrawingsPrint(StaticVariable.FlowCardDataTable);
            var totalnumber = new FlowCard().TotalNumber(StaticVariable.FlowCardDataTable);
            dtemp.Columns.Add("板件数量汇总", typeof(int));
            dtemp.Rows[0]["板件数量汇总"] = totalnumber;
            this.Grd1.ItemsSource = dtemp.AsDataView();
            FlowCardWindows fcw = new FlowCardWindows();
            fcw.ShowDialog();
        }
        private void DoorFlowCard(object sender, RoutedEventArgs e)
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            new FlowCard().InputFlowCardData(PublicTools.FileNames);
            if (StaticVariable.FlowCardDataTable == null || StaticVariable.FlowCardDataTable.Rows.Count == 0)
            {
                return;
            }
            DataRow[] dataRows = StaticVariable.FlowCardDataTable.Select("Class <> '门板'");
            foreach (DataRow dataRow in dataRows)
            {
                dataRow.BeginEdit();
                dataRow.Delete();
                dataRow.EndEdit();
            }
            StaticVariable.FlowCardDataTable.AcceptChanges();
            var dtemp = new FlowCard().DrawingsPrint(StaticVariable.FlowCardDataTable);
            var totalnumber = new FlowCard().TotalNumber(StaticVariable.FlowCardDataTable);
            dtemp.Columns.Add("板件数量汇总", typeof(int));
            dtemp.Rows[0]["板件数量汇总"] = totalnumber;
            this.Grd1.ItemsSource = dtemp.AsDataView();
            FlowCardWindows fcw = new FlowCardWindows();
            fcw.ShowDialog();
        }
        private void SingleSide(object sender, RoutedEventArgs e)
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            new FlowCard().InputFlowCardData(PublicTools.FileNames);
            if (StaticVariable.FlowCardDataTable == null || StaticVariable.FlowCardDataTable.Rows.Count == 0)
            {
                return;
            }
            DataRow[] dataRows = StaticVariable.FlowCardDataTable.Select("Class <> '单面吸塑'");
            foreach (DataRow dataRow in dataRows)
            {
                dataRow.BeginEdit();
                dataRow.Delete();
                dataRow.EndEdit();
            }
            StaticVariable.FlowCardDataTable.AcceptChanges();
            var dtemp = new FlowCard().DrawingsPrint(StaticVariable.FlowCardDataTable);
            var totalnumber = new FlowCard().TotalNumber(StaticVariable.FlowCardDataTable);
            dtemp.Columns.Add("板件数量汇总", typeof(int));
            dtemp.Rows[0]["板件数量汇总"] = totalnumber;
            this.Grd1.ItemsSource = dtemp.AsDataView();
            FlowCardWindows fcw = new FlowCardWindows();
            fcw.ShowDialog();
        }
        private void DoubleSide(object sender, RoutedEventArgs e)
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            new FlowCard().InputFlowCardData(PublicTools.FileNames);
            if (StaticVariable.FlowCardDataTable == null || StaticVariable.FlowCardDataTable.Rows.Count == 0)
            {
                return;
            }
            DataRow[] dataRows = StaticVariable.FlowCardDataTable.Select("Class <> '双面吸塑'");
            foreach (DataRow dataRow in dataRows)
            {
                dataRow.BeginEdit();
                dataRow.Delete();
                dataRow.EndEdit();
            }
            StaticVariable.FlowCardDataTable.AcceptChanges();
            var dtemp = new FlowCard().DrawingsPrint(StaticVariable.FlowCardDataTable);
            var totalnumber = new FlowCard().TotalNumber(StaticVariable.FlowCardDataTable);
            dtemp.Columns.Add("板件数量汇总", typeof(int));
            dtemp.Rows[0]["板件数量汇总"] = totalnumber;
            this.Grd1.ItemsSource = dtemp.AsDataView();
            FlowCardWindows fcw = new FlowCardWindows();
            fcw.ShowDialog();
        }
        private void FlowCard(object sender, RoutedEventArgs e)
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            new FlowCard().InputFlowCardData(PublicTools.FileNames);
            if (StaticVariable.FlowCardDataTable == null || StaticVariable.FlowCardDataTable.Rows.Count == 0)
            {
                return;
            }
            var dtemp = new FlowCard().DrawingsPrint(StaticVariable.FlowCardDataTable);
            var totalnumber = new FlowCard().TotalNumber(StaticVariable.FlowCardDataTable);
            dtemp.Columns.Add("板件数量汇总", typeof(int));
            dtemp.Rows[0]["板件数量汇总"] = totalnumber;
            this.Grd1.ItemsSource = dtemp.AsDataView();
            FlowCardWindows fcw = new FlowCardWindows();
            fcw.ShowDialog();
        }
        private void ConvertToMes1(object sender, RoutedEventArgs e)//汇总输出Mes
        {
            if (PublicTools.FileNames == null)
            {
                return;
            }
            DataTable dts = ConvertToMes.ConvertMes(PublicTools.FileNames);
            if (dts == null || dts.Rows.Count == 0)
            {
                return;
            }
            if (dts.Columns.Contains("项目编码")) dts.Columns.Remove("项目编码");
            if (dts.Columns.Contains("项目名称")) dts.Columns.Remove("项目名称");
            if (dts.Columns.Contains("订单编号")) dts.Columns.Remove("订单编号");
            if (dts.Columns.Contains("总套数")) dts.Columns.Remove("总套数");
            if (dts.Columns.Contains("总数量")) dts.Columns.Remove("总数量");
            if (dts.Columns.Contains("包装箱号")) dts.Columns.Remove("包装箱号");
            if (dts.Columns.Contains("预留1")) dts.Columns.Remove("预留1");
            if (dts.Columns.Contains("膜皮型号")) dts.Columns.Remove("膜皮型号");
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel 文件(*.xlsx)|*.xlsx|Excel 文件(*.xls)|*.xls",
                FilterIndex = 0,
                RestoreDirectory = true, //保存对话框是否记忆上次打开的目录
                Title = "导出Excel文件到"
            };
            DateTime now = DateTime.Now;
            saveFileDialog.FileName = $"Excel2Mes{now.Year,2}{now.Month,2}{now.Day,2}-{now.Hour,2}{now.Minute,2}{now.Second,2}";
            //点了保存按钮进入
            if (saveFileDialog.ShowDialog() == true)
            {
                if (saveFileDialog.FileName.Trim() == "")
                {
                    return;
                }
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(dts);
                Npoi.DataSetToExcel(dataSet, saveFileDialog.FileName);
                Notice.NoticeFunc("汇总并转换MES成功！已在当前目录生成文件！");
                Grd1.ItemsSource = dts.AsDataView();
                //System.Diagnostics.Process.Start(saveFileDialog.FileName);
            }
        }
        private void ConvertToMes2(object sender, RoutedEventArgs e)//单个输出Mes
        {
            if (PublicTools.FileNames == null) { return; }
            DataTable dts = ConvertToMes.ConvertMes(PublicTools.FileNames);
            if (dts == null || dts.Rows.Count == 0) { return; }
            if (dts.Columns.Contains("项目编码")) dts.Columns.Remove("项目编码");
            if (dts.Columns.Contains("项目名称")) dts.Columns.Remove("项目名称");
            if (dts.Columns.Contains("总套数")) dts.Columns.Remove("总套数");
            if (dts.Columns.Contains("总数量")) dts.Columns.Remove("总数量");
            if (dts.Columns.Contains("包装箱号")) dts.Columns.Remove("包装箱号");
            if (dts.Columns.Contains("预留1")) dts.Columns.Remove("预留1");
            if (dts.Columns.Contains("膜皮型号")) dts.Columns.Remove("膜皮型号");
            if (dts.Columns.Contains("工艺图纸")) dts.Columns.Remove("工艺图纸");
            if (dts.Columns.Contains("柜号")) dts.Columns.Remove("柜号");
            if (dts.Columns.Contains("柜宽")) dts.Columns.Remove("柜宽");
            if (dts.Columns.Contains("柜深")) dts.Columns.Remove("柜深");
            if (dts.Columns.Contains("柜高")) dts.Columns.Remove("柜高");
            if (dts.Columns.Contains("柜数量")) dts.Columns.Remove("柜数量");
            string FilePath = Path.GetDirectoryName(PublicTools.FileNames[0]);
            //点了保存按钮进入
            //var list0 = dts.AsEnumerable().GroupBy(g =>new {po = g["订单编号"] }).Select(s=>new { po = s.Select(p=>p["订单编号"]).First()}).ToList();
            var list = dts.AsEnumerable().Select(s => s["订单编号"].ToString()).Distinct().ToList();
            //List<string> list = (from r in dts.AsEnumerable() select r.Field<string>("订单编号")).Distinct().ToList<string>();
            foreach (string str in list)
            {
                string filename = FilePath + $@"\MES-{str}.xlsx";
                DataTable dt_temp = dts.AsEnumerable().Where(w => w["订单编号"].ToString().Equals(str)).CopyToDataTable();
                //DataTable dt_temp = (from r in dts.AsEnumerable() where r.Field<string>("订单编号") == str select r).CopyToDataTable();
                if (dt_temp.Columns.Contains("订单编号")) dt_temp.Columns.Remove("订单编号");
                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(dt_temp);
                Npoi.DataSetToExcel(dataSet, filename);
            }
            Notice.NoticeFunc("按合同号转换MES成功！已在当前目录生成文件！");
            if (dts.Columns.Contains("订单编号")) dts.Columns.Remove("订单编号");
            StaticVariable.MesDataTable = dts;
            Grd1.ItemsSource = dts.AsDataView();
        }
        #endregion
        #region 事件
        private void Shutdown(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void ButtonMax_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }
            else
            {
                this.WindowState = WindowState.Maximized;
            }
        }
        private void ButtonOpenMenu_Click(object sender, RoutedEventArgs e)
        {
            ButtonOpenMenu.Visibility = Visibility.Collapsed;
            ButtonCloseMenu.Visibility = Visibility.Visible;
        }
        private void ButtonCloseMenu_Click(object sender, RoutedEventArgs e)
        {
            ButtonOpenMenu.Visibility = Visibility.Visible;
            ButtonCloseMenu.Visibility = Visibility.Collapsed;
        }
        private void GridTitle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
        private void CloseMe_MouseEnter(object sender, MouseEventArgs e)
        {
            CloseMe.Background = new SolidColorBrush(Colors.Red);
        }
        private void CloseMe_MouseLeave(object sender, MouseEventArgs e)
        {
            CloseMe.Background = new SolidColorBrush();
        }
        private void PopupBox1Open(object sender, MouseButtonEventArgs e)
        {
            PopupBox1.IsPopupOpen = true;
        }
        #endregion
        private void Smile_Click(object sender, RoutedEventArgs e)
        {
            Notice.NoticeDemo();
        }
    }
}
