using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BoloniTools.Func
{
    public static class SummaryCabinet
    {
        #region 柜体汇总
        public static DataTable Cabinet_Summary(DataSet dss)
        {
            var dt = dss.Tables["柜体"];
            if (dt == null || dt.Rows.Count == 0)
            { return null; }
            Npoi.DeleteColumns(dt, 8, 18);
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                if (dt.Rows[i]["C10"] == DBNull.Value
                       || dt.Rows[i]["C11"] == DBNull.Value
                       || dt.Rows[i]["C15"] == DBNull.Value
                       || dt.Rows[i]["C15"].ToString() == "总数")
                    dt.Rows[i].Delete();
            }
            dt.AcceptChanges();
            var query = dt.AsEnumerable().GroupBy(
                c => new
                {
                    tzym = c["C8"],
                    gytz = c["C9"],
                    bjmc = c["C10"],
                    l = c["C11"],
                    w = c["C12"],
                    d = c["C13"],
                    cz = c["C16"],
                    fbms = c["C17"],
                    gybz = c["C18"]
                }
                ).Select(
                s => new
                {
                    tzym = s.Select(p => p["C8"]).First(),
                    gytz = s.Select(p => p["C9"]).First(),
                    bjmc = s.Select(p => p["C10"]).First(),
                    l = s.Select(p => p["C11"]).First(),
                    w = s.Select(p => p["C12"]).First(),
                    d = s.Select(p => p["C13"]).First(),
                    zs = s.Sum(p => Convert.ToInt32(p["C15"])),
                    cz = s.Select(p => p["C16"]).First(),
                    fbms = s.Select(p => p["C17"]).First(),
                    gybz = s.Select(p => p["C18"]).First()
                });
            //var query = from c in dt.AsEnumerable()
            //            group c by new
            //            {
            //                //图纸页码 工艺图纸    部件名称    长   宽   厚   数量  总数  材质  封边描述    工艺备注
            //                tzym = c["C8"],
            //                gytz = c["C9"],
            //                bjmc = c["C10"],
            //                l = c["C11"],
            //                w = c["C12"],
            //                d = c["C13"],
            //                cz = c["C16"],
            //                fbms = c["C17"],
            //                gybz = c["C18"],
            //            }
            //            into s
            //            select new
            //            {
            //                tzym = s.Select(p => p.Field<string>("C8")).First(),
            //                gytz = s.Select(p => p.Field<string>("C9")).First(),
            //                bjmc = s.Select(p => p.Field<string>("C10")).First(),
            //                l = s.Select(p => p.Field<string>("C11")).First(),
            //                w = s.Select(p => p.Field<string>("C12")).First(),
            //                d = s.Select(p => p.Field<string>("C13")).First(),
            //                zs = s.Sum(p => Convert.ToInt32(p.Field<string>("C15"))),
            //                cz = s.Select(p => p.Field<string>("C16")).First(),
            //                fbms = s.Select(p => p.Field<string>("C17")).First(),
            //                gybz = s.Select(p => p.Field<string>("C18")).First()
            //            };
            DataTable 柜体汇总 = dss.Tables["柜体"].Clone();
            query.ToList().ForEach(p => 柜体汇总.Rows.Add(p.tzym, p.gytz, p.bjmc, p.l, p.w, p.d, null, p.zs, p.cz, p.fbms, p.gybz));
            return 柜体汇总;
        }
        //public static DataTable InputData(string[] FileNames)
        //{
        //    DataTable 柜体汇总 = new DataTable();
        //    DataSet dss = new DataSet();
        //    foreach (string File in FileNames)
        //    {
        //        DataSet ds = Npoi.ExcelToDataSet(File);
        //        if (ds.Tables.Contains("柜体(标准版)"))
        //        {
        //            ds.Tables.Remove("柜体(标准版)");
        //        }
        //        dss.Merge(ds);
        //    }
        //    柜体汇总 = SummaryCabinet.Cabinet_Summary(dss);
        //    if (柜体汇总 == null)
        //    {
        //        return null;
        //    }
        //    柜体汇总.Columns["C8"].SetOrdinal(10);
        //    柜体汇总.Columns.Add("序号", typeof(int)); 柜体汇总.Columns["序号"].SetOrdinal(0);
        //    柜体汇总.Columns.Add("空列1", typeof(string)); 柜体汇总.Columns["空列1"].SetOrdinal(1);
        //    柜体汇总 = Npoi.ColumnsToDecimal(柜体汇总, 4, 5, 6, 8);
        //    柜体汇总.DefaultView.Sort = "C12,C10,C11";
        //    柜体汇总 = 柜体汇总.DefaultView.ToTable();
        //    string path = Path.GetDirectoryName(FileNames[0]) + @"\柜体汇总.xlsx";
        //    Npoi.SetSerialNumber(柜体汇总, "序号");
        //    Npoi.DataTableToTemplate(柜体汇总, 5, 0, "柜体汇总", path, 4, 5, 6, 8);
        //    //System.Diagnostics.Process.Start(path);
        //    return 柜体汇总;
        //}
        #endregion
        public static DataTable Hardware_Summary(DataSet dss)
        {
            var dt = dss.Tables["Page1"];
            if (dt == null || dt.Rows.Count == 0)
            { return null; }
            Npoi.DeleteColumns(dt, 3, 15);
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                if (dt.Rows[i]["C5"] == DBNull.Value|| !Regex.IsMatch(dt.Rows[i]["C11"].ToString(), @"[0-9]+"))
                    dt.Rows[i].Delete();
            }
            dt.AcceptChanges();
            dt.Columns.Remove("C4");
            dt.Columns.Remove("C6");
            dt.Columns.Remove("C7");
            dt.Columns.Remove("C8");
            dt.Columns.Remove("C12");
            dt.Columns.Remove("C14");

            var query = dt.AsEnumerable().GroupBy(
                c => new
                {
                    WLBM = c["C3"],
                    WLMC = c["C5"],
                    ZSL = c["C11"],
                    UNIT = c["C13"],
                    BZ = c["C15"]
                }
                ).Select(
                s => new
                {
                    WLBM = s.Select(p => p["C3"]).First(),
                    WLMC = s.Select(p => p["C5"]).First(),
                    ZSL = s.Sum(p => Convert.ToInt32(p["C11"])),
                    UNIT = s.Select(p => p["C13"]).First(),
                    BZ = s.Select(p => p["C15"]).First()
                });
            DataTable 五金汇总 = dt.Clone();
            query.ToList().ForEach(p => 五金汇总.Rows.Add(p.WLBM, p.WLMC, null, null, p.ZSL,p.UNIT, p.BZ));
            return 五金汇总;
        }
        public static DataTable InputData(string[] FileNames)
        {
            DataTable 五金汇总 = new DataTable();
            DataSet dss = new DataSet();
            foreach (string File in FileNames)
            {
                DataSet ds = Npoi.ExcelToDataSet(File);
                if (ds.Tables.Contains("柜体(标准版)"))
                {
                    ds.Tables.Remove("柜体(标准版)");
                }
                dss.Merge(ds);
            }
            五金汇总 = SummaryCabinet.Hardware_Summary(dss);
            if (五金汇总 == null)
            {
                return null;
            }
            五金汇总.Columns.Add("序号", typeof(int)); 五金汇总.Columns["序号"].SetOrdinal(0);
            五金汇总 = Npoi.ColumnsToDecimal(五金汇总,10,11);
            string path = Path.GetDirectoryName(FileNames[0]) + @"\五金汇总.xlsx";
            Npoi.SetSerialNumber(五金汇总, "序号");
            Npoi.DataTableToTemplate(五金汇总, 6, 0, "五金汇总", path,10,11);
            //System.Diagnostics.Process.Start(path);
            return 五金汇总;
        }
    }
}