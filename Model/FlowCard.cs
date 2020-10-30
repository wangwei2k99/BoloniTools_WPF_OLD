using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
namespace BoloniTools.Func
{
    public class FlowCard
    {
        public delegate DataTable FlowCardType(DataTable dataTable);
        public delegate void GroupByRule(ref DataTable dataTable, DataRow dataRow, int depth, int nums, decimal length, decimal width);
        public DataTable Cabinet(DataTable dataTable)
        {
            string[] 柜体汇总 = new string[] { "拍子号", "预留2", "工艺图纸", "部件名称", "开料长", "开料宽", "厚", "预留1", "总数量", "材质", "封边描述", "工艺备注", "图纸页码" };
            DataTable dt = Npoi.DeleteColumns(dataTable, 1, 13);
            string 项目名称 = dt.Columns[0].ColumnName;
            for (int i = 0; i < dt.Columns.Count; i++){dt.Columns[i].ColumnName = 柜体汇总[i];}
            dt.Columns.Add("Class", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("标段", typeof(string));
            dt.Columns.Add("套数", typeof(string));
            dt.Columns.Add("拆图员", typeof(string));
            dt.Columns.Add("整单备注", typeof(string));
            Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
            Npoi.ColumnSetValue(ref dt, dt.Rows[1]["开料长"].ToString(), "标段");
            Npoi.ColumnSetValue(ref dt, dt.Rows[1]["材质"].ToString(), "套数");
            Npoi.ColumnSetValue(ref dt, dt.Rows[1]["工艺备注"].ToString(), "拆图员");
            Npoi.ColumnSetValue(ref dt, dt.Rows[0]["拍子号"].ToString(), "整单备注");
            Npoi.ColumnSetValue(ref dt, "柜体", "Class");
            return dt;
        }
        public DataTable Door(DataTable dataTable)
        {
            string[] 门板汇总 = new string[] { "拍子号", "图纸页码", "门型", "部件名称", "开料长", "开料宽", "长", "宽", "厚", "预留1", "总数量", "材质", "封边描述", "开门方向", "工艺备注" };
            DataTable dt = Npoi.DeleteColumns(dataTable, 1, 15);
            string 项目名称 = dt.Columns[0].ColumnName;
            for (int i = 0; i < dt.Columns.Count; i++){dt.Columns[i].ColumnName = 门板汇总[i];}
            dt.Columns.Add("Class", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("标段", typeof(string));
            dt.Columns.Add("套数", typeof(string));
            dt.Columns.Add("拆图员", typeof(string));
            dt.Columns.Add("整单备注", typeof(string));
            dt.Columns.Add("纹路", typeof(string));
            string 套数 = Regex.Match(dt.Rows[1]["总数量"].ToString(), @"[0-9]+").Groups[0].ToString();
            string 标段 = dt.Rows[1]["图纸页码"].ToString();
            string 拆图员 = dt.Rows[1]["工艺备注"].ToString();
            string 整单备注 = dt.Rows[0]["拍子号"].ToString();
            string 纹路 = dt.Rows[3]["开料长"].ToString();
            //string 套数 = dt.Rows[0]["拍子号"].ToString() == "标段" ? Regex.Match(dt.Rows[0]["总数量"].ToString(), @"[0-9]+").Groups[0].ToString() : Regex.Match(dt.Rows[1]["总数量"].ToString(), @"[0-9]+").Groups[0].ToString();
            //string 标段 = dt.Rows[0]["拍子号"].ToString() == "标段" ? dt.Rows[0]["图纸页码"].ToString() : dt.Rows[1]["图纸页码"].ToString();
            //string 拆图员 = dt.Rows[0]["拍子号"].ToString() == "标段" ? dt.Rows[0]["工艺备注"].ToString() : dt.Rows[1]["工艺备注"].ToString();
            //string 整单备注 = dt.Rows[0]["拍子号"].ToString() == "标段" ? "无" : dt.Rows[0]["拍子号"].ToString();
            //string 纹路 = dt.Rows[0]["拍子号"].ToString() == "标段" ? dt.Rows[2]["开料长"].ToString() : dt.Rows[3]["开料长"].ToString();
            if (Regex.IsMatch(纹路, @"(宽|width)+")){纹路 = "宽";}
            Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
            Npoi.ColumnSetValue(ref dt, 标段, "标段");
            Npoi.ColumnSetValue(ref dt, 套数, "套数");
            Npoi.ColumnSetValue(ref dt, 拆图员, "拆图员");
            Npoi.ColumnSetValue(ref dt, 整单备注, "整单备注");
            Npoi.ColumnSetValue(ref dt, "门板", "Class");
            Npoi.ColumnSetValue(ref dt, 纹路, "纹路");
            return dt;
        }
        public DataTable PvcDoor(DataTable dataTable)
        {
            string[] 门板汇总 = new string[] { "拍子号", "图纸页码", "门型", "部件名称", "开料长", "开料宽", "长", "宽", "厚", "预留1", "总数量", "材质", "膜皮型号", "开门方向", "工艺备注" };
            DataTable dt = Npoi.DeleteColumns(dataTable, 1, 15);
            string 项目名称 = dt.Columns[0].ColumnName;
            for (int i = 0; i < dt.Columns.Count; i++){dt.Columns[i].ColumnName = 门板汇总[i];}
            dt.Columns.Add("Class", typeof(string));
            dt.Columns.Add("项目名称", typeof(string));
            dt.Columns.Add("标段", typeof(string));
            dt.Columns.Add("套数", typeof(string));
            dt.Columns.Add("拆图员", typeof(string));
            dt.Columns.Add("整单备注", typeof(string));
            dt.Columns.Add("纹路", typeof(string));
            string 套数 = Regex.Match(dt.Rows[1]["总数量"].ToString(), @"[0-9]+").Groups[0].ToString();
            string 标段 = dt.Rows[1]["图纸页码"].ToString();
            string 拆图员 = dt.Rows[1]["工艺备注"].ToString();
            string 整单备注 = dt.Rows[0]["拍子号"].ToString();
            string 纹路 = dt.Rows[3]["开料长"].ToString();
            if (Regex.IsMatch(纹路, @"(宽|width)+")){纹路 = "宽";}
            Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
            Npoi.ColumnSetValue(ref dt, 标段, "标段");
            Npoi.ColumnSetValue(ref dt, 套数, "套数");
            Npoi.ColumnSetValue(ref dt, 拆图员, "拆图员");
            Npoi.ColumnSetValue(ref dt, 整单备注, "整单备注");
            Npoi.ColumnSetValue(ref dt, Regex.IsMatch(dataTable.TableName, @"(单面)+(吸塑)+") ? "单面吸塑" : "双面吸塑", "Class");
            Npoi.ColumnSetValue(ref dt, 纹路, "纹路");
            return dt;
        }
        public void InputFlowCardData(string[] fileNames)
        {
            StaticVariable.FlowCardDataTable = null;
            DataTable flowCardTable = new Variable().FlowCardDataTalble();
            Dictionary<string, FlowCardType> assembly = new Dictionary<string, FlowCardType>
            {
                { "柜体汇总", Cabinet },
                { "门板汇总", Door },
                { "单面吸塑门板汇总", PvcDoor },
                { "双面吸塑门板汇总", PvcDoor }
            };
            for (int fileIndex = 0; fileIndex < fileNames.Length; fileIndex++)
            {
                DataSet dsp = Npoi.ExcelToDataSet(fileNames[fileIndex]);
                foreach (DataTable dataTable in dsp.Tables)
                {
                    if (assembly.ContainsKey(dataTable.TableName))
                    {
                        using (DataTable fdt = assembly[dataTable.TableName](dataTable))
                        {
                            for (int i = fdt.Rows.Count - 1; i >= 0; i--)
                            {
                                if (fdt.Rows[i]["部件名称"] == DBNull.Value
                                    || fdt.Rows[i]["开料长"] == DBNull.Value
                                    || !Regex.IsMatch(fdt.Rows[i]["总数量"].ToString(), @"[0-9]+"))
                                    fdt.Rows.RemoveAt(i);
                            }
                            flowCardTable.Merge(fdt);
                        }
                    }
                }
            }
            if (flowCardTable.Columns.Contains("预留1")) flowCardTable.Columns.Remove("预留1");
            if (flowCardTable.Columns.Contains("预留2")) flowCardTable.Columns.Remove("预留2");
            StaticVariable.FlowCardDataTable = FlowCardGroupBy(flowCardTable);
        }
        private DataTable FlowCardGroupBy(DataTable dataTable)
        {
            DataTable groupdt = Npoi.ColumnsToDecimal(dataTable, 7, 8, 9, 17).Clone();
            foreach (DataRow dataRow in Npoi.ColumnsToDecimal(dataTable, 7, 8, 9, 17).Rows)
            {
                int depth = (int)(decimal)dataRow["厚"];
                int nums = (int)(decimal)dataRow["总数量"];
                decimal length = (decimal)dataRow["开料长"];
                decimal width = (decimal)dataRow["开料宽"];
                Panel(ref groupdt, dataRow, depth, nums, length, width);
                //if (Math.Max(length, width) > 600 && Math.Min(length, width) <= 250)//小板件1
                //{
                //}
                //else if (Math.Max(length, width) < 350)//小板件2
                //{
                //}
                //else if (Math.Max(length, width) <= 600)//小板件3
                //{
                //}
                //else//大板件1
                //{
                //}
            }
            DataRow[] drcabinet = groupdt.Select("Class = '柜体'");
            DataRow[] drdoor = groupdt.Select("Class = '门板'");
            DataRow[] drsingle = groupdt.Select("Class = '单面吸塑'");
            DataRow[] drdouble = groupdt.Select("Class = '双面吸塑'");
            Npoi.SetPageNumber(ref drcabinet, "页码");
            Npoi.SetPageNumber(ref drdoor, "页码");
            Npoi.SetPageNumber(ref drsingle, "页码");
            Npoi.SetPageNumber(ref drdouble, "页码");
            groupdt.AcceptChanges();
            return groupdt;
        }
        private void Panel(ref DataTable dataTable, DataRow dataRow, int depth, int nums, decimal length, decimal width)
        {
            int max = (int)Math.Floor(600 / Math.Min(length, width));
            int min = (int)Math.Floor(900 / Math.Max(length, width));
            if (max == 0) { max = 1; }
            if (min == 0) { min = 1; }
            int layer = max * min;
            int stack;
            if ((depth * nums / layer) % StaticVariable.StackHeigth < 161
                && depth * nums / layer % StaticVariable.StackHeigth > 0
                && depth * nums / layer > StaticVariable.StackHeigth)
            {
                stack = (int)(Math.Ceiling((decimal)depth * (decimal)nums / (decimal)StaticVariable.StackHeigth / layer) - 1);
            }
            else
            {
                stack = (int)Math.Ceiling((decimal)depth * (decimal)nums / (decimal)StaticVariable.StackHeigth / layer);
            }
            DataTable dt_temp = dataTable.Clone();
            dt_temp.ImportRow(dataRow);
            DataRow dr = dt_temp.Rows[0];
            string order = dataRow["拍子号"].ToString();
            for (int i = 1; i <= stack; i++)
            {
                dr["拍子号"] = order + "-" + stack.ToString() + "-" + i.ToString();
                if (i != stack)
                {
                    dr["数量"] = StaticVariable.StackHeigth / depth * layer;
                }
                else
                {
                    dr["数量"] = nums - StaticVariable.StackHeigth / depth * layer * (stack - 1);
                }
                dataTable.ImportRow(dr);
            }
        }
        public DataTable DrawingsPrint(DataTable dataTable)
        {
            DataTable dtemp = dataTable.AsDataView().ToTable(false, "Class", "图纸页码");
            dtemp.Columns.Add("打印份数", typeof(int));
            Npoi.ColumnSetValue(ref dtemp, 1, "打印份数");
            var query = from c in dtemp.AsEnumerable()
                        where c["图纸页码"].ToString() != ""
                        group c by new
                        {
                            cls = c["Class"],
                            page = c["图纸页码"]
                        } into s
                        select new
                        {
                            cls = s.Select(p => p["Class"]).First(),
                            page = s.Select(p => p["图纸页码"]).First(),
                            print = s.Sum(p => (int)p["打印份数"])
                        };
            DataTable dtmp = dtemp.Clone();
            query.ToList().ForEach(p => dtmp.Rows.Add(p.cls, p.page, p.print));
            dtmp.Columns["Class"].ColumnName = "类型";
            return dtmp;
        }
        public int TotalNumber(DataTable dataTable)
        {
            List<string> list = (from r in dataTable.AsEnumerable() select r.Field<string>("数量")).ToList();
            var list1 = list.Select<string, int>(x => Convert.ToInt32(x));
            return list1.Sum();
        }
    }
}
