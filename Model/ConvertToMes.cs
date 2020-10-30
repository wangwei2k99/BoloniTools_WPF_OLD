using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
namespace BoloniTools.Func
{
    class ConvertToMes
    {
        public static DataTable ConvertMes(string[] Files)
        {
            var dts = new Variable().ExportMesDataTable();
            string[] 柜体 = new string[] { "柜体名称", "柜号", "柜宽", "柜深", "柜高", "柜数量", "柜体编号", "加工信息汇总", "工艺图纸", "物件名称", "长", "宽", "厚", "数量", "总数量", "基材", "封边", "铣型", "包装箱号" };
            string[] 铝材 = new string[] { "柜体编号", "成品长", "成品宽", "类别", "预留1", "采购编码", "物件名称", "规格", "长", "宽", "数量", "总数量", "单位", "铣型" };
            string[] 五金 = new string[] { "类别", "采购编码", "物件名称", "规格", "数量", "总数量", "单位", "铣型" };
            string[] 吸塑门板 = new string[] { "柜体编号", "加工信息汇总", "编码", "物件名称", "长", "宽", "成品长", "成品宽", "厚", "数量", "总数量", "基材", "膜皮型号", "开门方向", "铣型" };
            string[] 耐磨板门板 = new string[] { "柜体编号", "加工信息汇总", "编码", "物件名称", "长", "宽", "成品长", "成品宽", "厚", "数量", "总数量", "基材", "封边", "开门方向", "铣型" };
            string[] 工作表 = { "柜体", "铝材", "五金", "耐磨板门板", "单面吸塑门板", "双面吸塑门板" };
            for (int FileIndex = 0; FileIndex < Files.Length; FileIndex++)
            {
                string 项目编码;
                string 项目名称;
                string 订单编号;
                string 总套数;
                //MemoryStream ms = new MemoryStream(File.ReadAllBytes(Files[FileIndex]));
                DataSet dsp = Npoi.ExcelToDataSet(Files[FileIndex]);
                for (int Sht = 0; Sht < 工作表.Length; Sht++)
                {
                    foreach (DataTable dt_temp in dsp.Tables)
                    {
                        switch (Sht)
                        {
                            case 0:
                                {
                                    if (dt_temp.TableName.Trim() != 工作表[Sht]) break;
                                    DataTable dt = Npoi.DeleteColumns(dt_temp, 1, 19);
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    { dt.Columns[i].ColumnName = 柜体[i]; }
                                    项目编码 = dt.Rows[4][9].ToString();
                                    项目名称 = dt.Rows[4][12].ToString();
                                    订单编号 = dt.Rows[3][12].ToString();
                                    总套数 = dt.Rows[5][9].ToString();
                                    dt.Columns.Add("项目编码", typeof(string));
                                    dt.Columns.Add("项目名称", typeof(string));
                                    dt.Columns.Add("订单编号", typeof(string));
                                    dt.Columns.Add("总套数", typeof(string));
                                    dt.Columns.Add("类别", typeof(string));
                                    dt.Columns.Add("单位", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, 项目编码, "项目编码");
                                    Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
                                    Npoi.ColumnSetValue(ref dt, 订单编号, "订单编号");
                                    Npoi.ColumnSetValue(ref dt, 总套数, "总套数");
                                    Npoi.ColumnSetValue(ref dt, "板件", "类别");
                                    Npoi.ColumnSetValue(ref dt, "块", "单位");
                                    BomVlookup(ref dt, "工艺图纸", "物件名称", "基材");
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (dt.Rows[i]["物件名称"] == DBNull.Value
                                            || dt.Rows[i]["长"] == DBNull.Value
                                            || dt.Rows[i]["总数量"] == DBNull.Value
                                            || dt.Rows[i]["总数量"].ToString() == "总数")
                                        { dt.Rows[i].Delete(); }
                                    }
                                    dt.AcceptChanges();
                                    for (int i = 0; i < dt.Rows.Count; i++)
                                    {
                                        if (dt.Rows[i]["柜体名称"] == DBNull.Value)
                                        {
                                            if (dt.Rows[i]["物件名称"].ToString().Contains("辅助板"))
                                            {
                                                dt.Rows[i]["柜体名称"] = "辅助板";
                                            }
                                            else if (i >= 1)
                                            {
                                                dt.Rows[i]["柜体名称"] = dt.Rows[i - 1]["柜体名称"];
                                            }
                                        }
                                    }
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (i == 0 || dt.Rows[i]["柜号"].ToString() != "" && !dt.Rows[i]["柜体名称"].ToString().Contains("板"))
                                        {
                                            string 柜体名称, 柜号, 柜宽, 柜深, 柜高, 柜数量;
                                            柜体名称 = dt.Rows[i]["柜体名称"].ToString();
                                            柜号 = dt.Rows[i]["柜号"].ToString();
                                            柜宽 = dt.Rows[i]["柜宽"].ToString();
                                            柜深 = dt.Rows[i]["柜深"].ToString();
                                            柜高 = dt.Rows[i]["柜高"].ToString();
                                            柜数量 = dt.Rows[i]["柜数量"].ToString();
                                            DataRow drnew = dt.NewRow();
                                            drnew["柜体名称"] = 柜体名称;
                                            drnew["物件名称"] = 柜体名称;
                                            drnew["柜体编号"] = 柜号;
                                            drnew["类别"] = "柜体";
                                            drnew["单位"] = "个";
                                            drnew["厚"] = 柜高;
                                            drnew["长"] = 柜宽;
                                            drnew["宽"] = 柜深;
                                            drnew["数量"] = 柜数量;
                                            drnew["采购编码"] = "GT";
                                            drnew["订单编号"] = 订单编号;
                                            dt.Rows.InsertAt(drnew, i);
                                        }
                                    }
                                    dts.Merge(dt);
                                    break;
                                }
                            case 1:
                                {
                                    if (dt_temp.TableName.Trim() != 工作表[Sht]) break;
                                    DataTable dt = Npoi.DeleteColumns(dt_temp, 1, 14);
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    { dt.Columns[i].ColumnName = 铝材[i]; }
                                    项目编码 = dt.Rows[3][6].ToString();
                                    项目名称 = dt.Rows[3][8].ToString();
                                    订单编号 = dt.Rows[2][8].ToString();
                                    总套数 = dt.Rows[4][6].ToString();
                                    dt.Columns.Add("项目编码", typeof(string));
                                    dt.Columns.Add("项目名称", typeof(string));
                                    dt.Columns.Add("订单编号", typeof(string));
                                    dt.Columns.Add("总套数", typeof(string));
                                    dt.Columns.Add("柜体名称", typeof(string));
                                    dt.Columns.Add("基材编码", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, 项目编码, "项目编码");
                                    Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
                                    Npoi.ColumnSetValue(ref dt, 订单编号, "订单编号");
                                    Npoi.ColumnSetValue(ref dt, 总套数, "总套数");
                                    Npoi.ColumnSetValue(ref dt, "五金", "类别");
                                    Npoi.ColumnSetValue(ref dt, "铝材", "柜体名称");
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (dt.Rows[i]["采购编码"] == DBNull.Value
                                            || dt.Rows[i]["长"] == DBNull.Value
                                            || dt.Rows[i]["总数量"] == DBNull.Value
                                            || dt.Rows[i]["总数量"].ToString() == "总数")
                                        { dt.Rows[i].Delete(); }
                                        else
                                        {
                                            dt.Rows[i]["物件名称"] = dt.Rows[i]["物件名称"].ToString() + "(铝材)";
                                            dt.Rows[i]["基材编码"] = dt.Rows[i]["采购编码"];
                                        }
                                    }
                                    dt.AcceptChanges();
                                    dts.Merge(dt);
                                    break;
                                }
                            case 2:
                                {
                                    if (dt_temp.TableName.Trim() != 工作表[Sht]) break;
                                    DataTable dt = Npoi.DeleteColumns(dt_temp, 1, 8);
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    { dt.Columns[i].ColumnName = 五金[i]; }
                                    项目编码 = dt.Rows[2][4].ToString();
                                    项目名称 = dt.Rows[2][2].ToString();
                                    订单编号 = dt.Rows[1][2].ToString();
                                    总套数 = dt.Rows[3][4].ToString();
                                    dt.Columns.Add("项目编码", typeof(string));
                                    dt.Columns.Add("项目名称", typeof(string));
                                    dt.Columns.Add("订单编号", typeof(string));
                                    dt.Columns.Add("总套数", typeof(string));
                                    dt.Columns.Add("柜体名称", typeof(string));
                                    dt.Columns.Add("基材编码", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, 项目编码, "项目编码");
                                    Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
                                    Npoi.ColumnSetValue(ref dt, 订单编号, "订单编号");
                                    Npoi.ColumnSetValue(ref dt, 总套数, "总套数");
                                    Npoi.ColumnSetValue(ref dt, "五金", "类别");
                                    Npoi.ColumnSetValue(ref dt, "五金", "柜体名称");
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (dt.Rows[i]["采购编码"] == DBNull.Value
                                            || dt.Rows[i]["物件名称"] == DBNull.Value
                                            || dt.Rows[i]["总数量"] == DBNull.Value
                                            || dt.Rows[i]["总数量"].ToString() == "总数量")
                                        { dt.Rows[i].Delete(); }
                                        else
                                        { dt.Rows[i]["基材编码"] = dt.Rows[i]["采购编码"]; }
                                    }
                                    dt.AcceptChanges();
                                    dts.Merge(dt);
                                    break;
                                }
                            case 3:
                                {
                                    if (dt_temp.TableName.Trim() != 工作表[Sht]) break;
                                    DataTable dt = Npoi.DeleteColumns(dt_temp, 1, 15);
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    { dt.Columns[i].ColumnName = 耐磨板门板[i]; }
                                    项目编码 = dt.Rows[3][3].ToString();
                                    项目名称 = dt.Rows[3][6].ToString();
                                    订单编号 = dt.Rows[2][6].ToString();
                                    总套数 = dt.Rows[4][3].ToString();
                                    dt.Columns.Add("项目编码", typeof(string));
                                    dt.Columns.Add("项目名称", typeof(string));
                                    dt.Columns.Add("订单编号", typeof(string));
                                    dt.Columns.Add("总套数", typeof(string));
                                    dt.Columns.Add("类别", typeof(string));
                                    dt.Columns.Add("柜体名称", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, 项目编码, "项目编码");
                                    Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
                                    Npoi.ColumnSetValue(ref dt, 订单编号, "订单编号");
                                    Npoi.ColumnSetValue(ref dt, 总套数, "总套数");
                                    Npoi.ColumnSetValue(ref dt, "耐磨板门板", "类别");
                                    Npoi.ColumnSetValue(ref dt, "耐磨板门板", "柜体名称");
                                    dt.Columns.Add("单位", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, "块", "单位");
                                    BomVlookup(ref dt, "编码", "物件名称", "基材");
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (dt.Rows[i]["物件名称"] == DBNull.Value
                                            || dt.Rows[i]["长"] == DBNull.Value
                                            || dt.Rows[i]["总数量"] == DBNull.Value
                                            || dt.Rows[i]["总数量"].ToString() == "总数")
                                            dt.Rows[i].Delete();
                                    }
                                    dt.AcceptChanges();
                                    dts.Merge(dt);
                                    break;
                                }
                            case 4:
                                {
                                    if (dt_temp.TableName.Trim() != 工作表[Sht]) break;
                                    DataTable dt = Npoi.DeleteColumns(dt_temp, 1, 15);
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    { dt.Columns[i].ColumnName = 吸塑门板[i]; }
                                    项目编码 = dt.Rows[3][3].ToString();
                                    项目名称 = dt.Rows[3][6].ToString();
                                    订单编号 = dt.Rows[2][6].ToString();
                                    总套数 = dt.Rows[4][3].ToString();
                                    dt.Columns.Add("项目编码", typeof(string));
                                    dt.Columns.Add("项目名称", typeof(string));
                                    dt.Columns.Add("订单编号", typeof(string));
                                    dt.Columns.Add("总套数", typeof(string));
                                    dt.Columns.Add("类别", typeof(string));
                                    dt.Columns.Add("柜体名称", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, 项目编码, "项目编码");
                                    Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
                                    Npoi.ColumnSetValue(ref dt, 订单编号, "订单编号");
                                    Npoi.ColumnSetValue(ref dt, 总套数, "总套数");
                                    Npoi.ColumnSetValue(ref dt, "单面吸塑门板", "类别");
                                    Npoi.ColumnSetValue(ref dt, "单面吸塑门板", "柜体名称");
                                    dt.Columns.Add("单位", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, "块", "单位");
                                    BomVlookup(ref dt, "编码", "物件名称", "基材");
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (dt.Rows[i]["物件名称"] == DBNull.Value
                                            || dt.Rows[i]["长"] == DBNull.Value
                                            || dt.Rows[i]["总数量"] == DBNull.Value
                                            || dt.Rows[i]["总数量"].ToString() == "总数")
                                            dt.Rows[i].Delete();
                                    }
                                    dt.AcceptChanges();
                                    dts.Merge(dt);
                                    break;
                                }
                            case 5:
                                {
                                    if (dt_temp.TableName.Trim() != 工作表[Sht]) break;
                                    DataTable dt = Npoi.DeleteColumns(dt_temp, 1, 15);
                                    for (int i = 0; i < dt.Columns.Count; i++)
                                    { dt.Columns[i].ColumnName = 吸塑门板[i]; }
                                    项目编码 = dt.Rows[3][3].ToString();
                                    项目名称 = dt.Rows[3][6].ToString();
                                    订单编号 = dt.Rows[2][6].ToString();
                                    总套数 = dt.Rows[4][3].ToString();
                                    dt.Columns.Add("项目编码", typeof(string));
                                    dt.Columns.Add("项目名称", typeof(string));
                                    dt.Columns.Add("订单编号", typeof(string));
                                    dt.Columns.Add("总套数", typeof(string));
                                    dt.Columns.Add("类别", typeof(string));
                                    dt.Columns.Add("柜体名称", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, 项目编码, "项目编码");
                                    Npoi.ColumnSetValue(ref dt, 项目名称, "项目名称");
                                    Npoi.ColumnSetValue(ref dt, 订单编号, "订单编号");
                                    Npoi.ColumnSetValue(ref dt, 总套数, "总套数");
                                    Npoi.ColumnSetValue(ref dt, "双面吸塑门板", "类别");
                                    Npoi.ColumnSetValue(ref dt, "双面吸塑门板", "柜体名称");
                                    dt.Columns.Add("单位", typeof(string));
                                    Npoi.ColumnSetValue(ref dt, "块", "单位");
                                    BomVlookup(ref dt, "编码", "物件名称", "基材");
                                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                                    {
                                        if (dt.Rows[i]["物件名称"] == DBNull.Value
                                            || dt.Rows[i]["长"] == DBNull.Value
                                            || dt.Rows[i]["总数量"] == DBNull.Value
                                            || dt.Rows[i]["总数量"].ToString() == "总数")
                                            dt.Rows[i].Delete();
                                    }
                                    dt.AcceptChanges();
                                    dts.Merge(dt);
                                    break;
                                }
                            default:
                                break;
                        }
                    }
                }
            }
            Npoi.ColumnSetValue(ref dts, "柜体", "柜体类别");
            foreach (DataRow row in dts.Rows)
            {
                string str = row["封边"].ToString().Trim();
                if (Regex.IsMatch(str, @"^四.*[0-9.]+$"))
                {
                    string ef = Regex.Match(str, @"[$0-9.]+").Groups[0].ToString();
                    row["封边"] = string.Format("{0},{1},{2},{3}", ef, ef, ef, ef);
                    row["封边前厚度"] = ef;
                    row["封边后厚度"] = ef;
                    row["封边左厚度"] = ef;
                    row["封边右厚度"] = ef;
                }
                else if (Regex.Matches(str, @"[0-9.]+").Count == 2)
                {
                    MatchCollection ef = Regex.Matches(str, "(?<=[^0-9])[0-9.]+");
                    row["封边"] = string.Format("{0},{1},{2},{3}", ef[0], ef[1], ef[1], ef[1]);
                    row["封边前厚度"] = ef[0];
                    row["封边后厚度"] = ef[1];
                    row["封边左厚度"] = ef[1];
                    row["封边右厚度"] = ef[1];
                }
                string[] field = { "柜体编号", "成品长", "成品宽", "厚", "封边前厚度", "封边后厚度", "封边左厚度", "封边右厚度", "长", "宽", "前预铣值", "后预铣值", "左预铣值", "右预铣值" };
                foreach (string s in field) { if (row[s] == DBNull.Value) { row[s] = "0"; } }
            }
            foreach (DataRow row in dts.Rows)
            { row["规格"] = $"{row["长"]}*{row["宽"]}*{row["厚"]}"; }
            return dts;
        }
        public static void BomVlookup(ref DataTable dt, string BomCol1, string BomCol2, string Material)
        {
            if (!dt.Columns.Contains("采购编码"))
                dt.Columns.Add("采购编码", typeof(string));
            int bom_start = 0;
            Dictionary<string, string> Bom_dic = new Dictionary<string, string>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][BomCol1].ToString().Contains("物料编码"))
                    bom_start = i;
                if (i > bom_start && bom_start != 0 && dt.Rows[i][BomCol2].ToString() != "" && dt.Rows[i][BomCol1].ToString() != "")
                    Bom_dic.Add(dt.Rows[i][BomCol2].ToString(), dt.Rows[i][BomCol1].ToString());
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (Bom_dic.ContainsKey(dt.Rows[i][Material].ToString()))
                {
                    dt.Rows[i]["采购编码"] = Bom_dic[dt.Rows[i][Material].ToString()].PadLeft(9, '0');
                }
            }
        }
    }
}
