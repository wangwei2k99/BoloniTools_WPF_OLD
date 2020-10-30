using System.Data;
using BoloniTools.Controller;
namespace BoloniTools
{
    public class Variable
    {
        public DataTable ExportMesDataTable()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("物件名称", typeof(string));
            dataTable.Columns.Add("柜体类别", typeof(string));
            dataTable.Columns.Add("柜体编号", typeof(string));
            dataTable.Columns.Add("柜体名称", typeof(string));
            dataTable.Columns.Add("类别", typeof(string));
            dataTable.Columns.Add("单位", typeof(string));
            dataTable.Columns.Add("编码", typeof(string));
            dataTable.Columns.Add("基材编码", typeof(string));
            dataTable.Columns.Add("基材", typeof(string));
            dataTable.Columns.Add("颜色编码", typeof(string));
            dataTable.Columns.Add("颜色", typeof(string));
            dataTable.Columns.Add("表面处理", typeof(string));
            dataTable.Columns.Add("门板型号", typeof(string));
            dataTable.Columns.Add("开门方向", typeof(string));
            dataTable.Columns.Add("数量", typeof(string));
            dataTable.Columns.Add("成品长", typeof(string));
            dataTable.Columns.Add("成品宽", typeof(string));
            dataTable.Columns.Add("厚", typeof(string));
            dataTable.Columns.Add("铣型", typeof(string));
            dataTable.Columns.Add("开槽", typeof(string));
            dataTable.Columns.Add("打孔", typeof(string));
            dataTable.Columns.Add("加工信息汇总", typeof(string));
            dataTable.Columns.Add("封边", typeof(string));
            dataTable.Columns.Add("封边前厚度", typeof(string));
            dataTable.Columns.Add("封边后厚度", typeof(string));
            dataTable.Columns.Add("封边左厚度", typeof(string));
            dataTable.Columns.Add("封边右厚度", typeof(string));
            dataTable.Columns.Add("封边前颜色编码", typeof(string));
            dataTable.Columns.Add("封边后颜色编码", typeof(string));
            dataTable.Columns.Add("封边左颜色编码", typeof(string));
            dataTable.Columns.Add("封边右颜色编码", typeof(string));
            dataTable.Columns.Add("打孔图纸号", typeof(string));
            dataTable.Columns.Add("反面打孔图纸号", typeof(string));
            dataTable.Columns.Add("长", typeof(string));
            dataTable.Columns.Add("宽", typeof(string));
            dataTable.Columns.Add("备注", typeof(string));
            dataTable.Columns.Add("翻板", typeof(string));
            dataTable.Columns.Add("纹路", typeof(string));
            dataTable.Columns.Add("规格", typeof(string));
            dataTable.Columns.Add("门板系列", typeof(string));
            dataTable.Columns.Add("拉手类型", typeof(string));
            dataTable.Columns.Add("型材封边", typeof(string));
            dataTable.Columns.Add("生产加工数据", typeof(string));
            dataTable.Columns.Add("前预铣值", typeof(string));
            dataTable.Columns.Add("后预铣值", typeof(string));
            dataTable.Columns.Add("左预铣值", typeof(string));
            dataTable.Columns.Add("右预铣值", typeof(string));
            dataTable.Columns.Add("前铣型封边判断", typeof(string));
            dataTable.Columns.Add("后铣型封边判断", typeof(string));
            dataTable.Columns.Add("左铣型封边判断", typeof(string));
            dataTable.Columns.Add("右铣型封边判断", typeof(string));
            dataTable.Columns.Add("采购编码", typeof(string));
            return dataTable;
        }
        public DataTable FlowCardDataTalble()
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("项目名称", typeof(string));
            dataTable.Columns.Add("标段", typeof(string));
            dataTable.Columns.Add("套数", typeof(string));
            dataTable.Columns.Add("拆图员", typeof(string));
            dataTable.Columns.Add("拍子号", typeof(string));
            dataTable.Columns.Add("工艺图纸", typeof(string));
            dataTable.Columns.Add("部件名称", typeof(string));
            dataTable.Columns.Add("开料长", typeof(string));
            dataTable.Columns.Add("开料宽", typeof(string));
            dataTable.Columns.Add("厚", typeof(string));
            dataTable.Columns.Add("数量", typeof(string));
            dataTable.Columns.Add("材质", typeof(string));
            dataTable.Columns.Add("封边描述", typeof(string));
            dataTable.Columns.Add("膜皮型号", typeof(string));
            dataTable.Columns.Add("工艺备注", typeof(string));
            dataTable.Columns.Add("图纸页码", typeof(string));
            dataTable.Columns.Add("整单备注", typeof(string));
            dataTable.Columns.Add("总数量", typeof(string));
            dataTable.Columns.Add("纹路", typeof(string));
            dataTable.Columns.Add("开门方向", typeof(string));
            dataTable.Columns.Add("门型", typeof(string));
            dataTable.Columns.Add("页码", typeof(string));
            dataTable.Columns.Add("Class", typeof(string));
            return dataTable;
        }
    }
    public static class StaticVariable
    {
        private static DataTable mesDataTable;

        public static DataTable MesDataTable
        {
            get
            {
                if (mesDataTable == null || mesDataTable.Rows.Count == 0)
                {
                    Notice.NoticeFunc("您选择的不是拆图文件、或格式有误！");
                    return null;
                }
                else
                {
                    return mesDataTable;
                }
            }
            set { mesDataTable = value; }
        }

        private static DataTable flowCardDataTable;

        public static DataTable FlowCardDataTable
        {
            get {
                if (flowCardDataTable == null || flowCardDataTable.Rows.Count == 0)
                {
                    Notice.NoticeFunc("您选择的不是汇总文件、或格式有误！");
                    return null;
                }
                else
                {
                    return flowCardDataTable;
                }
            }
            set { flowCardDataTable = value; }
        }
        public static int StackHeigth { get { return stackHeigth; } }
        private static readonly int stackHeigth = 960;
    }
}
