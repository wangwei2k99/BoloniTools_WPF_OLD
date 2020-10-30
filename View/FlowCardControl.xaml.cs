using Microsoft.Reporting.WinForms;
using System.Drawing.Printing;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace BoloniTools.View
{
    /// <summary>
    /// ReportViewControl.xaml 的交互逻辑
    /// </summary>
    public partial class FlowCardControl : UserControl
    {
        public FlowCardControl()
        {
            InitializeComponent();
            this.flowcardcontrol.SetDisplayMode(DisplayMode.PrintLayout);
            this.Loaded += ReportViewControl_Loaded;
        }
        private void ReportViewControl_Loaded(object sender, RoutedEventArgs e)
        {
            ReportDataSource rds = new ReportDataSource
            {
                Name = "DataSetFlowCard",
                Value = StaticVariable.FlowCardDataTable
            };
            //PageSettings pageSettings = new PageSettings()
            //{
            //    Margins = new Margins(20, 20, 20, 24),
            //    PaperSize = new PaperSize("A5L", 827, 583),
            //    PaperSource = GetPaperSource("纸盘 1"),
            //    Landscape = false
            //};
            //flowcardcontrol.SetPageSettings(pageSettings);
            flowcardcontrol.LocalReport.ReportPath = Directory.GetCurrentDirectory() + @"\View\FlowCard.rdlc";
            flowcardcontrol.LocalReport.DataSources.Add(rds);
            flowcardcontrol.RefreshReport();
        }
        private PaperSource GetPaperSource(string sorceName)
        {
            PaperSource pageSorce = new PaperSource();
            PrinterSettings ps = new PrinterSettings();
            for (int i = 0; i < ps.PaperSources.Count; i++)
            {
                if (ps.PaperSources[i].SourceName == sorceName)
                {
                    return ps.PaperSources[i];
                }
            }
            return null;
        }
    }
}
