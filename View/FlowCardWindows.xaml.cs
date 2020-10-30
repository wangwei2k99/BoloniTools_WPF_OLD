using System.Windows;

namespace BoloniTools.View
{
    /// <summary>
    /// ReportViewWindows.xaml 的交互逻辑
    /// </summary>
    public partial class FlowCardWindows : Window
    {
        public FlowCardWindows()
        {
            InitializeComponent();
            this.Loaded += ReportViewWindows_Loaded;
        }

        private void ReportViewWindows_Loaded(object sender, RoutedEventArgs e)
        {
        }
    }
}
