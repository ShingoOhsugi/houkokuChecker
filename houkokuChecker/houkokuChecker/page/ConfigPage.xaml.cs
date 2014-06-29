using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;

namespace houkokuChecker
{
    /// <summary>
    /// ConfigPage.xaml の相互作用ロジック
    /// </summary>
    public partial class ConfigPage : Page
    {
        public ConfigPage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 設定ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSettei_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Content = new MainPage();
        }

        /// <summary>
        /// ファイルオープン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "フォルダを指定してください。";
            fbd.RootFolder = System.Environment.SpecialFolder.Desktop;
            fbd.SelectedPath = txtRootPath.Text;
            fbd.RootFolder = System.Environment.SpecialFolder.Desktop;

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                txtRootPath.Text = fbd.SelectedPath;
            }
        }

    }
}
