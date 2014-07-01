using System.Windows;
using System.Windows.Controls;

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

        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            txtSyainCd.Text = Properties.Settings.Default.SyainCode;
            txtSyainNm.Text = Properties.Settings.Default.SyainName;
            txtRootPath.Text = Properties.Settings.Default.RootPath;

            Properties.Settings.Default.Reset();
        }

        /// <summary>
        /// 設定ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSettei_Click(object sender, RoutedEventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txtSyainCd.Text) ||
               string.IsNullOrWhiteSpace(txtSyainNm.Text) ||
               string.IsNullOrWhiteSpace(txtRootPath.Text)) 
            {
                //入力エラー
                MessageBox.Show("すべて入力してください。");
                return;
            }

            // 設定ファイルに保存
            Properties.Settings.Default.SyainCode = txtSyainCd.Text;
            Properties.Settings.Default.SyainName = txtSyainNm.Text;
            Properties.Settings.Default.RootPath = txtRootPath.Text + "\\";
            Properties.Settings.Default.Save();

            this.NavigationService.Content = new MainPage();
        }

        /// <summary>
        /// ファイルオープン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd 
                = new System.Windows.Forms.FolderBrowserDialog();
            fbd.Description = "フォルダを指定してください。";
            fbd.RootFolder = System.Environment.SpecialFolder.Desktop;
            fbd.SelectedPath = txtRootPath.Text;
            fbd.RootFolder = System.Environment.SpecialFolder.Desktop;

            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtRootPath.Text = fbd.SelectedPath;
            }
        }

    }
}
