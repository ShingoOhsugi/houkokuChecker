using System.Windows;
using System.Windows.Controls;

namespace houkokuChecker
{
    /// <summary>
    /// WelcomePage.xaml の相互作用ロジック
    /// </summary>
    public partial class WelcomePage : Page
    {
        public WelcomePage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 開始ボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Content = new ConfigPage();
        }

    }
}
