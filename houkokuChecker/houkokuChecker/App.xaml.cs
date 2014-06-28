using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace houkokuChecker
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        private void StartupHandler(object sender, System.Windows.StartupEventArgs e)
        {
            SolidColorBrush btn = Elysium.AccentBrushes.Blue;
            SolidColorBrush mozi = new SolidColorBrush(Colors.White);

            Elysium.Manager.Apply(
                this,
                Elysium.Theme.Dark,
                btn,
                mozi);
        }
    }
}
