using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace houkokuChecker
{
    /// <summary>
    /// MainPage.xaml の相互作用ロジック
    /// </summary>
    public partial class MainPage : Page
    {
        private string _rootPath = string.Empty;

        private const string SINSEI_S = "申請書_";
        private const string SINSEI_E = ".xls";

        private const string SINSEI_FMT = "申請書_{0}.xls";
        private const string HOUKOKU_FMT = "WR0240【{0}年{1}月】{2}.xls";
        private const string AD_FMT = "{0},{1}"; //行,列

        private const int R_IDX_TIME_KIHON = 9;         //行No 基本稼働時間
        private const int R_IDX_TIME_MINASHI = 36;      //行No みなし稼働時間
        private const int R_IDX_TIME_KOZYO_T = 21;      //行No 控除(遅刻)
        private const int R_IDX_TIME_KOZYO_S = 26;      //行No 控除(早退)
        private const int R_IDX_TIME_KOZYO_G = 31;      //行No 控除(外出)
        private const int R_IDX_TIME_FUTUZAN = 40;      //行No 普通残業
        private const int R_IDX_TIME_SINZAN = 44;       //行No 深夜残業
        private const int R_IDX_TIME_SOUZAN = 48;       //行No 早朝残業
        private const int R_IDX_TIME_HOTEIZAN = 52;     //行No 法定休日稼働
        private const int R_IDX_TIME_HOTEISINZAN = 56;  //行No 法定休日深夜稼働

        private const int C_IDX_HIDUKE_START = 21;      //列No 「1日」のセル

        public MainPage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 初期表示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            _rootPath = Properties.Settings.Default.RootPath;
            
            calCheckTaisyo.SelectedDate = DateTime.Now;
            calSyukeiTaisyo.SelectedDate = DateTime.Now;

            btnTorikomi_Click(null, null);
        }

        /// <summary>
        /// 更新ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTorikomi_Click(object sender, RoutedEventArgs e)
        {
            List<string> locMember = new List<string>();
            DirectoryInfo rootDicInfo = new DirectoryInfo(_rootPath);

            foreach (FileInfo fileInfo in rootDicInfo.GetFiles())
            {
                string fileNameL = fileInfo.Name.ToLower();

                if (fileNameL.Contains(SINSEI_S))
                {
                    string dispName = fileNameL.Replace(SINSEI_S, string.Empty)
                                               .Replace(SINSEI_E, string.Empty);

                    locMember.Add(dispName);
                }
            }
            lbMember.ItemsSource = locMember;
        }

        /// <summary>
        /// ALLボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAll_Click(object sender, RoutedEventArgs e)
        {
            lbMember.SelectAll();
        }

        /// <summary>
        /// チェックボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCheck_Click(object sender, RoutedEventArgs e)
        {
            txtCheckResult.Clear();

            if (!inputCheck(btnCheck.Name))
            {
                return;
            }

            loopLogic(btnCheck.Name, calCheckTaisyo.SelectedDate.Value);

            MessageBox.Show("処理が完了しました！");

        }

        /// <summary>
        /// 集計ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSyukei_Click(object sender, RoutedEventArgs e)
        {
            dgSyukeiResult.ItemsSource = null;

            if (!inputCheck(btnSyukei.Name))
            {
                return;
            }

            loopLogic(btnSyukei.Name, calSyukeiTaisyo.SelectedDate.Value);


            MessageBox.Show("処理が完了しました！");

        }

        /// <summary>
        /// 代打ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDaida_Click(object sender, RoutedEventArgs e)
        {
            //checkLogic(btnDaida.Name);

            //loopLogic(btnDaida.Name);

            MessageBox.Show("処理が完了しました！");

        }

        /// <summary>
        /// 設定ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfig_Click(object sender, RoutedEventArgs e)
        {
            this.NavigationService.Content = new ConfigPage();
        }

        /// <summary>
        /// 入力チェック
        /// </summary>
        /// <param name="btnName"></param>
        /// <returns></returns>
        private bool inputCheck(string btnName)
        {
            if (lbMember.SelectedItems.Count == 0)
            {
                MessageBox.Show("メンバを選択してください");
                return false;
            }


            if (btnCheck.Name.Equals(btnName))
            {
                if (calCheckTaisyo.SelectedDate.HasValue == false)
                {
                    MessageBox.Show("日付を選択してください");
                    return false;
                }
            }
            else if (btnSyukei.Name.Equals(btnName))
            {
                if (calSyukeiTaisyo.SelectedDate.HasValue == false)
                {
                    MessageBox.Show("日付を選択してください");
                    return false;
                }
            }

            return true;

        }

        /// <summary>
        /// 申請、報告書取り込み
        /// </summary>
        /// <param name="btnName"></param>
        /// <returns></returns>
        private bool loopLogic(string btnName, DateTime selDate)
        {
            Microsoft.Office.Interop.Excel.Application xlsSinsei
                = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Application xlsHokoku
                = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wrkBookSinsei = null;
            Microsoft.Office.Interop.Excel.Workbook wrkBookHokoku = null;

            foreach (string selMember in lbMember.SelectedItems)
            {
                try
                {
                    //申請書Open
                    string sinseiFilePath =
                        _rootPath + string.Format(SINSEI_FMT, selMember);

                    wrkBookSinsei = xlsSinsei.Workbooks.Open(sinseiFilePath);

                    //報告書Open
                    string hokokuFilePath =
                        _rootPath + string.Format(HOUKOKU_FMT,
                                                  selDate.Year.ToString(),
                                                  selDate.Month.ToString("00"),
                                                  selMember);

                    //報告書なしエラー
                    if (File.Exists(hokokuFilePath) == false)
                    {
                        MessageBox.Show(selMember + " の報告書を格納してください。:" + hokokuFilePath);

                        return false;
                    }
                    wrkBookHokoku = xlsHokoku.Workbooks.Open(hokokuFilePath);


                    //チェック
                    if (btnCheck.Name.Equals(btnName))
                    {
                        checkMain(wrkBookSinsei, wrkBookHokoku, selDate);
                    }
                    else if (btnSyukei.Name.Equals(btnName))
                    {
                        syukeiMain(wrkBookSinsei, wrkBookHokoku, selDate);
                    }
                }
                finally
                {
                    //Close
                    if (wrkBookSinsei != null)
                    {
                        wrkBookSinsei.Close(false);
                    }
                    if (wrkBookHokoku != null)
                    {
                        wrkBookHokoku.Close(false);
                    }
                }
            }

            return true;

        }

        /// <summary>
        /// 報告書チェックメイン
        /// </summary>
        /// <param name="sinseiBook"></param>
        /// <param name="houkokuBook"></param>
        private void checkMain(
            Microsoft.Office.Interop.Excel.Workbook sinseiBook,
            Microsoft.Office.Interop.Excel.Workbook houkokuBook,
            DateTime selDate)
        {

            //申請情報取込



            //報告情報チェック

            // NGチェック
            Dictionary<string, string> dicHoukokuAlldata = getAllData(houkokuBook)[0];

            foreach (KeyValuePair<string, string> hitData in dicHoukokuAlldata.Where(p => p.Value == ("NG")))
            {
                string[] hitAd = splitAdress(hitData.Key);
                txtCheckResult.Text += string.Format("NGがあります。 行No: {0} 列No: {1} \n", hitAd[0], hitAd[1]);
            };

            // 指定期間に限定
            for (int i = 0; i < selDate.Day; i++)
            {
                // PJNo必須チェック
                // 休日必須チェック

                // 申請チェック
            }

            return;
        }

        /// <summary>
        /// 集計処理メイン
        /// </summary>
        /// <param name="sinseiBook"></param>
        /// <param name="houkokuBook"></param>
        private void syukeiMain(
            Microsoft.Office.Interop.Excel.Workbook sinseiBook,
            Microsoft.Office.Interop.Excel.Workbook houkokuBook,
            DateTime selDate)
        {
            DataTable dtKekka = new DataTable();
            dtKekka.Columns.Add("氏名", typeof(string));
            dtKekka.Columns.Add("総稼動", typeof(string));
            dtKekka.Columns.Add("基本稼働", typeof(string));
            dtKekka.Columns.Add("みなし稼働", typeof(string));
            dtKekka.Columns.Add("控除", typeof(string));
            dtKekka.Columns.Add("普通残業", typeof(string));
            dtKekka.Columns.Add("深夜残業", typeof(string));
            dtKekka.Columns.Add("早朝稼働", typeof(string));
            dtKekka.Columns.Add("法定休日稼働", typeof(string));
            dtKekka.Columns.Add("作業内容", typeof(string));

            //申請情報解析
            List<Dictionary<string, string>> dicSinseiAlldata = getAllData(sinseiBook);

            foreach (Dictionary<string, string> dicSheet in dicSinseiAlldata)
            {

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Contains("日付")))
                {
                    string[] hitAd = splitAdress(hitData.Key);


                };
            }

            //報告情報解析
            Dictionary<string, string> dicHoukokuAlldata = getAllData(houkokuBook)[0];
            TimeSpan totalKihon = new TimeSpan();

            // 指定期間に限定
            for (int i = 0; i < selDate.Day; i++)
            {

                //基本稼働時間
                string adKihon = string.Format(AD_FMT, R_IDX_TIME_KIHON, C_IDX_HIDUKE_START + i);
                double dblKihon = 0;
                if (double.TryParse(dicHoukokuAlldata[adKihon], out dblKihon))
                {
                    DateTime dtKihon = DateTime.FromOADate(dblKihon);
                    totalKihon += dtKihon.TimeOfDay;
                }
            }

            //バインド用結果設定
            DataRow drKekka = dtKekka.NewRow();
            drKekka["氏名"] = dicHoukokuAlldata["4,7"];
            drKekka["総稼動"] = string.Empty;
            drKekka["基本稼働"] = totalKihon.TotalHours.ToString("0") + totalKihon.ToString(@"\:mm");
            drKekka["みなし稼働"] = string.Empty;
            drKekka["控除"] = string.Empty;
            drKekka["普通残業"] = string.Empty;
            drKekka["深夜残業"] = string.Empty;
            drKekka["早朝稼働"] = string.Empty;
            drKekka["法定休日稼働"] = string.Empty;
            drKekka["作業内容"] = string.Empty;
            dtKekka.Rows.Add(drKekka);

            dgSyukeiResult.ItemsSource = dtKekka.DefaultView;

            return;
        }

        /// <summary>
        /// Excelデータ → Dic変換
        /// </summary>
        /// <param name="wkBook"></param>
        /// <returns></returns>
        private List<Dictionary<string, string>> getAllData(Microsoft.Office.Interop.Excel.Workbook wkBook)
        {
            List<Dictionary<string, string>> res = new List<Dictionary<string, string>>();

            foreach (Microsoft.Office.Interop.Excel.Worksheet wkSheet in wkBook.Sheets)
            {
                Dictionary<string, string> dicSheet = new Dictionary<string, string>();

                Microsoft.Office.Interop.Excel.Range aRange = wkSheet.UsedRange;

                object[,] dataAll = aRange.get_Value();

                if (dataAll == null)
                {
                    continue;
                }

                long iRowCnt = dataAll.GetUpperBound(0);
                long iColCnt = dataAll.GetUpperBound(1);

                for (long r = 1; r <= iRowCnt; r++)
                {
                    for (long c = 1; c <= iColCnt; c++)
                    {
                        dicSheet.Add(
                            createAdress(r, c),
                            dataAll[r, c] == null ? string.Empty : dataAll[r, c].ToString());
                    }
                }

                res.Add(dicSheet);
            }

            return res;
        }

        /// <summary>
        /// Dicアドレス分割
        /// </summary>
        /// <param name="dicKey"></param>
        /// <returns>[0]:行　[1]:列</returns>
        private static string[] splitAdress(string dicKey)
        {
            return dicKey.Split(',');
        }

        /// <summary>
        /// Dicアドレス作成
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static string createAdress(object row, object col)
        {
            return string.Format(AD_FMT, row, col);
        }



        //申請チェックサンプル
        private void testc()
        {

            List<string> outList = new List<string>();//記入対象日：申請内容：新記入日

            //Microsoft.Office.Interop.Excel
            List<string[]> mainList = new List<string[]>();
            List<string[]> subList = new List<string[]>();

            mainList.Add(new string[] { "2015/1/1", "2015/2/1" });

            subList.Add(new string[] { "2015/1/1", "2015/2/3" });
            subList.Add(new string[] { "2015/1/1", "2015/2/4" });

            foreach (string[] row in mainList)
            {
                outList.Add("対象" + row[0] + ":振出" + row[1]);

                string oldDate = row[1];

                foreach (string[] subRow in subList)
                {
                    if (row[0].Equals(subRow[0]))
                    {
                        outList.Add("対象" + oldDate + ":振休訂" + subRow[1]);
                        oldDate = subRow[1];
                    }
                }
                outList.Add("対象" + oldDate + ":振休" + row[0]);

            }
            outList.Sort();

            MessageBox.Show("ユーザ情報セット");

        }

    }
}
