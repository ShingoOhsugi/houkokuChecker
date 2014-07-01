using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

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

        //申請書
        private const long C_IDX_FF_KBN1 = 3;      //列No 振出・振休 区分1
        private const long C_IDX_FF_STK1 = 4;      //列No 振出・振休 取得日1
        private const long C_IDX_FF_KBN2 = 5;      //列No 振出・振休 区分2
        private const long C_IDX_FF_STK2 = 6;      //列No 振出・振休 取得日2

        //報告書
        private const long R_IDX_TIME_KIHON = 15;        //行No 基本稼働時間
        private const long R_IDX_TIME_MINASHI_H = 14;    //行No みなし補足時間
        private const long R_IDX_TIME_MINASHI = 36;      //行No みなし稼働時間
        private const long R_IDX_TIME_KOZYO_T = 21;      //行No 控除(遅刻)
        private const long R_IDX_TIME_KOZYO_S = 26;      //行No 控除(早退)
        private const long R_IDX_TIME_KOZYO_G = 31;      //行No 控除(外出)
        private const long R_IDX_TIME_FUTUZAN = 40;      //行No 普通残業
        private const long R_IDX_TIME_SINZAN = 44;       //行No 深夜残業
        private const long R_IDX_TIME_SOUZAN = 48;       //行No 早朝残業
        private const long R_IDX_TIME_HOTEIZAN = 52;     //行No 法定休日稼働
        private const long R_IDX_TIME_HOTEISINZAN = 56;  //行No 法定休日深夜稼働

        private const long C_IDX_HIDUKE_START = 21;      //列No 「1日」のセル

        SyukeiTable _dtKekka;
        SinseiTable _dtSinsei;

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

            if (lbMember.SelectedItems.Count == 0)
            {
                MessageBox.Show("メンバを選択してください");
                return;
            }

            if (calCheckTaisyo.SelectedDate.HasValue == false)
            {
                MessageBox.Show("日付を選択してください");
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
            _dtKekka = new SyukeiTable();

            if (lbMember.SelectedItems.Count == 0)
            {
                MessageBox.Show("メンバを選択してください");
                return;
            }

            if (calSyukeiTaisyo.SelectedDate.HasValue == false)
            {
                MessageBox.Show("日付を選択してください");
                return;
            }

            loopLogic(btnSyukei.Name, calSyukeiTaisyo.SelectedDate.Value);

            dgSyukeiResult.ItemsSource = _dtKekka.DefaultView;

            MessageBox.Show("処理が完了しました！");

        }

        /// <summary>
        /// 申請確認ボタン押下時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSinseiKaku_Click(object sender, RoutedEventArgs e)
        {
            dgSinseiResult.ItemsSource = null;
            _dtSinsei = new SinseiTable();

            if (lbMember.SelectedItems.Count == 0)
            {
                MessageBox.Show("メンバを選択してください");
                return;
            }

            loopLogic(btnSinseiKaku.Name);

            dgSinseiResult.ItemsSource = _dtSinsei.DefaultView;

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
        /// 申請、報告書取り込み
        /// </summary>
        /// <param name="btnName"></param>
        /// <returns></returns>
        private bool loopLogic(string btnName, DateTime selDate = new DateTime())
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

                        //return false;
                    }
                    else
                    {
                        wrkBookHokoku = xlsHokoku.Workbooks.Open(hokokuFilePath);
                    }


                    //チェック
                    if (btnCheck.Name.Equals(btnName))
                    {
                        checkMain(wrkBookSinsei, wrkBookHokoku, selDate);
                    }
                    else if (btnSyukei.Name.Equals(btnName))
                    {
                        syukeiMain(wrkBookSinsei, wrkBookHokoku, selDate);
                    }
                    else if (btnSinseiKaku.Name.Equals(btnName))
                    {
                        sinseiKakuninMain(wrkBookSinsei);
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
                long[] hitAd = splitAdress(hitData.Key);
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

            //報告情報解析
            Dictionary<string, string> dicHoukokuAlldata = getAllData(houkokuBook)[0];
            TimeSpan totalKihon = new TimeSpan();
            TimeSpan totalMinasi = new TimeSpan();
            TimeSpan totalKojo = new TimeSpan();
            TimeSpan totalFTZan = new TimeSpan();
            TimeSpan totalSNZan = new TimeSpan();
            TimeSpan totalSOZan = new TimeSpan();
            TimeSpan totalKYZan = new TimeSpan();

            // 指定期間に限定
            for (int i = 0; i < selDate.Day; i++)
            {
                //基本稼働時間
                double dblVal = 0;
                string exAd = createAdress(R_IDX_TIME_KIHON, C_IDX_HIDUKE_START + i);

                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalKihon += dtVal.TimeOfDay;
                }

                //みなし稼働
                // みなし補足時間
                exAd = createAdress(R_IDX_TIME_MINASHI_H, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalMinasi += dtVal.TimeOfDay;
                }
                // みなし稼働
                exAd = createAdress(R_IDX_TIME_MINASHI, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalMinasi += dtVal.TimeOfDay;
                }

                //控除
                // 控除(遅刻)
                exAd = createAdress(R_IDX_TIME_KOZYO_T, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalKojo += dtVal.TimeOfDay;
                }
                // 控除(早退)
                exAd = createAdress(R_IDX_TIME_KOZYO_S, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalKojo += dtVal.TimeOfDay;
                }
                // 控除(外出)
                exAd = createAdress(R_IDX_TIME_KOZYO_G, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalKojo += dtVal.TimeOfDay;
                }

                //普通残業
                exAd = createAdress(R_IDX_TIME_FUTUZAN, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalFTZan += dtVal.TimeOfDay;
                }

                //深夜残業
                exAd = createAdress(R_IDX_TIME_SINZAN, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalSNZan += dtVal.TimeOfDay;
                }

                //早朝稼働
                exAd = createAdress(R_IDX_TIME_SOUZAN, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalSOZan += dtVal.TimeOfDay;
                }

                //法定休日稼働
                // 法定休日稼働
                exAd = createAdress(R_IDX_TIME_HOTEIZAN, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalKYZan += dtVal.TimeOfDay;
                }
                // 法定休日深夜稼働
                exAd = createAdress(R_IDX_TIME_HOTEISINZAN, C_IDX_HIDUKE_START + i);
                if (double.TryParse(dicHoukokuAlldata[exAd], out dblVal))
                {
                    DateTime dtVal = DateTime.FromOADate(dblVal);
                    totalKYZan += dtVal.TimeOfDay;
                }

                //作業内容

            }

            //バインド用結果設定
            //Math.Floor(totalKihon.TotalHours).ToString("0") + totalKihon.ToString(@"\:mm");
            DataRow drKekka = _dtKekka.NewRow();
            drKekka["氏名"] = dicHoukokuAlldata["4,7"];
            TimeSpan tsTotal = totalKihon + totalMinasi + totalKojo + totalFTZan + totalSNZan + totalSOZan + totalKYZan;
            drKekka["総稼動"] = tsTotal.TotalHours.ToString("0.00");
            drKekka["基本稼働"] = totalKihon.TotalHours.ToString("0.00");
            drKekka["みなし稼働"] = totalMinasi.TotalHours.ToString("0.00");
            drKekka["控除"] = totalKojo.TotalHours.ToString("0.00");
            drKekka["普通残業"] = totalFTZan.TotalHours.ToString("0.00");
            drKekka["深夜残業"] = totalSNZan.TotalHours.ToString("0.00");
            drKekka["早朝稼働"] = totalSOZan.TotalHours.ToString("0.00");
            drKekka["法定休日稼働"] = totalKYZan.TotalHours.ToString("0.00");
            drKekka["作業内容"] = dicHoukokuAlldata["80,2"] + " : " + dicHoukokuAlldata["81,6"];
            _dtKekka.Rows.Add(drKekka);

            return;
        }

        /// <summary>
        /// 申請確認処理メイン
        /// </summary>
        /// <param name="sinseiBook"></param>
        /// <param name="houkokuBook"></param>
        private void sinseiKakuninMain(
            Microsoft.Office.Interop.Excel.Workbook sinseiBook)
        {

            //申請情報解析
            List<Dictionary<string, string>> dicSinseiAlldata = getAllData(sinseiBook);
            //List<string[]> ffDic = new List<string[]>();  // 振出・振休申請まとめ

            foreach (Dictionary<string, string> dicSheet in dicSinseiAlldata)
            {
                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Contains("振出・振休")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);

                    for (int i = 2; i < 100; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //回数が空だったら次のセクションへ
                        if (dicSheet.ContainsKey(createAdress(rowBase, 1)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, 1)]))
                        {
                            break;
                        }

                        //区分１が空だったら次の行へ
                        string ffKbn1 = dicSheet[createAdress(rowBase, C_IDX_FF_KBN1)];
                        string ffStk1 = dicSheet[createAdress(rowBase, C_IDX_FF_STK1)];
                        string ffKbn2 = dicSheet[createAdress(rowBase, C_IDX_FF_KBN2)];
                        string ffStk2 = dicSheet[createAdress(rowBase, C_IDX_FF_STK2)];
                        if (string.IsNullOrWhiteSpace(ffKbn1))
                        {
                            continue;
                        }

                        switch (ffKbn1)
                        {
                            case "振出":
                            case "振休":
                                if ("振休".Equals(ffKbn2) || "振出".Equals(ffKbn2))
                                {
                                    //ffDic.Add(new string[] { ffStk1, ffKbn1, ffStk2 });//1/2 振出　 3/1
                                    //ffDic.Add(new string[] { ffStk2, ffKbn2, ffStk1 });//3/1 振休　 1/2
                                    DataRow dr = _dtSinsei.NewRow();
                                    dr["氏名"] = dicSheet["3,8"];
                                    dr["対象日"] = ffStk1;
                                    dr["出休"] = ffKbn1;
                                    dr["出休日付"] = ffStk2;
                                    _dtSinsei.Rows.Add(dr);

                                    dr = _dtSinsei.NewRow();
                                    dr["氏名"] = dicSheet["3,8"];
                                    dr["対象日"] = ffStk2;
                                    dr["出休"] = ffKbn2;
                                    dr["出休日付"] = ffStk1;
                                    _dtSinsei.Rows.Add(dr);

                                }
                                else if ("振休訂正".Equals(ffKbn2) || "振出訂正".Equals(ffKbn2))
                                {
                                    for (int p = 0; p < _dtSinsei.Rows.Count; p++)
                                    {
                                        DataRow rowVal = _dtSinsei.Rows[p];
                                        if (ffStk1.Equals(rowVal["出休日付"]))
                                        {
                                            DataRow dr = _dtSinsei.NewRow();
                                            dr["氏名"] = rowVal["氏名"];
                                            dr["対象日"] = rowVal["対象日"];
                                            dr["訂正"] = ffKbn2;
                                            dr["訂正日付"] = ffStk2;
                                            _dtSinsei.Rows.Add(dr);

                                            _dtSinsei.Rows.Remove(rowVal);

                                            //rowVal["出休"] = new string[] { rowVal[0], ffKbn2, ffStk2 };//3/1 振休訂正 4/1
                                            break;
                                        }
                                    }
                                    DataRow dr2 = _dtSinsei.NewRow();
                                    dr2["氏名"] = dicSheet["3,8"];
                                    dr2["対象日"] = ffStk2;
                                    dr2["出休"] = ffKbn2.Substring(0, 2);
                                    dr2["出休日付"] = ffStk1;
                                    _dtSinsei.Rows.Add(dr2);

                                    //ffDic.Add(new string[] { ffStk2, ffKbn2.Substring(0, 2), ffStk1 });// 4/1 振休　 1/2
                                }
                                else
                                {
                                    //取消
                                }
                                break;

                            case "振出取消":
                                //振休取消

                                break;

                            case "振休取消":
                                //振出取消

                                break;
                        }
                    }

                };
            }

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
        private static long[] splitAdress(string dicKey)
        {
            string[] sp = dicKey.Split(',');
            return new long[] { long.Parse(sp[0]), long.Parse(sp[1]) }; 
        }

        /// <summary>
        /// Dicアドレス作成
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static string createAdress(long row, long col)
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
