using houkokuChecker.util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
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
        private const long C_IDX_KAISU = 1;        //列No 回数
        private const long C_IDX_TODOKE = 2;       //列No 届出日

        private const long C_IDX_YK_STK = 3;       //列No 有給休暇 取得日
        private const long C_IDX_YK_KBN = 4;       //列No 有給休暇 区分

        private const long C_IDX_TK_STK = 3;       //列No 特別休 取得日
        private const long C_IDX_TK_KBN = 5;       //列No 特別休 区分

        private const long C_IDX_DK_STK = 4;       //列No 代休暇 取得日
        private const long C_IDX_DK_KBN = 5;       //列No 代休暇 区分

        private const long C_IDX_FF_KBN1 = 3;      //列No 振出・振休 区分1
        private const long C_IDX_FF_STK1 = 4;      //列No 振出・振休 取得日1
        private const long C_IDX_FF_KBN2 = 5;      //列No 振出・振休 区分2
        private const long C_IDX_FF_STK2 = 6;      //列No 振出・振休 取得日2

        //報告書
        private const long R_IDX_TIME_KIHON = 17;        //行No 基本稼働時間
        private const long R_IDX_TIME_MINASHI_H = 16;    //行No みなし補足時間
        private const long R_IDX_TIME_MINASHI = 39;      //行No みなし稼働時間
        private const long R_IDX_TIME_KOZYO_T = 24;      //行No 控除(遅刻)
        private const long R_IDX_TIME_KOZYO_S = 29;      //行No 控除(早退)
        private const long R_IDX_TIME_KOZYO_G = 34;      //行No 控除(外出)
        private const long R_IDX_TIME_FUTUZAN = 51;      //行No 普通残業
        private const long R_IDX_TIME_SINZAN = 55;       //行No 深夜残業
        private const long R_IDX_TIME_SOUZAN = 47;       //行No 早朝残業
        private const long R_IDX_TIME_HOTEIZAN = 63;     //行No 法定休日稼働
        private const long R_IDX_TIME_HOTEISINZAN = 67;  //行No 法定休日深夜稼働
        private const string R_IDX_SAGYOCD = "97,2";     //行列No PJコード
        private const string R_IDX_SAGYONAIYO = "98,6";  //行列No PJ内容

        private const long C_IDX_HIDUKE_START = 21;      //列No 「1日」のセル

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

            foreach (string selMember in lbMember.SelectedItems)
            {
                //申請書Open
                string sinseiFilePath =
                    _rootPath + string.Format(SINSEI_FMT, selMember);

                Dictionary<string, Dictionary<string, string>> dicSinsei =
                    getExcelData(sinseiFilePath);

                //申請書なしエラー
                if (dicSinsei.Count == 0)
                {
                    MessageBox.Show(selMember + " の申請書を格納してください。:" + sinseiFilePath);
                    continue;
                }

                //報告書Open
                string hokokuFilePath =
                    _rootPath + string.Format(HOUKOKU_FMT,
                                              calCheckTaisyo.SelectedDate.Value.Year.ToString(),
                                              calCheckTaisyo.SelectedDate.Value.Month.ToString("00"),
                                              selMember);

                Dictionary<string, Dictionary<string, string>> dicHoukoku =
                    getExcelData(hokokuFilePath);

                //報告書なしエラー
                if (dicHoukoku.Count == 0)
                {
                    MessageBox.Show(selMember + " の報告書を格納してください。:" + hokokuFilePath);
                    continue;
                }

                //チェック結果追記
                txtCheckResult.Text += CheckMain(dicSinsei, dicHoukoku, calCheckTaisyo.SelectedDate.Value);
            }

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
            SyukeiTable dtKekka = new SyukeiTable();

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

            foreach (string selMember in lbMember.SelectedItems)
            {
                //報告書Open
                string hokokuFilePath =
                    _rootPath + string.Format(HOUKOKU_FMT,
                                              calSyukeiTaisyo.SelectedDate.Value.Year.ToString(),
                                              calSyukeiTaisyo.SelectedDate.Value.Month.ToString("00"),
                                              selMember);

                Dictionary<string, Dictionary<string, string>> dicHoukoku =
                    getExcelData(hokokuFilePath);

                //報告書なしエラー
                if (dicHoukoku.Count == 0)
                {
                    MessageBox.Show(selMember + " の報告書を格納してください。:" + hokokuFilePath);
                    continue;
                }

                //集計データ取得
                dtKekka.Merge(GetSyukeiData(dicHoukoku, calSyukeiTaisyo.SelectedDate.Value));
            }

            dgSyukeiResult.ItemsSource = dtKekka.DefaultView;

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
            SinseiTable dtSinsei = new SinseiTable();

            if (lbMember.SelectedItems.Count == 0)
            {
                MessageBox.Show("メンバを選択してください");
                return;
            }

            foreach (string selMember in lbMember.SelectedItems)
            {
                //申請書Open
                string sinseiFilePath =
                    _rootPath + string.Format(SINSEI_FMT, selMember);

                Dictionary<string, Dictionary<string, string>> dicSinsei = 
                    getExcelData(sinseiFilePath);

                //申請書なしエラー
                if (dicSinsei.Count == 0)
                {
                    MessageBox.Show(selMember + " の申請書を格納してください。:" + sinseiFilePath);
                    continue;
                }

                //申請情報ロード
                dtSinsei.Merge(GetSinseiData(dicSinsei), true);
            }

            //ソート
            dtSinsei.DefaultView.Sort = "対象日";

            dgSinseiResult.ItemsSource = dtSinsei.DefaultView;

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
        /// 報告書チェックメイン
        /// </summary>
        /// <param name="sinseiBook"></param>
        /// <param name="houkokuBook"></param>
        private string CheckMain(
            Dictionary<string, Dictionary<string, string>> sinseiBook,
            Dictionary<string, Dictionary<string, string>> houkokuBook,
            DateTime selDate)
        {
            string result = string.Empty;

            //申請情報取込
            SinseiTable sinseiData = GetSinseiData(sinseiBook);

            //報告情報チェック

            // NGチェック
            Dictionary<string, string> dicHoukokuAlldata = houkokuBook["社員"];
            int syainCode = int.Parse(dicHoukokuAlldata["4,3"]);
            string syainName = dicHoukokuAlldata["4,7"];

            foreach (KeyValuePair<string, string> hitData in dicHoukokuAlldata.Where(p => p.Value == ("NG")))
            {
                long[] hitAd = splitAdress(hitData.Key);
                result += 
                    string.Format("{0} NGがあります。 行No: {1} 列No: {2} \n", syainName, hitAd[0], hitAd[1]);
            };

            // 指定期間に限定
            for (int i = 0; i < selDate.Day; i++)
            {
                // PJNo必須チェック
                // 休日必須チェック

                // 申請チェック
                foreach(DataRow dr in sinseiData.Select(string.Format("社員コード = {0} AND ", syainCode)))
                {

                }

            }

            return result;
        }

        /// <summary>
        /// 集計処理メイン
        /// </summary>
        /// <param name="dicSinsei"></param>
        /// <param name="selDate"></param>
        private SyukeiTable GetSyukeiData(
            Dictionary<string, Dictionary<string, string>> dicHoukoku,
            DateTime selDate)
        {
            SyukeiTable locKekka = new SyukeiTable();

            //報告情報解析
            Dictionary<string, string> dicHoukokuAlldata = dicHoukoku["作業報告書"];
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
            DataRow drKekka = locKekka.NewRow();
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
            drKekka["作業内容"] = dicHoukokuAlldata[R_IDX_SAGYOCD] + " : " + dicHoukokuAlldata[R_IDX_SAGYONAIYO];
            locKekka.Rows.Add(drKekka);

            return locKekka;
        }

        /// <summary>
        /// 申請書ロード
        /// </summary>
        /// <returns></returns>
        private SinseiTable GetSinseiData(Dictionary<string, Dictionary<string, string>> dicSinsei)
        {
            //申請情報解析

            SinseiTable locSinsei = new SinseiTable();

            //ブック毎にループ
            foreach (Dictionary<string, string> dicSheet in dicSinsei.Values)
            {
                //社員コード、名前が入ってなかったらスキップ
                if (dicSheet.ContainsKey("3,6") == false ||
                    dicSheet.ContainsKey("3,8") == false )
                {
                    break;
                }

                int syainCode = int.Parse(dicSheet["3,6"]); // 社員コード
                string syainName = dicSheet["3,8"]; // 社員名

                #region 有給休暇の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("有給休暇")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 8; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_YK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = dicSheet[createAdress(rowBase, C_IDX_YK_KBN)];

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                #region 病気有給休暇の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("病気有給休暇")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_YK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = string.Format("病{0}", dicSheet[createAdress(rowBase, C_IDX_YK_KBN)]);

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                #region 特別有休の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("特別有休")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_TK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = string.Format("特有{0}", dicSheet[createAdress(rowBase, C_IDX_TK_KBN)]);

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                //特別有休（連日）の解析　→自分で見てと通知だけする

                #region 特別無休の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("特別無休")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_TK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = string.Format("特無{0}", dicSheet[createAdress(rowBase, C_IDX_TK_KBN)]);

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                //特別無休（連日）の解析　→自分で見てと通知だけする

                #region 振出・振休の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Contains("振出・振休")))
                {

                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「回数」が空だったら次のタイトルセルへ
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_KAISU)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_KAISU)]))
                        {
                            break;
                        }

                        //「区分１」が空だったら次の行へ
                        string ffKbn1 = dicSheet[createAdress(rowBase, C_IDX_FF_KBN1)];
                        if (string.IsNullOrWhiteSpace(ffKbn1))
                        {
                            continue;
                        }

                        //行情報を取得
                        string ffStk1 = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_FF_STK1)]).ToString("yyyy/MM/dd");
                        string ffKbn2 = dicSheet[createAdress(rowBase, C_IDX_FF_KBN2)];
                        string ffStk2 = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_FF_STK2)]).ToString("yyyy/MM/dd");

                        //「区分1」を判定
                        switch (ffKbn1)
                        {
                            case "振出":
                            case "振休":

                                switch (ffKbn2)
                                {
                                    case "振休":
                                    case "振出":
                                        //振出 → 振休
                                        //振休 → 振休

                                        // x/x　「振出」 (y/y(振休日))
                                        DataRow dr1 = locSinsei.NewRow();
                                        dr1["社員コード"] = syainCode;
                                        dr1["氏名"] = syainName;
                                        dr1["対象日"] = ffStk1;
                                        dr1["出休"] = ffKbn1;
                                        dr1["出休日付"] = ffStk2;
                                        if (ffKbn1.Equals("振出"))
                                        {
                                            dr1["元申請日"] = ffStk1;
                                        }
                                        else
                                        {
                                            dr1["元申請日"] = ffStk2;
                                        }
                                        locSinsei.Rows.Add(dr1);

                                        // y/y　「振休」 (x/x(振出日))
                                        dr1 = locSinsei.NewRow();
                                        dr1["社員コード"] = syainCode;
                                        dr1["氏名"] = syainName;
                                        dr1["対象日"] = ffStk2;
                                        dr1["出休"] = ffKbn2;
                                        dr1["出休日付"] = ffStk1;
                                        if (ffKbn1.Equals("振出"))
                                        {
                                            dr1["元申請日"] = ffStk1;
                                        }
                                        else
                                        {
                                            dr1["元申請日"] = ffStk2;
                                        }
                                        locSinsei.Rows.Add(dr1);

                                        break;

                                    case "振休訂正":
                                    case "振出訂正":
                                        //振出 → 振休訂正
                                        //振休 → 振出訂正
                                        // y/y　「振休」 (x/x(振出日))　→ y/y　「振休訂正」 (y'/y'(訂正振休日))　
                                        for (int r = 0; r < locSinsei.Rows.Count; r++)
                                        {
                                            DataRow rowVal = locSinsei.Rows[r];

                                            if (ffStk1.Equals(rowVal["出休日付"]) &&
                                                ffStk1.Equals(rowVal["元申請日"]))
                                            {
                                                //訂正書換
                                                DataRow dr21 = locSinsei.NewRow();
                                                dr21["社員コード"] = rowVal["社員コード"];
                                                dr21["氏名"] = rowVal["氏名"];
                                                dr21["対象日"] = rowVal["対象日"];
                                                dr21["訂正"] = ffKbn2;
                                                dr21["訂正日付"] = ffStk2;
                                                dr21["元申請日"] = rowVal["元申請日"];
                                                locSinsei.Rows.Add(dr21);

                                                //消す
                                                locSinsei.Rows.Remove(rowVal);

                                                break;
                                            }
                                        }

                                        // y'/y' 「振休」 (x/x(振出日))
                                        DataRow dr22 = locSinsei.NewRow();
                                        dr22["社員コード"] = syainCode;
                                        dr22["氏名"] = syainName;
                                        dr22["対象日"] = ffStk2;
                                        dr22["出休"] = ffKbn2.Substring(0, 2);
                                        dr22["出休日付"] = ffStk1;
                                        dr22["元申請日"] = ffStk1;
                                        locSinsei.Rows.Add(dr22);

                                        break;

                                    //振出 → 振出取消　はNG
                                    //振出 → 振休取消　はNG
                                    //振休 → 振出取消　はNG
                                    //振休 → 振休取消　はNG

                                }
                                break;

                            case "振出取消":
                            case "振休取消":

                                //振出取消 → 振休取消
                                //振休取消 → 振出取消
                                for (int r = 0; r < locSinsei.Rows.Count; r++)
                                {
                                    DataRow rowVal = locSinsei.Rows[r];

                                    if (ffStk1.Equals(rowVal["元申請日"]) ||
                                        ffStk2.Equals(rowVal["元申請日"]))
                                    {
                                        // 削除
                                        locSinsei.Rows.Remove(rowVal);
                                    }
                                }

                                break;

                            //振休訂正取消は無視。振休訂正すればいいじゃん。

                        }
                    }
                }

                #endregion

                #region 休日出勤の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("休日出勤")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_YK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = dicSheet[createAdress(rowBase, C_IDX_YK_KBN)];

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                #region 代休暇の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("代休暇")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_DK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = dicSheet[createAdress(rowBase, C_IDX_DK_KBN)];

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                #region 遅刻・早退・外出の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("遅刻・早退・外出")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_YK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = dicSheet[createAdress(rowBase, C_IDX_YK_KBN)];

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["遅刻・早退・外出"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["遅刻・早退・外出"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                #region 欠勤の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("欠勤")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_YK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = dicSheet[createAdress(rowBase, C_IDX_YK_KBN)];

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["その1,2"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["その1,2"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

                //欠勤（連日）の解析 →自分で見てと通知だけする

                #region 時間外対応の解析

                foreach (KeyValuePair<string, string> hitData in dicSheet.Where(p => p.Value.Equals("時間外対応")))
                {
                    long[] furiTitleAd = splitAdress(hitData.Key);  //解析対象タイトルセル

                    //タイトルセルの次の行から解析開始。
                    for (int i = 2; i < 500; i++)
                    {
                        long rowBase = furiTitleAd[0] + i;

                        //「届出日」が空だったら終了
                        if (dicSheet.ContainsKey(createAdress(rowBase, C_IDX_TODOKE)) == false ||
                            string.IsNullOrWhiteSpace(dicSheet[createAdress(rowBase, C_IDX_TODOKE)]))
                        {
                            break;
                        }

                        //行情報を取得
                        string ykStk = DatetimeUtil.ToDate(dicSheet[createAdress(rowBase, C_IDX_YK_STK)]).ToString("yyyy/MM/dd");
                        string ykKbn = dicSheet[createAdress(rowBase, C_IDX_YK_KBN)];

                        //「区分」を判定
                        if (ykKbn.Contains("取消"))
                        {
                            string ykKbnBase = ykKbn.Substring(0, ykKbn.Length - 2);

                            //取り消し申請の場合
                            for (int r = 0; r < locSinsei.Rows.Count; r++)
                            {
                                DataRow rowVal = locSinsei.Rows[r];

                                if (ykStk.Equals(rowVal["対象日"]) &&
                                    ykKbnBase.Equals(rowVal["当番・緊急・計画"]))
                                {
                                    //消す
                                    locSinsei.Rows.Remove(rowVal);
                                }
                            }
                        }
                        else
                        {
                            //通常申請
                            DataRow drYK = locSinsei.NewRow();
                            drYK["社員コード"] = syainCode;
                            drYK["氏名"] = syainName;
                            drYK["対象日"] = ykStk;
                            drYK["当番・緊急・計画"] = ykKbn;
                            locSinsei.Rows.Add(drYK);
                        }
                    }
                }

                #endregion

            }

            return locSinsei;
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

        /// <summary>
        /// Excel → Dictionary変換
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>dicのキーはシート名</returns>
        private Dictionary<string, Dictionary<string, string>> getExcelData(string filePath)
        {
            Dictionary<string, Dictionary<string, string>> res =
                new Dictionary<string, Dictionary<string, string>>();

            if (File.Exists(filePath) == false)
            {
                return res;
            }

            Microsoft.Office.Interop.Excel.Application xlsFile
                = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook xlsBook = null;

            try
            {
                xlsBook = xlsFile.Workbooks.Open(filePath);

                foreach (Microsoft.Office.Interop.Excel.Worksheet xlsSheet in xlsBook.Sheets)
                {
                    Dictionary<string, string> dicSheet = new Dictionary<string, string>();

                    Microsoft.Office.Interop.Excel.Range aRange = xlsSheet.UsedRange;

                    object[,] dataAll = aRange.get_Value();

                    if (dataAll == null)
                    {
                        res.Add(xlsSheet.Name, dicSheet);
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

                    res.Add(xlsSheet.Name, dicSheet);
                }

                return res;

            }
            finally
            {
                //Close
                if (xlsBook != null)
                {
                    xlsBook.Close(false);
                }
            }
        }

    }
}
