using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace houkokuChecker.box
{
    class ShinseiBox
    {
        /// <summary>
        /// 社員コード
        /// </summary>
        public string ShainCd { get; set; }
        /// <summary>
        /// 氏名
        /// </summary>
        public string Shimei { get; set; }
        /// <summary>
        /// 対象日
        /// </summary>
        public string TaisyoDt { get; set; }
        /// <summary>
        /// 遅刻・早退・外出
        /// </summary>
        public string TikokuEtc { get; set; }
        /// <summary>
        /// その1,2
        /// </summary>
        public string Sono12 { get; set; }
        /// <summary>
        /// 出休
        /// </summary>
        public string Syutukyu { get; set; }
        /// <summary>
        /// 出休日付
        /// </summary>
        public string SyutukyuDt { get; set; }
        /// <summary>
        /// 訂正
        /// </summary>
        public string Teisei { get; set; }
        /// <summary>
        /// 訂正日付
        /// </summary>
        public string TeiseiDt { get; set; }
        /// <summary>
        /// 当番・緊急・計画
        /// </summary>
        public string Toban { get; set; }
        /// <summary>
        /// 元申請日
        /// </summary>
        public string MotoShinseiDt { get; set; }

    }
}
