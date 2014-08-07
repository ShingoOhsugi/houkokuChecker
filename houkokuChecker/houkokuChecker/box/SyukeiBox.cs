using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace houkokuChecker.box
{
    class SyukeiBox
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
        /// 総稼動
        /// </summary>
        public string Soukadou { get; set; }
        /// <summary>
        /// 基本稼働
        /// </summary>
        public string KihonKado { get; set; }
        /// <summary>
        /// みなし稼働
        /// </summary>
        public string MinashiKado { get; set; }
        /// <summary>
        /// 控除
        /// </summary>
        public string Kojo { get; set; }
        /// <summary>
        /// 普通残業
        /// </summary>
        public string FutsuZan { get; set; }
        /// <summary>
        /// 深夜残業
        /// </summary>
        public string ShinyaZan { get; set; }
        /// <summary>
        /// 早朝稼働
        /// </summary>
        public string SotyoZan { get; set; }
        /// <summary>
        /// 法定休日稼働
        /// </summary>
        public string HouteiKyuzituKado { get; set; }
        /// <summary>
        /// 作業内容
        /// </summary>
        public string SagyoNaiyo { get; set; }

    }
}
