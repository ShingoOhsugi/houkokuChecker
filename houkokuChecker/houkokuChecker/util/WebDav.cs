using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace houkokuChecker.util
{
    public class WebDAV
    {
        public static void Connect()
        {
            // HttpWebRequestを初期化、メゾットなど必要な設定を行う
            HttpWebRequest webReq = (HttpWebRequest)HttpWebRequest.Create("http://share.digicomnet.co.jp/ndcngy/");
            webReq.Method = "COPY";
            webReq.Headers.Add("Destination", @"file:///C:\_work");

            //　認証設定
            webReq.Credentials = new NetworkCredential("diginagoya", "72UIwGzQ");

            // 結果ステータス取得
            HttpWebResponse res = (HttpWebResponse)webReq.GetResponse();
            Console.WriteLine("Status Code: {0}", res.StatusCode);
            res.Close();

            Console.ReadLine();
        }
    }
}
