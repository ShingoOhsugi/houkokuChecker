using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace houkokuChecker
{
    class SinseiTable : DataTable
    {
        public SinseiTable()
        {
            Columns.Add("氏名", typeof(string));
            Columns.Add("対象日", typeof(string));
            Columns.Add("その1", typeof(string));
            Columns.Add("その2", typeof(string));
            Columns.Add("出休", typeof(string));
            Columns.Add("出休日付", typeof(string));
            Columns.Add("訂正", typeof(string));
            Columns.Add("訂正日付", typeof(string));
            Columns.Add("当番・緊急・計画", typeof(string));
        }

    }
}
