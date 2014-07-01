using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace houkokuChecker
{
    class SyukeiTable : DataTable
    {
        public SyukeiTable()
        {
            Columns.Add("氏名", typeof(string));
            Columns.Add("総稼動", typeof(string));
            Columns.Add("基本稼働", typeof(string));
            Columns.Add("みなし稼働", typeof(string));
            Columns.Add("控除", typeof(string));
            Columns.Add("普通残業", typeof(string));
            Columns.Add("深夜残業", typeof(string));
            Columns.Add("早朝稼働", typeof(string));
            Columns.Add("法定休日稼働", typeof(string));
            Columns.Add("作業内容", typeof(string));
        }



    }
}
