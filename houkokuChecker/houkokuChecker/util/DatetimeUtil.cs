using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace houkokuChecker.util
{
    class DatetimeUtil
    {
        public static DateTime ToDate(string dt)
        {
            return DateTime.ParseExact(dt, "yyyy/MM/dd H:mm:ss", CultureInfo.InvariantCulture);

        }
    }
}
