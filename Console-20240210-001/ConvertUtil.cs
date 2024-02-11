using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console_20240210_001
{
    public static class ConvertUtil
    {
        public static decimal ToDecimal(object value)
        {
            if (value == null || value == DBNull.Value) return 0m;

            return Math.Round(Convert.ToDecimal(value),3);
        }
    }
}
