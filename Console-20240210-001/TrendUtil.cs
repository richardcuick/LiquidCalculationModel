using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console_20240210_001
{
    public static class TrendUtil
    {
        public static int MannKendallTest(decimal[] data)
        {
            int n = data.Length;
            int S = 0;

            for (int i = 0; i < n - 1; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    int sign = Math.Sign(data[j] - data[i]);
                    S += sign;
                }
            }

            return S;
        }
    }
}
