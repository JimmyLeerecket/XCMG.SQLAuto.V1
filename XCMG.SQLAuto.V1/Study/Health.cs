using NPOI.SS.Formula.Functions;
using RekTec.Crm.Common.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XCMG.SQLAuto.V1.Study
{
    public static class Health
    {
        public static string GetFatBurningHeartRateZone(string age)
        {
            int ageNew = Cast.ConToInt(age);
            if (ageNew <= 0)
            {
                return "年龄输入有问题！";
            }

            // （207-0.7×年龄）×0.6~0.8
            double maxHeartRate = 220;
            double maxHeartRateZone = 0;
            double minHeartRateZone = 0;

            //maxHeartRateZone = (maxHeartRate - 0.7 * (ageNew * 1.0)) * 0.8;
            //minHeartRateZone = (maxHeartRate - 0.7 * (ageNew * 1.0)) * 0.6;
            maxHeartRateZone = (maxHeartRate - ageNew * 1.0) * 0.8;
            minHeartRateZone = (maxHeartRate - ageNew * 1.0) * 0.6;

            return $"您的燃脂区间：{minHeartRateZone:0.00}-{maxHeartRateZone:0.00}次/分钟。";
        }
    }
}
