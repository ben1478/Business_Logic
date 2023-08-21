using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Business_Logic.Entity
{
    public class ReportEntity
    {

        public class Achievement
        {
            public String check_amount { get; set; }
            public String yyyymm { get; set; }
            public String DisplayName { get; set; }
            public String CheckVal { get; set; }
        }

        public class YearAmountInfo
        {
            public int yyyy { get; set; }
            public List<AmountInfo> AmountInfos=new List<AmountInfo>();
        }
        public class YearAmountInfoByYear
        {
            public int yyyy { get; set; }
            public int amount { get; set; }
        }



        public class AmountInfo
        {
            public int yyyy { get; set; }
            public int mm { get; set; }
            public int amount { get; set; }
        }
       

    }
}
