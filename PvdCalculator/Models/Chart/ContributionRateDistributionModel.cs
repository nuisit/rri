using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models.Chart
{
    public class ContributionRateDistributionModel : DistributionBaseModel
    {
        public decimal ContributionRate { get; set; }
        public string Key { get; set; }
        public int Value { get; set; }
    }
}
