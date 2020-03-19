using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class RetirementCostForecastModel
    {
        public int Age { get; set; }
        public int Year { get; set; }
        public decimal EstimatedExp { get; set; }
        public decimal FundedStatusActual { get; set; }
        public decimal FundedStatusAmberZone { get; set; }
        public int FundEndedAge { get; set; }
    }
}
