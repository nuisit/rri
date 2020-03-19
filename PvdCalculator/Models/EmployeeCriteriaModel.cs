using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class EmployeeCriteriaModel
    {
        public decimal CostMonthlyAmount { get; set; }
        public int Lifespan { get; set; }

        public decimal EmployersRatePercentage { get; set; }
        public decimal EmployersRate => EmployersRatePercentage / 100;
        public int RetirementAge { get; set; }
        public decimal SalaryIncreaseRatePercentage { get; set; }
        public decimal SalaryIncreaseRate => SalaryIncreaseRatePercentage / 100;
        public decimal SalaryCap { get; set; }

        public decimal CostYearlyAmountBetween60And85 => CostMonthlyAmount * 12;
        public decimal CostYearlyAmountWhenAfter85 => CostMonthlyAmount * 12;

        public int AutomaticEscalation { get; set; }
        public decimal CollectAdjustRatePercentage { get; set; }
        public decimal CollectAdjustRate => CollectAdjustRatePercentage / 100;
        public bool IsDoTargetDate { get; set; }
    }
}
