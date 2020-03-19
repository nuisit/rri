using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class EmployeeModel
    {
        public string EmployeeCode { get; set; }
        public string Position { get; set; }
        public decimal CurrentSalary { get; set; }

        public int Age { get; set; }
        /// <summary>
        /// ปีที่เกษียณ
        /// </summary>
        public int RetirementYear { get; set; }
        /// <summary>
        /// ปีที่ตาย
        /// </summary>
        public int DeadYears { get; set; }
        /// <summary>
        /// เงินเมื่อเกษียณ
        /// </summary>
        public decimal ValueAtRetirement => PvdForecastList?.FirstOrDefault((o) => o.WorkYear == RetirementYear).PortfolioValueActual ?? 0;
        /// <summary>
        /// อายุหลังเกษียณ
        /// </summary>
        public int RemainAge { get; set; }
        /// <summary>
        /// เงิน ณ ปีที่ตาย ActualStatus
        /// </summary>
        public decimal ValueAtDead { get; set; }
        /// <summary>
        /// เงิน ณ ปีที่ตาย Amber
        /// </summary>
        public decimal ValueAtDeadAmber { get; set; }
        public int OutOfMoneyYears { get; set; }
        public int AgeOutOfMoney => 0 < OutOfMoneyYears ? OutOfMoneyYears : 100;
        /// <summary>
        /// ระยะเวลาที่สามารถดำรงชีวิตอยู่ได้ด้วยเงินกองทุน
        /// </summary>
        public int LivingWithFund => AgeOutOfMoney - RetirementYear;

        public List<PvdForecastModel> PvdForecastList { get; set; }
        public List<RetirementCostForecastModel> RetirementCostForecastList { get; set; }

        public string Gender { get; set; }
        public int? WorkYears { get; set; }
        public string Status { get; set; }
        public int? NumOfChild { get; set; }
        public decimal RawCollectRate { get; set; }
        public decimal CollectRatePercentage { get; set; }
        public decimal CollectRate => CollectRatePercentage / 100;
        public decimal? EmployersRate { get; set; }
        public decimal PvdAmount { get; set; }

        public decimal RawLifePartPolicyPercentage { get; set; }
        public decimal LifePartPolicyPercentage { get; set; }
        public decimal LifePartPolicy => LifePartPolicyPercentage / 100;

        public decimal RawMarketPolicyPercentage { get; set; }
        public decimal MarketPolicyPercentage { get; set; }
        public decimal MarketPolicy => MarketPolicyPercentage / 100;

        public decimal RawBondPolicyPercentage { get; set; }
        public decimal BondPolicyPercentage { get; set; }
        public decimal BondPolicy => BondPolicyPercentage / 100;

        public decimal RawThaiStockPolicyPercentage { get; set; }
        public decimal ThaiStockPolicyPercentage { get; set; }
        public decimal ThaiStockPolicy => ThaiStockPolicyPercentage / 100;

        public decimal RawRealEstatePolicyPercentage { get; set; }
        public decimal RealEstatePolicyPercentage { get; set; }
        public decimal RealEstatePolicy => RealEstatePolicyPercentage / 100;

        public decimal RawForeignFundPolicyPercentage { get; set; }
        public decimal ForeignFundPolicyPercentage { get; set; }
        public decimal ForeignFundPolicy => ForeignFundPolicyPercentage / 100;

        public decimal CostAmontAt85 { get; set; }
        public decimal CostAmontAtRetirement { get; set; }

        public int ZoneLevel => 0 < ValueAtDead ? 2 : 0 < ValueAtDeadAmber ? 1 : 0;
        public decimal Score => ZoneLevel == 2 ? 1 : ZoneLevel == 1 ? 0.85M : 0;

        public decimal Equities => 1 - (MarketPolicy + BondPolicy);

        public string Group => LifePartPolicy == 1 ? "Target Date" : Equities == 0.5M ? "Mixed" : 0.5M < Equities ? "Fixed Income Oriented" : "Equities Oriented";

        //decimal costAmontAt85 = (decimal)Math.Pow(decimal.ToDouble(_employeeCriteria.CostYearlyAmountWhenAfter85 * Inflation), age + _criterionNumber.BeforCriterion + 1);
        //decimal costAmontAtRetirement = (decimal)Math.Pow(decimal.ToDouble(_employeeCriteria.CostYearlyAmountBetween60And85 * Inflation), retirementAge);
    }
}
