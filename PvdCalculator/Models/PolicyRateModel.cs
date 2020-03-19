using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class PolicyRateModel
    {
        public decimal MarketPolicyPercentage { get; set; }
        public decimal MarketPolicy => MarketPolicyPercentage / 100;
        public decimal BondPolicyPercentage { get; set; }
        public decimal BondPolicy => BondPolicyPercentage / 100;
        public decimal ThaiStockPolicyPercentage { get; set; }
        public decimal ThaiStockPolicy => ThaiStockPolicyPercentage / 100;
        public decimal RealEstatePolicyPercentage { get; set; }
        public decimal RealEstatePolicy => RealEstatePolicyPercentage / 100;
        public decimal ForeignFundPolicyPercentage { get; set; }
        public decimal ForeignFundPolicy => ForeignFundPolicyPercentage / 100;

        public RangeModel<int> AgeRange { get; set; }
        public decimal MMPercentage { get; set; }
        public decimal MM => MMPercentage / 100;
        public decimal FI => 1 - (EQ + REITs + GlobalEQ) - MM;
        public decimal EQPercentage { get; set; }
        public decimal EQ => EQPercentage / 100;
        public decimal REITsPercentage { get; set; }
        public decimal REITs => REITsPercentage / 100;
        public decimal GlobalEQPercentage { get; set; }
        public decimal GlobalEQ => GlobalEQPercentage / 100;

        public decimal Value => (MarketPolicy * MM) + (BondPolicy * FI) + (ThaiStockPolicy * EQ) + (RealEstatePolicy * REITs) + (ForeignFundPolicy * GlobalEQ);
    }
}
