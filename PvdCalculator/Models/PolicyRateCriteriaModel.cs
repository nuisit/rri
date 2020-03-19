using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class PolicyRateCriteriaModel
    {
        public decimal MarketPolicyPercentage { get; set; } = 1.75M;
        public decimal MarketPolicy => MarketPolicyPercentage / 100;
        public decimal BondPolicyPercentage { get; set; } = 2.5M;
        public decimal BondPolicy => BondPolicyPercentage / 100;
        public decimal ThaiStockPolicyPercentage { get; set; } = 6M;
        public decimal ThaiStockPolicy => ThaiStockPolicyPercentage / 100;
        public decimal RealEstatePolicyPercentage { get; set; } = 6M;
        public decimal RealEstatePolicy => RealEstatePolicyPercentage / 100;
        public decimal ForeignFundPolicyPercentage { get; set; } = 5M;
        public decimal ForeignFundPolicy => ForeignFundPolicyPercentage / 100;
    }
}
