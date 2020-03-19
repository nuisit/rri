using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class CriterionNumberModel
    {
        private decimal _pppAvgAustAndThai = 2.840906034M;
        private decimal _exchangeRate = 21.5M;
        private int _ageOfCriterio = 85;
        private int _beforCriterion = 25;
        public decimal PPPAvgAustAndThai
        {
            get => _pppAvgAustAndThai;
            set => _pppAvgAustAndThai = value;
        }

        public decimal ExchangeRate
        {
            get => _exchangeRate;
            set => _exchangeRate = value;
        }
        public int AgeOfCriterio
        {
            get => _ageOfCriterio;
            set => _ageOfCriterio = value;
        }
        public int BeforCriterion
        {
            get => _beforCriterion;
            set => _beforCriterion = value;
        }
    }
}
