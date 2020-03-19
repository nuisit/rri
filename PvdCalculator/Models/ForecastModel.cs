using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models
{
    public class PvdForecastModel
    {
        public int Age { get; set; }
        public int WorkYear { get; set; }
        public decimal Salary { get; set; }
        /// <summary>
        /// ผลตอบแทนโดยประมาณ
        /// </summary>
        public decimal ReturnActual { get; set; }
        /// <summary>
        /// รายได้ต่อปีจนถึงขั้นเกษียณ
        /// </summary>
        public decimal Contribution { get; set; }
        /// <summary>
        /// อัตราสะสมก่อนเกษียณ
        /// </summary>
        public decimal PortfolioValueActual { get; set; }

        /// <summary>
        /// Exp
        /// </summary>
        //public decimal EstimatedExp { get; set; }
        /// <summary>
        /// กระแสการเงิน
        /// </summary>
        public decimal FundFlowSummary { get; set; }
    }
}
