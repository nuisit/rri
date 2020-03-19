using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models.Chart
{
    public class ZoneDistributionModel : DistributionBaseModel
    {
        public int Value { get; set; }
        public decimal ValuePercentage { get; set; }
    }
}
