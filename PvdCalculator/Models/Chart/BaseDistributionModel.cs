using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PvdCalculator.Models.Chart
{
    public class DistributionBaseModel
    {
        private Color _color;

        public int ZoneLevel { get; set; }
        public Color Color
        {
            get
            {
                switch (ZoneLevel)
                {
                    case 0:
                        _color = ColorTranslator.FromHtml("#FA5B65");
                        break;
                    case 1:
                        _color = ColorTranslator.FromHtml("#FFC107");
                        break;
                    case 2:
                        _color = ColorTranslator.FromHtml("#42DDA0");
                        break;
                }

                return _color;
            }
        }
    }
}
