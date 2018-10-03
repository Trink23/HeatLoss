using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace HeatLoss
{
    class Calculations
    {
        /**
         * Room Watts calculation
         */
        public static double WattsCalculate(double volum, float airChanges,int tempDif)
        {
            
            return volum * airChanges * tempDif * 0.33;
        }

        /**
         * Wall  Watts calculation
         */
        public static double WallWattsCalculate(double area,int tempDif,float uValue)
        {
            return area * tempDif * uValue;
        }
    }


}
