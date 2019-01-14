using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DeckingReportCompiler
{
    public class Clearance
    {
        public int Number { get; private set; }
        public double Result { get; private set; }
        public double DistanceToMarried { get; private set; }

        public Clearance(int number, double result, double distanceToMarried)
        {
            Number = number;
            Result = result;
            DistanceToMarried = distanceToMarried;
        }
    }
}
