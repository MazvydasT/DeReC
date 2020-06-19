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
