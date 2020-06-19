using System.Collections.Generic;

namespace DeckingReportCompiler
{
    public class ClearanceFileMetaData
    {
        public string Version { get; private set; }
        public string MagicNumber { get; private set; }
        public string Header { get; private set; }

        public Dictionary<int, string> DataRows { get; private set; }
        public Dictionary<string, PartPair> PartPairs { get; private set; }

        public ClearanceFileMetaData(string version, string magicNumber, string header, int numberOfDataRows)
        {
            Version = version;
            MagicNumber = magicNumber;
            Header = header;
            DataRows = new Dictionary<int, string>(numberOfDataRows);
            PartPairs = new Dictionary<string, PartPair>();
        }
    }
}
