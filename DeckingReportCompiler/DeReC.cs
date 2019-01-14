using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DeckingReportCompiler
{
    public class DeReC
    {
        public delegate void Error(string message);
        public event Error OnError;

        public delegate void Progress(float value);
        public event Progress OnProgress;

        public Dictionary<string, PartPair> Compile(string inputPath, double stepSize)
        {
            using (var reader = new StreamReader(inputPath))
            {
                // skip 4 lines
                for (var i = 0; i < 4; ++i) reader.ReadLine();

                if (reader.EndOfStream && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not have column names row");
                    OnError("Provided clearance file does not have column names row");
                    return null;
                }

                var columnNames = rowToValues(reader.ReadLine());

                var numberColumnIndex = Array.IndexOf(columnNames, "Number");
                if (numberColumnIndex == -1 && OnError != null)
                {
                    OnError("Provide clearance file does not contain Number column");
                    return null;
                }

                var resultColumnIndex = Array.IndexOf(columnNames, "Result");
                if (resultColumnIndex == -1 && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not contain Result column");
                    OnError("Provide clearance file does not contain Number column");
                    return null;
                }

                var frameColumnIndex = Array.IndexOf(columnNames, "Frame");
                if (frameColumnIndex == -1 && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not contain Frame column");
                    OnError("Provided clearance file does not contain Frame column");
                    return null;
                }

                var cadId1ColumnIndex = Array.IndexOf(columnNames, "CADID1");
                if (cadId1ColumnIndex == -1 && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not contain CADID1 column");
                    OnError("Provided clearance file does not contain CADID1 column");
                    return null;
                }

                var cadId2ColumnIndex = Array.IndexOf(columnNames, "CADID2");
                if (cadId1ColumnIndex == -1 && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not contain CADID1 column");
                    OnError("Provided clearance file does not contain CADID1 column");
                    return null;
                }

                var xform1ColumnIndex = Array.IndexOf(columnNames, "XFORM1");
                if (xform1ColumnIndex == -1 && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not contain XFORM1 column");
                    OnError("Provided clearance file does not contain XFORM1 column");
                    return null;
                }

                var xform2ColumnIndex = Array.IndexOf(columnNames, "XFORM2");
                if (xform2ColumnIndex == -1 && OnError != null)
                {
                    //throw new Exception("Provided clearance file does not contain XFORM2 column");
                    OnError("Provided clearance file does not contain XFORM2 column");
                    return null;
                }

                var maxIndex = new int[]
                {
                    numberColumnIndex,
                    resultColumnIndex,
                    frameColumnIndex,
                    cadId1ColumnIndex,
                    cadId2ColumnIndex,
                    xform1ColumnIndex,
                    xform2ColumnIndex
                }.Max();

                var rowCounter = 5;

                var partPairs = new Dictionary<string, PartPair>();

                var numericCleanupRegex = new Regex(@"[^\d\.-]", RegexOptions.Compiled);

                while (!reader.EndOfStream)
                {
                    var dataRow = reader.ReadLine();
                    rowCounter++;

                    if (dataRow == "###EndItem") break;

                    var dataValues = rowToValues(dataRow);

                    if (dataValues.Length <= maxIndex && OnError != null)
                    {
                        //throw new Exception("Row " + rowCounter + " does not conatain all required values");
                        OnError("Row " + rowCounter + " does not conatain all required values");
                        return null;
                    }

                    var number = int.Parse(dataValues[numberColumnIndex]);
                    var result = double.Parse(numericCleanupRegex.Replace(dataValues[resultColumnIndex], ""));

                    var frame = 0;
                    int.TryParse(dataValues[frameColumnIndex], out frame);

                    var part1Hierarchy = cadIdToPartHierarchy(dataValues[cadId1ColumnIndex]);
                    var part2Hierarchy = cadIdToPartHierarchy(dataValues[cadId2ColumnIndex]);
                    var xform1 = dataValues[xform1ColumnIndex];
                    var xform2 = dataValues[xform2ColumnIndex];

                    var partPairId = string.Join("|", part1Hierarchy) + "#" + xform1 + "#" + string.Join("|", part2Hierarchy) + "#" + xform2;

                    if (!partPairs.ContainsKey(partPairId))
                    {
                        var partPair = new PartPair(part1Hierarchy, xform1, part2Hierarchy, xform2);
                        var test = partPair.Part1DS;

                        partPairs.Add(partPairId, partPair);
                    }

                    var partPairClearances = partPairs[partPairId].Clearances;

                    if (!partPairClearances.ContainsKey(frame))
                        partPairClearances.Add(frame, new Clearance(number, result, frame * stepSize));

                    if (OnProgress != null)
                        OnProgress((float)reader.BaseStream.Position / (float)reader.BaseStream.Length);
                }

                return partPairs;
            }
        }

        private static Regex rowCleanupRegex = new Regex(@"^\d+", RegexOptions.Compiled);
        private static string[] rowToValues(string row)
        {
            return rowCleanupRegex.Replace(row, "").Trim(new char[] { '"' }).Split(new string[] { "\",\"" }, StringSplitOptions.None);
        }

        private static Regex cadIdCleanupRegex = new Regex(@"^.*?\0", RegexOptions.Compiled);
        private static Regex cadIdSplitRegex = new Regex(@"\.(?:part|asm);-?\d+;-?\d+:\0", RegexOptions.Compiled);
        private static string[] cadIdToPartHierarchy(string cadId)
        {
            var partHierarchy = cadIdSplitRegex.Split(cadIdCleanupRegex.Replace(cadId, ""));

            Array.Resize(ref partHierarchy, partHierarchy.Length - 1);

            return partHierarchy;
        }
    }
}