using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DeckingReportCompiler
{
    public class ExcelManager
    {
        public delegate void Progress(float value);
        public event Progress OnProgress;

        public void Write(Dictionary<string, PartPair> partPairs, Stream excelOutputStream)
        {
            if (partPairs == null)
                throw new Exception("partPairs is null");

            if (excelOutputStream == null)
                throw new Exception("outputStream is null");

            if (!excelOutputStream.CanWrite)
                throw new Exception("Cannot write to provided stream");

            using (var templateStream = new MemoryStream(Properties.Resources.Template))
            using (var templatePackage = new ExcelPackage(templateStream))
            {
                var workbook = templatePackage.Workbook;

                var summarySheet = workbook.Worksheets["Summary"];
                var spaklinesDataSheet = workbook.Worksheets["SparklinesData"];

                var partPairsCount = partPairs.Count;

                var summaryRowCount = partPairsCount + 1;

                bool cpsc1Present = false, cpsc2Present = false, mfgArea1Present = false, mfgArea2Present = false;

                var summaryTableRangeHeaderRow = 4;
                var summaryTableRangeFirstDataRow = summaryTableRangeHeaderRow + 1;
                var summaryTableRangeLastRow = summaryRowCount + summaryTableRangeHeaderRow - 1;

                int summaryTableRowNumber = 1;

                summarySheet.InsertRow(summaryTableRangeFirstDataRow, partPairsCount - 1, summaryTableRangeLastRow);

                //var resultsColumnsAddress = String.Format("J{0}:J{1},L{0}:M{1}", summaryTableRangeFirstDataRow, summaryTableRangeLastRow);
                var resultsColumnsAddress = String.Format("J{0}:K{1}", summaryTableRangeFirstDataRow, summaryTableRangeLastRow);

                var conditionalFormatting = summarySheet.ConditionalFormatting;
                conditionalFormatting[0].Address = new ExcelAddress(resultsColumnsAddress);
                conditionalFormatting[1].Address = new ExcelAddress(resultsColumnsAddress);
                //conditionalFormatting[2].Address = new ExcelAddress(String.Format("I{0}:I{1}", summaryTableRangeFirstDataRow, summaryTableRangeLastRow));

                var sparklinesDataRow = 1;

                foreach (var partPairKeyValue in partPairs)
                {
                    var partPairId = partPairKeyValue.Key;
                    var partPair = partPairKeyValue.Value;

                    if (partPair.Part1CPSC != null) cpsc1Present = true;
                    if (partPair.Part2CPSC != null) cpsc2Present = true;
                    if (partPair.Part1MfgArea != null) mfgArea1Present = true;
                    if (partPair.Part2MfgArea != null) mfgArea2Present = true;

                    OnProgress?.Invoke((float)summaryTableRowNumber / (float)partPairsCount);

                    var currentRowNumber = summaryTableRangeHeaderRow + summaryTableRowNumber++;

                    var summaryTableRow = summarySheet.Cells[String.Format("A{0}:N{0}", currentRowNumber)];
                    summaryTableRow.LoadFromArrays(new List<object[]>
                    {
                        new object[]
                        {
                            //partPairId,
                            partPair.WorstCaseClearance.Number,
                            //partPair.Number,
                            
                            //null,
                            
                            partPair.Part1MfgArea,
                            partPair.Part1CPSC,
                            partPair.Part1Hierarchy[partPair.Part1Hierarchy.Length - 1],
                            
                            //null,
                            
                            partPair.Part2MfgArea,
                            partPair.Part2CPSC,
                            partPair.Part2Hierarchy[partPair.Part2Hierarchy.Length - 1],

                            null,

                            //partPair.FirstDetectedClearance == null ? (object)null : partPair.FirstDetectedClearance.DistanceToMarried,
                            //partPair.FirstDetectedClearance == null ? (object)null : partPair.FirstDetectedClearance.Result,

                            partPair.WorstCaseClearance == null ? (object)null : partPair.WorstCaseClearance.DistanceToMarried,
                            partPair.WorstCaseClearance == null ? (object)null : partPair.WorstCaseClearance.Result,

                            partPair.ClearanceAtDecked == null ? (object)null : partPair.ClearanceAtDecked.Result,

                            null
                        }
                    });

                    if (partPair.Part1Process != null)
                        summarySheet.Cells[currentRowNumber, 2].AddComment(partPair.Part1Process, "_").AutoFit = true;

                    if (partPair.Part1DS != null)
                        summarySheet.Cells[currentRowNumber, 3].AddComment(partPair.Part1DS, "_").AutoFit = true;

                    if (partPair.Part1Ancestors != null)
                        summarySheet.Cells[currentRowNumber, 4].AddComment(partPair.Part1Ancestors, "_").AutoFit = true;

                    if (partPair.Part2Process != null)
                        summarySheet.Cells[currentRowNumber, 5].AddComment(partPair.Part2Process, "_").AutoFit = true;

                    if (partPair.Part2DS != null)
                        summarySheet.Cells[currentRowNumber, 6].AddComment(partPair.Part2DS, "_").AutoFit = true;

                    if (partPair.Part2Ancestors != null)
                        summarySheet.Cells[currentRowNumber, 7].AddComment(partPair.Part2Ancestors, "_").AutoFit = true;


                    //###########
                    // Sparklines
                    //###########

                    var sortedClearances = partPair.Clearances.OrderBy(x => x.Key).Reverse();

                    var minResult = sortedClearances.Min(x => x.Value.Result);
                    var maxResult = sortedClearances.Max(x => x.Value.Result);

                    var clearancesCount = sortedClearances.Count();

                    var clearanceValuesToDistanceRange = new List<Tuple<double, double, double?>>(clearancesCount);
                    Nullable<double> previousResult = null;

                    int maxDistanceToMarriedPadding = 0;
                    int minDistanceToMarriedPadding = 0;

                    foreach (var keyValuePair in sortedClearances)
                    {
                        var clearance = keyValuePair.Value;

                        int paddingLength = clearance.DistanceToMarried.ToString().Length + /* mm*/ 3;

                        if (previousResult != clearance.Result)
                        {
                            previousResult = clearance.Result;
                            clearanceValuesToDistanceRange.Add(new Tuple<double, double, double?>(clearance.Result, clearance.DistanceToMarried, null));

                            if (paddingLength > maxDistanceToMarriedPadding)
                                maxDistanceToMarriedPadding = paddingLength;
                        }

                        else
                        {
                            var lastClearanceValuesToDistanceRangeItem = clearanceValuesToDistanceRange[clearanceValuesToDistanceRange.Count - 1];
                            clearanceValuesToDistanceRange[clearanceValuesToDistanceRange.Count - 1] = new Tuple<double, double, double?>(lastClearanceValuesToDistanceRangeItem.Item1, lastClearanceValuesToDistanceRangeItem.Item2, clearance.DistanceToMarried);

                            if (paddingLength > maxDistanceToMarriedPadding)
                                minDistanceToMarriedPadding = paddingLength;
                        }
                    }

                    var resultPadding = /*negative sign*/ (minResult < 0 ? 1 : 0)
                        + /*result value*/ Math.Max(Math.Abs(minResult), Math.Abs(maxResult)).ToString("0.00").Length
                        + /* mm*/ 3;

                    var sparklineCell = summarySheet.Cells[currentRowNumber, 8];
                    var sparklineDataRange = spaklinesDataSheet.Cells[String.Format("A{0}:A{1}", sparklinesDataRow, sparklinesDataRow + clearancesCount - 1)];

                    summarySheet.SparklineGroups.Add(eSparklineType.Line, sparklineCell, sparklineDataRange);

                    sparklineDataRange.LoadFromArrays(sortedClearances.Select(x => new object[] { x.Value.Result }).ToList());

                    var sparklineComment = sparklineCell.AddComment(string.Join("\r\n", clearanceValuesToDistanceRange.Select(x =>
                    {
                        var paddedValue = (Math.Abs(x.Item1).ToString("0.00") + " mm").PadLeft(resultPadding, ' ');

                        if (x.Item1 < 0)
                            paddedValue = "-" + paddedValue.Remove(0, 1);

                        paddedValue += " @ "
                            + (x.Item2 + " mm").PadLeft(maxDistanceToMarriedPadding, ' ')
                            + (x.Item3 != null ? (" - " + (x.Item3 + " mm").PadLeft(minDistanceToMarriedPadding, ' ')) : "");

                        return paddedValue;
                    }).ToList()), "_");
                    sparklineComment.Font.FontName = "Courier New";
                    sparklineComment.AutoFit = true;

                    sparklinesDataRow += clearancesCount;
                }

                //###############
                // END Sparklines
                //###############

                if (!mfgArea1Present) summarySheet.Column(2).Hidden = true;
                if (!mfgArea2Present) summarySheet.Column(5).Hidden = true;
                if (!cpsc1Present) summarySheet.Column(3).Hidden = true;
                if (!cpsc2Present) summarySheet.Column(6).Hidden = true;

                templatePackage.Compression = CompressionLevel.BestCompression;

                templatePackage.SaveAs(excelOutputStream);
            }
        }
    }
}
