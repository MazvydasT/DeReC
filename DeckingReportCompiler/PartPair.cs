using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace DeckingReportCompiler
{
    public class PartPair
    {
        public int Number { get; private set; }

        public string[] Part1Hierarchy { get; private set; }
        public string[] Part2Hierarchy { get; private set; }

        public string Part1Ancestors
        {
            get
            {
                return GetAncestors(Part1Hierarchy);
            }
        }

        public string Part2Ancestors
        {
            get
            {
                return GetAncestors(Part2Hierarchy);
            }
        }

        public string Part1Transforms { get; private set; }
        public string Part2Transforms { get; private set; }

        public string Part1Process
        {
            get
            {
                return GetProcess(Part1Hierarchy);
            }
        }

        public string Part2Process
        {
            get
            {
                return GetProcess(Part2Hierarchy);
            }
        }

        public string Part1MfgArea
        {
            get
            {
                var part1Process = Part1Process;

                if (part1Process != null)
                {
                    var matchGroups = mfgAreaRegex.Match(part1Process).Groups;

                    return matchGroups[1].Value + " " + matchGroups[2].Value;
                }

                return null;
            }
        }

        public string Part2MfgArea
        {
            get
            {
                var part2Process = Part2Process;

                if (part2Process != null)
                {
                    var matchGroups = mfgAreaRegex.Match(part2Process).Groups;

                    return matchGroups[1].Value + " " + matchGroups[2].Value;
                }

                return null;
            }
        }

        public string Part1DS
        {
            get
            {
                return GetDSNumber(Part1Hierarchy);
            }
        }

        public string Part2DS
        {
            get
            {
                return GetDSNumber(Part2Hierarchy);
            }
        }

        public Nullable<int> Part1CPSC
        {
            get
            {
                var part1DS = Part1DS;

                return part1DS == null ? (Nullable<int>)null : int.Parse(cpscRegex.Match(part1DS).Groups[1].Value);
            }
        }

        public Nullable<int> Part2CPSC
        {
            get
            {
                var part2DS = Part2DS;

                return part2DS == null ? (Nullable<int>)null : int.Parse(cpscRegex.Match(part2DS).Groups[1].Value);
            }
        }

        /*public double[] ClearanceValues
        {
            get
            {
                var clearanceValues = new double[Clearances.Count];
                
                var clearanceKeys = new List<int>(Clearances.Keys);
                clearanceKeys.Sort();
                clearanceKeys.Reverse();

                for (int i = 0, c = clearanceValues.Length; i < c; ++i)
                {
                    clearanceValues[i] = Clearances[clearanceKeys[i]].Result;
                }

                return clearanceValues;
            }
        }*/

        private static readonly Regex mfgAreaRegex = new Regex(@".*?(F[a-zA-Z]{2})[\s-_]{0,2}(\d{5}).*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex dsRegex = new Regex(@".*?ds.*?\d{6}.*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex cpscRegex = new Regex(@".?(\d{6}).*", RegexOptions.Compiled);

        public Dictionary<int, Clearance> Clearances { get; private set; }

        public Clearance WorstCaseClearance
        {
            get
            {
                var keys = new List<int>(Clearances.Keys);
                keys.Sort();

                Clearance worstCaseClearance = null;

                for (var i = keys.Count - 1; i > -1; --i)
                {
                    var key = keys[i];

                    var clearance = Clearances[key];

                    if (worstCaseClearance == null || clearance.Result < worstCaseClearance.Result)
                        worstCaseClearance = clearance;
                }

                return worstCaseClearance;
            }
        }

        public Clearance FirstDetectedClearance
        {
            get
            {
                var keys = new List<int>(Clearances.Keys);
                keys.Sort();
                return Clearances[keys[keys.Count - 1]];

                /*if (keys.Count > 1)
                {
                    keys.Sort();

                    return Clearances[keys[keys.Count - 1]];
                }

                else
                    return null;*/
            }
        }

        public Clearance ClearanceAtDecked
        {
            get
            {
                return Clearances.ContainsKey(0) ? Clearances[0] : null;
            }
        }

        public PartPair(string[] part1Hierarchy, string part1Transforms, string[] part2Hierarchy, string part2Transforms, int number)
        {
            Number = number;

            Clearances = new Dictionary<int, Clearance>();

            Part1Hierarchy = part1Hierarchy;
            Part2Hierarchy = part2Hierarchy;

            Part1Transforms = part1Transforms;
            Part2Transforms = part2Transforms;
        }

        private string GetProcess(string[] partHierarchy)
        {
            for (var i = partHierarchy.Length - 1; i > -1; --i)
            {
                var match = mfgAreaRegex.Match(partHierarchy[i]);

                if (match.Success)
                    return match.Value;
            }

            return null;
        }

        private string GetDSNumber(string[] partHierarchy)
        {
            for (var i = partHierarchy.Length - 1; i > -1; --i)
            {
                var match = dsRegex.Match(partHierarchy[i]);

                if (match.Success)
                    return match.Value;
            }

            return null;
        }

        private string GetAncestors(string[] partHierarchy)
        {
            var stringBuilder = new StringBuilder(partHierarchy.Length - 1);

            for (int i = 0, c = partHierarchy.Length - 1; i < c; ++i)
            {
                stringBuilder.Append(new String(' ', 2 * i));
                stringBuilder.Append(partHierarchy[i]);

                if (i + 1 < c) stringBuilder.Append("\r\n");
            }

            return stringBuilder.Length == 0 ? null : stringBuilder.ToString();
        }
    }
}
