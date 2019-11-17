using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace RevManCovidenceValidation
{
    static class EnrichedRevManToSubgroup
    {
        private static IEnumerable<dynamic> ParseThaTkaData()
        {
            var dataInput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB study THA TKA data.csv";

            //Outcome,PNB Events, PNB Total,NO - PNB Events,NO - PNB Totals,Study,Hip,Knee
            //Arrhythmias(Not specified),0,35,2,36,Baranovi 2011,,1
            //Aspiration,0,20,1,20,Andersen 2012,,1
            using (var parser = new CsvParser(File.OpenText(dataInput)))
            {
                string[] f;

                // skip header
                parser.Read();

                while ((f = parser.Read()) != null)
                {
                    yield return new
                    {
                        Outcome = f[0],
                        PNB = int.Parse(f[1]),
                        PNB_Total = int.Parse(f[2]),
                        NO_PNB = int.Parse(f[3]),
                        NO_PNB_Total = int.Parse(f[4]),
                        Study = f[5].Replace(' ', '-'),
                        Hip = !string.IsNullOrEmpty(f[6]),
                        Knee = !string.IsNullOrEmpty(f[7])
                    };
                }
            }
        }
 
        public static void Process_THA_TKA()
        {
            var thaTkaData = ParseThaTkaData().ToList();

            var revmanInput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 clean.rm5";
            var revmanOutput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191115 all.rm5";
            var revmanXml = XDocument.Parse(File.ReadAllText(revmanInput));

            var subgroup = new List<Tuple<string, string>>();

            var observational = revmanXml.XPathSelectElements("//QUALITY_ITEM")
                .Where(qualityItem => qualityItem.XPathSelectElement("./NAME").Value ==
                    "Observational Study: Failure to develop and apply appropriate eligibility criteria")
                .SelectMany(qualityItem =>
                    qualityItem.XPathSelectElements(".//QUALITY_ITEM_DATA_ENTRY")
                        .Select(e => e.Attribute("STUDY_ID").Value))
                .ToHashSet();

            foreach (var study in revmanXml.XPathSelectElements("//STUDY"))
            {
                var id = study.Attribute("ID").Value;
                if (!(id.StartsWith("STD-H-") || id.StartsWith("STD-K-")))
                    continue;

                var nameOnly = Regex.Match(id, "-([^-]+-[0-9a-z]+$)").Groups[1].Value;
                var newId = "STD-" + nameOnly;

                var postfix = " | " + (observational.Contains(id) ? "Observational" : "RCT");

                if (id.Contains("-H-"))
                    subgroup.Add(Tuple.Create(newId, "THA" + postfix));
                if (id.Contains("-K-"))
                    subgroup.Add(Tuple.Create(newId, "TKA" + postfix));

                // H K Type multiple Comp CNB SA or GA vs SA or GA vs X Memtsoudis 2016

                var match = Regex.Match(id, "Comp-(.*)-vs-(.*)-vs-(.*)-");
                var left = match.Groups[1].Value;
                var right = match.Groups[2].Value + match.Groups[3].Value;

                //AddSubgroup(outcome, data, "PNB vs no-PNB", none);
                //AddSubgroup(outcome, data, "PNB vs infiltration", rightOnly);
                //AddSubgroup(outcome, data, "PNB + infiltration vs no-infiltration", leftOnly);
                //AddSubgroup(outcome, data, "PNB + infiltration vs infiltration", both);

                if (Regex.Match(left, "(LIA|PAI)").Success && Regex.Match(right, "(LIA|PAI)").Success)
                    subgroup.Add(Tuple.Create(newId, "PNB + infiltration vs infiltration" + postfix));
                else
                {
                    if (Regex.Match(left, "(LIA|PAI)").Success)
                        subgroup.Add(Tuple.Create(newId, "PNB + infiltration vs no-infiltration" + postfix));
                    else
                    {
                        if (Regex.Match(right, "(LIA|PAI)").Success)
                            subgroup.Add(Tuple.Create(newId, "PNB vs infiltration" + postfix));
                        else
                            subgroup.Add(Tuple.Create(newId, "PNB vs no-infiltration" + postfix));
                    }
                }

                study.SetAttributeValue("ID", newId);
                study.SetAttributeValue("NAME", newId.Replace("-", " ").Substring(4));

                // patch studyId
                foreach (var studyId in revmanXml.XPathSelectElements($"//*[@STUDY_ID = '{id}']"))
                    studyId.SetAttributeValue("STUDY_ID", newId);
            }

            int outcomeCount = 128; // revmanXml.XPathSelectElements($"//DICH_OUTCOME").Count() +
                //revmanXml.XPathSelectElements($"//CONT_OUTCOME").Count()
                //+ 1;

            // introduce subgroups
            foreach (var type in new[] { "DICH", "CONT" })
            {
                var newOutcomes = new List<XElement>();

                // All, THA, TKA, Infiltration, ...
                foreach (var group in new []
                {
                    new
                    {
                        Name = "THA/TKA",
                        Prefixes = new [] {
                            "THA",
                            "TKA"
                        }
                    },
                    new
                    {
                        Name = "Infiltration",
                        Prefixes = new [] {
                            "PNB + infiltration vs infiltration",
                            "PNB + infiltration vs no-infiltration",
                            "PNB vs infiltration",
                            "PNB vs no-infiltration"
                        }
                    }
                })
                {
                    // THA + TKA, THA, TK,...
                    //foreach (var expandedGroup in new[] { group }.Union(group.Prefixes.Select(p => new { Name = p, Prefixes = new[] { p } })))
                    foreach (var expandedGroup in group.Prefixes.Select(p => new { Name = p, Prefixes = new[] { p } }))
                    {
                        foreach (var cleanOutcome in revmanXml.XPathSelectElements($"//{type}_OUTCOME"))
                        {
                            var outcome = new XElement(cleanOutcome);
                            outcomeCount++;
                            outcome.Attribute("ID").Value = $"CMP-001.{outcomeCount:D2}";

                            var data = outcome.XPathSelectElements($"{type}_DATA").ToList();

                            var outcomeName = outcome.XPathSelectElement("./NAME").Value;
                            outcome.XPathSelectElement("./NAME").Value = outcomeName + " | " + expandedGroup.Name;

                            if (!(expandedGroup.Name.Contains("THA") || expandedGroup.Name.Contains("TKA")))
                                outcome.XPathSelectElement("./NAME").Value += " (THA/TKA)";

                            if (data.Count == 0)
                                continue;

                            foreach (var d in data)
                                d.Remove();

                            foreach (var g in subgroup.GroupBy(t => t.Item2))
                            {
                                if (!expandedGroup.Prefixes.Any(prefix => g.Key.StartsWith(prefix)))
                                    continue;

                                if (expandedGroup.Name.Contains("THA") || expandedGroup.Name.Contains("TKA"))
                                {
                                    var filteredData = data.Where(e =>
                                    {
                                        if (type == "CONT")
                                            return true;

                                        var events1 = int.Parse(e.Attribute("EVENTS_1").Value);
                                        var events2 = int.Parse(e.Attribute("EVENTS_2").Value);
                                        var total1 = int.Parse(e.Attribute("TOTAL_1").Value);
                                        var total2 = int.Parse(e.Attribute("TOTAL_2").Value);
                                        var s1 = thaTkaData.Where(a => e.Attribute("STUDY_ID").Value.Contains(a.Study)).ToList();
                                        var s2 = s1.Where(a => a.Outcome == outcomeName).ToList();
                                        var s3 = s2.Where(a => (g.Key.Contains("THA") && a.Hip) || 
                                                               (g.Key.Contains("TKA") && a.Knee))
                                                                .ToList();
                                        var s4 = s3.Where(a => a.PNB == events1 && a.PNB_Total == total1 &&
                                                               a.NO_PNB == events2 && a.NO_PNB_Total == total2).ToList();

                                        var matches = s4.Count();
                                        if (matches > 1)
                                            throw new Exception("Multiple matches");

                                        return matches > 0;
                                    }).ToList();

                                    AddSubgroup(outcome, filteredData, g.Key, g.Select(t => t.Item1).ToList());
                                }
                                else
                                {
                                    AddSubgroup(outcome, data, g.Key, g.Select(t => t.Item1).ToList());
                                }
                            }

                            // don't add if we don't have any data
                            if (outcome.XPathSelectElements($".//{type}_DATA").Count() > 0)
                                newOutcomes.Add(outcome);

                            //AddSubgroup(outcome, data, "THA", tha);
                            //AddSubgroup(outcome, data, "TKA", tka);

                            //CheckStudyData(data, outcome);
                        }
                    }
                }

                foreach (var newOutcome in newOutcomes)
                    revmanXml.XPathSelectElements($"//{type}_OUTCOME").First().AddAfterSelf(newOutcome);
            }

            revmanXml.Save(revmanOutput);
        }

        private static void CheckStudyData(List<XElement> before, XElement outcome)
        {
            var type = outcome.Name.LocalName.Equals("DICH_OUTCOME") ? "DICH" : "CONT";

            var studiesBefore = before.Select(e => e.Attribute("STUDY_ID").Value).Distinct().ToHashSet();
            var studiesAfter = outcome.XPathSelectElements($".//{type}_DATA").Select(e => e.Attribute("STUDY_ID").Value).Distinct().ToHashSet();

            studiesAfter.SymmetricExceptWith(studiesBefore);

            if (studiesAfter.Count > 0)
                throw new Exception("dropped studies " + studiesAfter);
        }

        public static void Process_Infiltration()
        {
            var revmanInput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 clean.rm5";
            var revmanOutput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 Infiltration.rm5";
            var revmanXml = XDocument.Parse(File.ReadAllText(revmanInput));

            var both = new List<string>();
            var leftOnly = new List<string>();
            var rightOnly = new List<string>();
            var none = new List<string>();

            foreach (var study in revmanXml.XPathSelectElements("//STUDY"))
            {
                var id = study.Attribute("ID").Value;
                if (!(id.StartsWith("STD-H-") || id.StartsWith("STD-K-")))
                    continue;

                var nameOnly = Regex.Match(id, "-([^-]+-[0-9a-z]+$)").Groups[1].Value;
                var newId = "STD-" + nameOnly;

                // H K Type multiple Comp CNB SA or GA vs SA or GA vs X Memtsoudis 2016

                var match = Regex.Match(id, "Comp-(.*)-vs-(.*)-vs-(.*)-");
                var left = match.Groups[1].Value;
                var right = match.Groups[2].Value + match.Groups[3].Value;

                if (Regex.Match(left, "(LIA|PAI)").Success && Regex.Match(right, "(LIA|PAI)").Success)
                    both.Add(newId);
                else
                {
                    if (Regex.Match(left, "(LIA|PAI)").Success)
                        leftOnly.Add(newId);
                    else
                    {
                        if (Regex.Match(right, "(LIA|PAI)").Success)
                            rightOnly.Add(newId);
                        else
                            none.Add(newId);
                    }
                }

                study.SetAttributeValue("ID", newId);
                study.SetAttributeValue("NAME", newId.Replace("-", " ").Substring(4));

                // patch studyId
                foreach (var studyId in revmanXml.XPathSelectElements($"//*[@STUDY_ID = '{id}']"))
                    studyId.SetAttributeValue("STUDY_ID", newId);
            }

            // introduce subgroups
            foreach (var type in new[] { "DICH", "CONT" })
            {
                foreach (var outcome in revmanXml.XPathSelectElements($"//{type}_OUTCOME"))
                {
                    var data = outcome.XPathSelectElements($"{type}_DATA").ToList();
                    if (data.Count == 0)
                        continue;

                    foreach (var d in data)
                        d.Remove();

                    AddSubgroup(outcome, data, "PNB vs no-PNB", none);
                    AddSubgroup(outcome, data, "PNB vs infiltration", rightOnly);
                    AddSubgroup(outcome, data, "PNB + infiltration vs no-infiltration", leftOnly);
                    AddSubgroup(outcome, data, "PNB + infiltration vs infiltration", both);

                    CheckStudyData(data, outcome);
                }
            }

            revmanXml.Save(revmanOutput);
        }

        private static void AddSubgroup(XElement outcome, IEnumerable<XElement> data, string name, List<string> studyIds)
        {
            outcome.SetAttributeValue("SUBGROUPS", "YES");

            var type = outcome.Name.LocalName.Equals("DICH_OUTCOME") ? "DICH" : "CONT";

            var outcomeId = outcome.Attribute("ID").Value;
            var subgroupCount = outcome.XPathSelectElements($"{type}_SUBGROUP").Count() + 1; 
            var subgroupId = $"{outcomeId}.{subgroupCount:D2}";

            var dataCount = data.Where(e => studyIds.Contains(e.Attribute("STUDY_ID").Value)).Count();

            if (dataCount > 0)
            {
                var subgroup = type == "DICH" ? 
                    XElement.Parse($"<DICH_SUBGROUP CHI2=\"0.0\" CI_END=\"0.0\" CI_START=\"0.0\" DF=\"0\" EFFECT_SIZE=\"0.0\" ESTIMABLE=\"NO\" EVENTS_1=\"0\" EVENTS_2=\"0\" I2=\"0.0\" ID=\"{subgroupId}\" LOG_CI_END=\"-Infinity\" LOG_CI_START=\"-Infinity\" LOG_EFFECT_SIZE=\"-Infinity\" MODIFIED=\"2019-11-02 18:56:37 -0400\" MODIFIED_BY=\"[Empty name]\" NO=\"2\" P_CHI2=\"0.0\" P_Z=\"0.0\" STUDIES=\"{dataCount}\" TAU2=\"0.0\" TOTAL_1=\"0\" TOTAL_2=\"0\" WEIGHT=\"0.0\" Z=\"0.0\"><NAME>{name}</NAME></DICH_SUBGROUP>") :
                    XElement.Parse($"<CONT_SUBGROUP CHI2=\"0.0\" CI_END=\"0.0\" CI_START=\"0.0\" DF=\"0\" EFFECT_SIZE=\"0.0\" ESTIMABLE=\"NO\" I2=\"0.0\" ID=\"{subgroupId}\" MODIFIED=\"2019-11-02 21:31:43 -0400\" MODIFIED_BY=\"[Empty name]\" NO=\"2\" P_CHI2=\"0.0\" P_Z=\"0.0\" STUDIES=\"{dataCount}\" TAU2=\"0.0\" TOTAL_1=\"0\" TOTAL_2=\"0\" WEIGHT=\"0.0\" Z=\"0.0\"><NAME>{name}</NAME></CONT_SUBGROUP>");


                foreach (var d in data.Where(e => studyIds.Contains(e.Attribute("STUDY_ID").Value)))
                    subgroup.Add(new XElement(d));

                outcome.Add(subgroup);
            }
        }
    }
}
