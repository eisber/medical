using F23.StringSimilarity;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;

namespace RevManCovidenceValidation
{
    static class RevManSummaryToCSV
    {
        public static void Process()
        {
            using (var csv = new StreamWriter(@"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20201015 summary.csv"))
            {
                csv.WriteLine("outcome,subgroup,n_studies,OR,CI_lower,CI_upper,I2,heterogeneity_p_value,OR_p_value,events1,total1,events2,total2,endnote,effect_measure");

                foreach (var effect_measure in new[] { "OR", "RR" })
                    ProcessFile(csv, $@"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20201015 all {effect_measure}.rm5", true, effect_measure);

                // ProcessFile(csv, @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 THA TKA.rm5", true);
                // ProcessFile(csv, @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 Infiltration.rm5", false);
            }              
        }

        public static void ExtractRiskOfBias()
        {
            var revmanInput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 clean.rm5";

            var revmanXml = XDocument.Parse(File.ReadAllText(revmanInput));

            using (var csv = new StreamWriter(@"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 risk of bias.csv"))
            {
                csv.WriteLine("Study,RiskName,Result");

                foreach (var riskOfBias in
                    revmanXml.XPathSelectElements("//QUALITY_ITEM").SelectMany(
                        qualityItem =>
                            qualityItem.XPathSelectElements(".//QUALITY_ITEM_DATA_ENTRY")
                                .Select(e => new
                                {
                                    Name = qualityItem.XPathSelectElement("./NAME").Value,
                                    StudyId = Regex.Match(e.Attribute("STUDY_ID").Value, "-([^-]+-[0-9a-z]+$)").Groups[1].Value.Replace("-", " "),
                                    Result = e.Attribute("RESULT").Value
                                })))
                {
                    csv.WriteLine($"{riskOfBias.Name},{riskOfBias.StudyId},{riskOfBias.Result}");
                }
            }
        }

        public static void ExtractOutcomes()
        {
            var revmanInput = @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 clean.rm5";

            var revmanXml = XDocument.Parse(File.ReadAllText(revmanInput));

            using (var csv = new StreamWriter(@"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20191102 outcomes.csv"))
            {
                csv.WriteLine("Study,Outcome");

                foreach (var type in new []{ "DICH" }) // don't care for continuious
                {
                    foreach (var outcome in
                        revmanXml.XPathSelectElements($"//{type}_OUTCOME").SelectMany(
                            outcome =>
                                outcome.XPathSelectElements($".//{type}_DATA")
                                    .Select(e => new
                                    {
                                        Name = outcome.XPathSelectElement("./NAME").Value,
                                        StudyId = Regex.Match(e.Attribute("STUDY_ID").Value, "-([^-]+-[0-9a-z]+$)").Groups[1].Value.Replace("-", " "),
                                    })))
                    {
                        csv.WriteLine($"{outcome.StudyId},{outcome.Name}");
                    }
                }
            }
        }

        public class EndNoteEntry
        {
            public string RecNumber;

            public List<string> Titles;

            public string Year;

            public List<string> Authors;
        }

        public static List<EndNoteEntry> GetEndNoteRecords()
        {
            var endnoteExportXml = XDocument.Parse(File.ReadAllText(@"C:\work\medical\RevManCovidenceValidation\Cris EndNote Library.xml"));

            return endnoteExportXml.XPathSelectElements("//record").Select(
                record =>
                {
                    var recNumber = record.XPathSelectElement("./rec-number")?.Value;
                    if (recNumber == null)
                        return null;

                    var titles = record.XPathSelectElements("./titles/descendant::*")
                        .Select(e => e.Value)
                        .Distinct()
                        .ToList();

                    var authors = record.XPathSelectElements(".//author/descendant::*")
                        .Select(e => e.Value)
                        .SelectMany(auth => auth.Split(','))
                        .Distinct()
                        .ToList();

                    return new EndNoteEntry
                    {
                        RecNumber = recNumber,
                        Titles = titles,
                        Year = record.XPathSelectElement("./dates/year/descendant::*")?.Value,
                        Authors = authors
                    };
                })
                .Where(x => x != null)
                .ToList();
        }

        private static void ProcessFile(StreamWriter csv, string revmanInput, bool total, string effect_measure)
        {
            var revmanXml = XDocument.Parse(File.ReadAllText(revmanInput));

            // find total number of studies
            {
                var studies =
                    revmanXml.XPathSelectElements("//DICH_DATA").Select(e => e.Attribute("STUDY_ID").Value).Union(
                        revmanXml.XPathSelectElements("//CONT_DATA").Select(e => e.Attribute("STUDY_ID").Value))
                    .Distinct();
                Console.WriteLine("Studies: " + studies.Count());

                var observational = revmanXml.XPathSelectElements("//QUALITY_ITEM")
                    .Where(qualityItem => qualityItem.XPathSelectElement("./NAME").Value ==
                        "Observational Study: Failure to develop and apply appropriate eligibility criteria")
                    .SelectMany(qualityItem =>
                        qualityItem.XPathSelectElements(".//QUALITY_ITEM_DATA_ENTRY")
                            .Select(e => e.Attribute("STUDY_ID").Value))
                    .ToHashSet();

                foreach(var g in studies.GroupBy(s => observational.Contains(s)))
                    Console.WriteLine((g.Key ? "Observational": "RCT") + ":" + g.Count());
            }

            var recordsText = GetEndNoteRecords();
            ProcessTypeSingleFile(csv, revmanXml, "DICH", recordsText, effect_measure);
            ProcessTypeSingleFile(csv, revmanXml, "CONT", recordsText, effect_measure);
        }

        private static void ProcessType(StreamWriter csv, XDocument revmanXml, string type, bool total)
        {
            foreach (var outcome in revmanXml.XPathSelectElements($"//{type}_OUTCOME"))
            {
                var outcomeName = outcome.XPathSelectElement("./NAME").Value;

                foreach (var subgroup in outcome.XPathSelectElements($"./{type}_SUBGROUP"))
                {
                    var subgroupName = subgroup.XPathSelectElement("./NAME").Value;
                    var studies = subgroup.XPathSelectElements($"./{type}_DATA")
                        .Select(e => e.Attribute("STUDY_ID").Value)
                        .Distinct()
                        .Count();
                    var or = subgroup.Attribute("EFFECT_SIZE").Value;
                    var ciLower = subgroup.Attribute("CI_START").Value;
                    var ciUpper = subgroup.Attribute("CI_END").Value;
                    var i2 = subgroup.Attribute("I2").Value;
                    var heterogeneity_p_value = subgroup.Attribute("P_CHI2").Value;
                    var OR_p_value = subgroup.Attribute("P_Z").Value;

                    csv.WriteLine($"{outcomeName},{subgroupName},{studies},{or},{ciLower},{ciUpper},{i2},{heterogeneity_p_value},{OR_p_value}");
                }

                if (total)
                {
                    var studies = outcome.XPathSelectElements($".//{type}_DATA")
                        .Select(e => e.Attribute("STUDY_ID").Value)
                        .Distinct()
                        .Count();
                    var or = outcome.Attribute("EFFECT_SIZE").Value;
                    var ciLower = outcome.Attribute("CI_START").Value;
                    var ciUpper = outcome.Attribute("CI_END").Value;
                    var i2 = outcome.Attribute("I2").Value;
                    var heterogeneity_p_value = outcome.Attribute("P_CHI2").Value;
                    var OR_p_value = outcome.Attribute("P_Z").Value;
                    csv.WriteLine($"{outcomeName},total,{studies},{or},{ciLower},{ciUpper},{i2},{heterogeneity_p_value},{OR_p_value}");
                }
            }
        }

        static Levenshtein textDistance = new Levenshtein();

        public static HashSet<string> notFoundStudy = new HashSet<string>();
        public static HashSet<string> foundStudy = new HashSet<string>();

        private static string NormalizeTitle(string title)
        {
            return Regex.Replace(title.Trim().ToLowerInvariant(), "[^a-zA-Z0-9 ]", string.Empty);
        }

        public static string GetEndNoteStudies(List<EndNoteEntry> endnoteRecords, XDocument revmanXml, String studyId)
        {
            var studyElement = revmanXml.XPathSelectElement($"//STUDY[@ID='{studyId}']");

            var title = studyElement.XPathSelectElement(".//TI")?.Value?.Trim();
            var year = studyElement.Attribute("YEAR").Value; //.XPathSelectElement(".//YR").Value;
            var authors = studyElement.XPathSelectElement(".//AU")?.Value;
            var author = studyId.Split('-')[1];

            var authorMatches = endnoteRecords
                // .Where(e => e.Authors.Any(a => a == author))
                .ToList();

            var matches = authorMatches
                .Where(e => e.Year == year)
                .SelectMany(e => e.Titles
                    .Where(t => title == null || NormalizeTitle(title) == NormalizeTitle(t))
                    .Select(t =>
                        new
                        {
                            e,
                            title = t,
                        }))
                        .ToList();

            if (matches.Count == 0)
            {
                var approxMatch = endnoteRecords
                    .Where(e => e.Year == year)
                    .SelectMany(e => e.Titles
                        .Select(t =>
                            new
                            {
                                e,
                                title = t,
                                similiarty = title == null ? 1 : textDistance.Distance(title, t)
                            }))
                            .OrderBy(e => e.similiarty)
                            .ToList();

                if (!notFoundStudy.Contains(studyId)) {
                    notFoundStudy.Add(studyId);
                    Console.WriteLine($"NOT FOUND {studyId} {year} vs {approxMatch.First().e.Year}");
                    Console.WriteLine($"\tRevMan:  {authors}");
                    var endNoteAuthors = string.Join(", ", approxMatch.First().e.Authors);
                    Console.WriteLine($"\tEndNote: {endNoteAuthors}");
                    Console.WriteLine($"\tRevMan:  {title}");
                    Console.WriteLine($"\tEndNote: {approxMatch.First().title}");
                }
                return null;
            }

            foundStudy.Add(studyId);

            var matchedStudyAuthors = string.Join(", ", matches.First().e.Authors);
            //Console.WriteLine($"{studyId} <-> {matchedStudyAuthors}");
            //Console.WriteLine($"\tRevMan  {title}");
            //Console.WriteLine($"\tEndNote {matches.First().e.Titles.First()}");
            //Console.WriteLine();

            return "{#" + matches.First().e.RecNumber + "}";
        }

        private static void ProcessTypeSingleFile(StreamWriter csv, XDocument revmanXml, string type, List<EndNoteEntry> endnoteRecords, string effect_measure)
        {
            foreach (var outcome in revmanXml.XPathSelectElements($"//{type}_OUTCOME"))
            {
                var outcomeName = outcome.XPathSelectElement("./NAME").Value;

                // cut subgroup...
                var sepIdx = outcomeName.IndexOf(" | ");
                var totalPostfix = "";
                if (sepIdx > 0)
                {
                    totalPostfix = " " + outcomeName.Substring(sepIdx + 3);
                    outcomeName = outcomeName.Substring(0, sepIdx);
                }

                var subgroups = outcome.XPathSelectElements($"./{type}_SUBGROUP").ToList();

                int events1, events2, total1, total2;
                foreach (var subgroup in subgroups)
                {
                    events1 = events2 = total1 = total2 = 0;

                    var subgroupName = subgroup.XPathSelectElement("./NAME").Value;

                    var studyElements = subgroup.XPathSelectElements($"./{type}_DATA")
                        .Select(e => e.Attribute("STUDY_ID").Value)
                        .Distinct()
                        .ToList();
                    var studies = studyElements
                        .Count();
                    var or = subgroup.Attribute("EFFECT_SIZE").Value;
                    var ciLower = subgroup.Attribute("CI_START").Value;
                    var ciUpper = subgroup.Attribute("CI_END").Value;
                    var i2 = subgroup.Attribute("I2").Value;
                    var heterogeneity_p_value = subgroup.Attribute("P_CHI2").Value;
                    var OR_p_value = subgroup.Attribute("P_Z").Value;

                    if (type == "DICH")
                    {
                        events1 = int.Parse(subgroup.Attribute("EVENTS_1").Value);
                        events2 = int.Parse(subgroup.Attribute("EVENTS_2").Value);
                    }

                    total1 = int.Parse(subgroup.Attribute("TOTAL_1").Value);
                    total2 = int.Parse(subgroup.Attribute("TOTAL_2").Value);

                    var endNoteStudies = string.Join(" ", studyElements.Select(s => GetEndNoteStudies(endnoteRecords, revmanXml, s)));

                    csv.WriteLine($"{outcomeName},{subgroupName},{studies},{or},{ciLower},{ciUpper},{i2},{heterogeneity_p_value},{OR_p_value},{events1},{total1},{events2},{total2},{endNoteStudies},{effect_measure}");
                }

                events1 = events2 = total1 = total2 = 0;

                // if (subgroups.Count == 0)
                {
                    var studyElements = outcome.XPathSelectElements($".//{type}_DATA")
                        .Select(e => e.Attribute("STUDY_ID").Value)
                        .Distinct();
                    var studies = studyElements
                        .Count();
                    var or = outcome.Attribute("EFFECT_SIZE").Value;
                    var ciLower = outcome.Attribute("CI_START").Value;
                    var ciUpper = outcome.Attribute("CI_END").Value;
                    var i2 = outcome.Attribute("I2").Value;
                    var heterogeneity_p_value = outcome.Attribute("P_CHI2").Value;
                    var OR_p_value = outcome.Attribute("P_Z").Value;

                    if (type == "DICH")
                    {
                         events1 = int.Parse(outcome.Attribute("EVENTS_1").Value);
                         events2 = int.Parse(outcome.Attribute("EVENTS_2").Value);
                    }

                    total1 = int.Parse(outcome.Attribute("TOTAL_1").Value);
                    total2 = int.Parse(outcome.Attribute("TOTAL_2").Value);

                    var endNoteStudies = string.Join(" ", studyElements.Select(s => GetEndNoteStudies(endnoteRecords, revmanXml, s)));

                    csv.WriteLine($"{outcomeName},Total{totalPostfix},{studies},{or},{ciLower},{ciUpper},{i2},{heterogeneity_p_value},{OR_p_value},{events1},{total1},{events2},{total2},{endNoteStudies},{effect_measure}");
                }
            }
        }
    }
}
