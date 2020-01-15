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
    static class RevManSummaryToCSV
    {
        public static void Process()
        {
            using (var csv = new StreamWriter(@"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20200115 summary.csv"))
            {
                csv.WriteLine("outcome,subgroup,n_studies,OR,CI_lower,CI_upper,I2,heterogeneity_p_value,OR_p_value,events1,total1,events2,total2");

                ProcessFile(csv, @"C:\Users\marcozo\OneDrive\Cris\201909 HSS Consensus PNB\PNB versus no PNB 20200115 all.rm5", true);

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

        private static void ProcessFile(StreamWriter csv, string revmanInput, bool total)
        {
            var revmanXml = XDocument.Parse(File.ReadAllText(revmanInput));

            ProcessTypeSingleFile(csv, revmanXml, "DICH");
            ProcessTypeSingleFile(csv, revmanXml, "CONT");
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

        private static void ProcessTypeSingleFile(StreamWriter csv, XDocument revmanXml, string type)
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

                    if (type == "DICH")
                    {
                        events1 = int.Parse(subgroup.Attribute("EVENTS_1").Value);
                        events2 = int.Parse(subgroup.Attribute("EVENTS_2").Value);
                    }

                    total1 = int.Parse(subgroup.Attribute("TOTAL_1").Value);
                    total2 = int.Parse(subgroup.Attribute("TOTAL_2").Value);
                    csv.WriteLine($"{outcomeName},{subgroupName},{studies},{or},{ciLower},{ciUpper},{i2},{heterogeneity_p_value},{OR_p_value},{events1},{total1},{events2},{total2}");
                }

                events1 = events2 = total1 = total2 = 0;

                // if (subgroups.Count == 0)
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

                    if (type == "DICH")
                    {
                         events1 = int.Parse(outcome.Attribute("EVENTS_1").Value);
                         events2 = int.Parse(outcome.Attribute("EVENTS_2").Value);
                    }

                    total1 = int.Parse(outcome.Attribute("TOTAL_1").Value);
                    total2 = int.Parse(outcome.Attribute("TOTAL_2").Value);

                    csv.WriteLine($"{outcomeName},Total{totalPostfix},{studies},{or},{ciLower},{ciUpper},{i2},{heterogeneity_p_value},{OR_p_value},{events1},{total1},{events2},{total2}");
                }
            }
        }
    }
}
