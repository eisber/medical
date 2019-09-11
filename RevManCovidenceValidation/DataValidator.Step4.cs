using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;

namespace RevManCovidenceValidation
{
    public partial class DataValidator
    {
        private static string JoinAlternatives(ILookup<string, string> options, string key)
        {
            var list = new List<string>() { key };

            if (options.Contains(key))
                list.AddRange(options[key]);

            var sb = new StringBuilder();

            foreach (var e in list)
                sb.AppendFormat("\t\t{0}\n", e);

            return sb.ToString();
        }

        private void PrintState(string text, dynamic rev, dynamic covidenceStudyData, 
            ILookup<string, string> combinedOutcomes, ILookup<string, string> mappedComparison,
            dynamic matchedList,
            string n, string N)
        {
            Console.WriteLine(@"Unable to find {0} data
    Study           {1} 
    RevMan Outcome  {2}
    RevMan Subgroup {3}
    Matches         {4}
    Data            {5}/{6}",
    text, 
     rev.StudyId, rev.Outcome, rev.Subgroup, matchedList.Count, n, N);

            if (matchedList.Count > 0)
            {
                Console.WriteLine("Covidence matched candidates are");
                foreach (var item in matchedList)
                {
                    Console.WriteLine("\t{0} / {1}: {2}/{3}",
                        item.Comparison, item.Subgroup, item.n, item.N);
                }
            }
            else
            {
                Console.WriteLine("Covidence candidates are");
                foreach (var covData in covidenceStudyData)
                {
                    Console.WriteLine(@"
####### Outcom '{0}' not in 
{1}
        Compar '{2}' not in 
{3}
        Data {4}/{5}
        Covidence Subgroup '{6}' (should start with PNB or NO-PNB)",
                        covData.Outcome, JoinAlternatives(combinedOutcomes, rev.Outcome),
                        covData.Comparison, JoinAlternatives(mappedComparison, rev.Subgroup),
                        covData.n, covData.N, covData.Subgroup);
                }
            }

            Console.WriteLine("Press key to continue");
            // Console.ReadKey();
        }

        public void Step4_ValidateRevMan()
        {
            var combinedOutcomes = File.ReadAllLines(mappingFileOutcomes)
                    .Select(s => s.Split('\t'))
                    .Select(f => new
                    {
                        CombinedOutcome = f[0].Trim(),
                        Outcome = f[1].Trim()
                    })
                    .ToLookup(f => f.CombinedOutcome, f => f.Outcome);

            var mappedComparison = File.ReadAllLines(mappingFileComparison)
                    .Select(s => s.Split('\t'))
                    .Select(f => new
                    {
                        RevmanSubgroup = f[0].Trim(),
                        CovidenceComparison = f[1].Trim()
                    })
                    .ToLookup(f => f.RevmanSubgroup, f => f.CovidenceComparison);

            // load covidence extracted data
            var data = File.ReadAllLines(CovidenceDataPath)
                .Select(l => JsonConvert.DeserializeObject(l))
                .Cast<dynamic>()
                .Select(c => new
                {
                    RevManStudyId = (string)c.RevManStudyId.ToString(),
                    StudyName = (string)c.StudyName.ToString(),
                    Outcome = (string)c.Outcome.ToString(),
                    Comparison = (string)c.Comparison.ToString(),
                    Subgroup = (string)c.Subgroup.ToString(),
                    n = (string)c.n.ToString(),
                    N = (string)c.N.ToString()
                })
                .ToList();

            // load revman data
            var revmanData = revmanXml.XPathSelectElements("//DICH_DATA")
                .Select(n => new
                {
                    StudyId = n.Attribute("STUDY_ID").Value,
                    PNB_n = n.Attribute("EVENTS_1").Value,
                    PNB_N = n.Attribute("TOTAL_1").Value,
                    NO_PNB_n = n.Attribute("EVENTS_2").Value,
                    NO_PNB_N = n.Attribute("TOTAL_2").Value,
                    Outcome = n.XPathSelectElement("ancestor::DICH_OUTCOME/NAME").Value,
                    Subgroup = n.XPathSelectElement("ancestor::DICH_SUBGROUP/NAME").Value
                })
                .ToList();

            foreach (var revmanStudyData in revmanData.GroupBy(r => r.StudyId))
            {
                // per study
                var covidenceStudyData = data.Where(s => s.RevManStudyId == revmanStudyData.Key).ToList();

                // per datapoint
                foreach (var rev in revmanStudyData)
                {
                    var matchedEntries = covidenceStudyData
                        .Where(c =>
                            (c.Comparison == rev.Subgroup || 
                               (mappedComparison.Contains(rev.Subgroup) && mappedComparison[rev.Subgroup].Contains(c.Comparison)))
                            &&
                            // either exact match or it comes from the combined outcomes
                            (c.Outcome == rev.Outcome || 
                            // combined outcome exists?
                             (combinedOutcomes.Contains(rev.Outcome) && combinedOutcomes[rev.Outcome].Contains(c.Outcome)))
                        )
                        .ToList();

                    // TODO: need subgroup mapping too
                    Predicate<string> pnbCondition = c => c.StartsWith("PNB") || c == "Adductor Canal Catheter" || c == "FEMI/SCI";


                    var pnb = matchedEntries.Where(c => pnbCondition(c.Subgroup)).ToList();
                    var no_pnb = matchedEntries.Where(c => !pnbCondition(c.Subgroup)).ToList();

                    // search for PNB
                    if (pnb.Count == 0)
                    {
                        PrintState("PNB", rev, covidenceStudyData, combinedOutcomes, mappedComparison, pnb, rev.PNB_n, rev.PNB_N);
                    }
                    else
                    {
                        bool found = false;

                        foreach (var covData in pnb)
                        {
                            if (rev.PNB_n == covData.n && rev.PNB_N == covData.N)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                            PrintState("PNB", rev, covidenceStudyData, combinedOutcomes, mappedComparison, pnb, rev.PNB_n, rev.PNB_N);
                    }

                    // search for NO-PNB
                    if (no_pnb.Count == 0)
                        PrintState("NO-PNB", rev, covidenceStudyData, combinedOutcomes, mappedComparison, no_pnb, rev.NO_PNB_n, rev.NO_PNB_N);
                    else
                    {
                        bool found = false;
                        foreach (var item in no_pnb)
                        {
                            if (rev.NO_PNB_n == item.n && rev.NO_PNB_N == item.N)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            Console.WriteLine("Unable to find NO-PNB data {0} / {1} / {2}",
                                rev.StudyId, rev.Outcome, rev.Subgroup);

                            PrintState("NO-PNB", rev, covidenceStudyData, combinedOutcomes, mappedComparison, no_pnb, rev.NO_PNB_n, rev.NO_PNB_N);
                        }
                    }
                }

                int x = 1;
            }

            /*
            // match revman <-> covidence
            foreach (var covidenceStudyData in data.GroupBy(k => k.RevManStudyId))
            {
                var revmanStudyData = revmanData
                    .Where(s => s.StudyId == covidenceStudyData.Key)
                    .Select((r, i) =>
                    new
                    {
                        Revman = r,
                        Index = i
                    })
                    .ToList();

                var processedPNB = new HashSet<int>();
                var processedNO_PNB = new HashSet<int>();

                foreach (var cov in covidenceStudyData)
                {
                    foreach (var rev in revmanStudyData)
                    {
                        HashSet<int> processed;
                        string revManEvents;
                        string revManTotal;

                        if (cov.Subgroup == "PNB")
                        {
                            processed = processedPNB;
                            revManEvents = rev.Revman.PNB_n;
                            revManTotal = rev.Revman.PNB_N;
                        }
                        else
                        {
                            processed = processedNO_PNB;
                            revManEvents = rev.Revman.NO_PNB_n;
                            revManTotal = rev.Revman.NO_PNB_N;
                        }

                        if (processed.Contains(rev.Index))
                            continue;

                        if (cov.Outcome == rev.Revman.Outcome &&
                            cov.Comparison == rev.Revman.Subgroup)
                        {
                            if (cov.n != revManEvents)
                            {
                                Console.WriteLine("Mismatched {0} / {1} / {2} / {3}: n={4} vs {5}", 
                                    cov.RevManStudyId, cov.Outcome, cov.Comparison, cov.Subgroup,
                                    cov.n, revManEvents);
                            }

                            if (cov.N != revManTotal)
                            {
                                Console.WriteLine("Mismatched {0} / {1} / {2} / {3}: n={4} vs {5}",
                                    cov.RevManStudyId, cov.Outcome, cov.Comparison, cov.Subgroup,
                                    cov.N, revManTotal);
                            }

                            processed.Add(rev.Index);

                            continue;
                        }

                        Console.WriteLine("Unable to match covidence data {0} / {1} / {2}",
                            cov.RevManStudyId, cov.Outcome, cov.Comparison);
                    }

                    var unmatchedPNB = revmanStudyData.Where(r => !processedPNB.Contains(r.Index)).ToList();
                    var unmatchedNO_PNB = revmanStudyData.Where(r => !processedNO_PNB .Contains(r.Index)).ToList();

                    if (unmatchedPNB.Count > 0)
                    {
                        Console.WriteLine("Unable to match PNB events");
                        foreach (var item in unmatchedPNB)
                            Console.WriteLine("\t{0} / {1}",
                                item.Revman.Outcome, item.Revman.Subgroup);
                    }

                    if (unmatchedNO_PNB.Count > 0)
                    {
                        Console.WriteLine("Unable to match NO_PNB events");
                        foreach (var item in unmatchedNO_PNB)
                            Console.WriteLine("\t{0} / {1}",
                                item.Revman.Outcome, item.Revman.Subgroup);
                    }

                    int x = 1;
                }

                // match exactly
                //revmanStudyData.Intersect(cov)

            }*/

        }
    }
}
