using F23.StringSimilarity;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;

namespace RevManCovidenceValidation
{
    public partial class DataValidator
    {
        private List<Study> studies;

        private List<Study> CovidenceParseMasterList()
        {
            var covidenceMasterWorksheet = ((Worksheet)covidenceExcel.Worksheets["Studies"]);

            var masterList = new List<Study>();
            for (int row = 1; row < 8 * 1024; row++)
            {
                var studyName = covidenceMasterWorksheet.Cells[row, 1].Value;
                var studyTitle = covidenceMasterWorksheet.Cells[row, 2].Value;

                if (string.IsNullOrWhiteSpace(studyName) && string.IsNullOrWhiteSpace(studyTitle))
                    break;

                masterList.Add(new Study
                {
                    Name = studyName,
                    Title = studyTitle,
                    WorksheetIndex = row + 1
                });
            }

            return masterList;
        }

        private static string GetCovidenceYear(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            return name.Substring(name.Length - 4);
        }

        public void Step2_MatchRevmanAndCovidenceStudies()
        {
            var studyIdsWithData = new HashSet<string>(
                revmanXml.XPathSelectElements("//*[@STUDY_ID]")
                   .Where(d => d.Name == "DICH_DATA" || d.Name == "CONT_DATA")
                   .Select(d => d.Attribute("STUDY_ID").Value)
                   .Distinct());

            // correlate studies
            var revmanStudies = revmanXml.XPathSelectElements("//STUDY")
                .Select(n => new
                {
                    Id = n.Attribute("ID").Value,
                    Year = n.Attribute("YEAR").Value,
                    Name = n.Attribute("NAME").Value,
                    Title = n.XPathSelectElement(".//TI").Value
                })
                .Where(n => studyIdsWithData.Contains(n.Id))
                .ToList();

            studies = CovidenceParseMasterList();
            var masterSearchList = new List<Study>(studies);

            var matchedStudies = new HashSet<string>();

            foreach (var revmanStudy in revmanStudies)
            {
                var matches = masterSearchList
                    .Where(s => s.Name == revmanStudy.Name)
                    .ToList();

                if (matches.Count == 1)
                {
                    matches.First().RevManStudyId = revmanStudy.Id;

                    matchedStudies.Add(revmanStudy.Id);
                    masterSearchList.Remove(matches.First());
                }
            }

            var textDistance = new Levenshtein();

            foreach (var revmanStudy in revmanStudies)
            {
                if (matchedStudies.Contains(revmanStudy.Id))
                    continue;

                var matches = masterSearchList.Select(s => new
                {
                    Study = s,
                    DistanceTitle = textDistance.Distance(revmanStudy.Title, s.Title),
                    DistanceName = s.Name == null ? int.MaxValue : textDistance.Distance(revmanStudy.Name, s.Name),
                    Year = GetCovidenceYear(s.Name)
                })
                .ToList();

                var orderedMatches = matches
                    .Where(s => s.Year == null || s.Year == revmanStudy.Year)
                    .OrderBy(s => s.DistanceTitle).ThenBy(s => s.DistanceName)
                    .ToList();

                var firstMatch = orderedMatches.First();

                if (firstMatch.DistanceName + firstMatch.DistanceTitle < 5)
                {
                    firstMatch.Study.RevManStudyId = revmanStudy.Id;

                    matchedStudies.Add(revmanStudy.Id);
                    masterSearchList.Remove(firstMatch.Study);

                    Console.WriteLine(@"
Matching by title/name RevMan/Covidence

RevMan Study:    {0}
Covidence Study: {1}
RevMan Title:    {2}
Covidence Title: {3}
",
        revmanStudy.Name,
        firstMatch.Study.Name,
        revmanStudy.Title,
        firstMatch.Study.Title);

                    continue;
                }

                var ignoreYearMatches = matches
                    .OrderBy(s => s.DistanceName).ThenBy(s => s.DistanceTitle)
                    .Take(5)
                    .ToList();

                Console.WriteLine(@"
WARNING --- UNABLE TO MATCH

RevMan Study: {0}
RevMan Title: {1}
Matches:      {2}
{3}
",
    revmanStudy.Name,
    revmanStudy.Title,
    matches.Count,
    string.Join("\n\t", ignoreYearMatches));
            }

            Console.WriteLine(@"
Matched studies:   {0}
RevMan studies:    {1}
Covidence studies: {2}",
    studies.Count(x => x.RevManStudyId != null),
    revmanStudies.Count,
    studies.Count);


            // filter the ones we didn't match (should be the excluded ones)
            studies = studies.Where(s => s.RevManStudyId != null).ToList();
        }
    }
}
