using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevManCovidenceValidation
{
    public partial class DataValidator
    {
        private int FindRow(Worksheet sheet, int start, string value)
        {
            for (int row = start; row < 1024; row++)
            {
                Range r = sheet.Cells[row, 1];

                if (r.Value?.ToString()?.Trim() == value)
                    return row;
            }

            throw new NotSupportedException("Unable to find value: '" + value + "'");
        }

        public void Step3_EnrichOutcomes()
        {
            using (var outputFile = new StreamWriter(CovidenceDataPath))
            {
                foreach (var study in studies)
                {
                    var sheet = (Worksheet)covidenceExcel.Worksheets[study.WorksheetIndex];

                    Console.WriteLine("Processing: {0}", study.Name);

                    int outcomesRow = FindRow(sheet, 15, "Outcomes");
                    int studyIdentificationRow = FindRow(sheet, outcomesRow + 1, "Study Identification");

                    // Outcome header
                    // blank | comparison
                    // blank | n | N
                    // subgroup | n-value | N-value

                    for (int row = outcomesRow + 1; row < studyIdentificationRow - 4; row++)
                    {
                        string outcome = sheet.Cells[row, 1].Value?.ToString()?.Trim();
                        row++;

                        var innerRowEnd = row;
                        for (int column = 0; column < 16; column += 2)
                        {
                            var innerRow = row;

                            string comparison = sheet.Cells[innerRow, column + 2].Value?.ToString()?.Trim();
                            if (string.IsNullOrWhiteSpace(comparison))
                                break;

                            innerRow++;

                            string header1 = sheet.Cells[innerRow, column + 2].Value?.ToString()?.Trim();

                            if (header1 == "n")
                            {
                                // dico
                                string header2 = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                if (header2 != "N")
                                    throw new NotSupportedException("Unsupported header: " + header2);
                                innerRow++;

                                for (; innerRow < studyIdentificationRow - 4; innerRow++)
                                {
                                    var subgroup = sheet.Cells[innerRow, 1].Value?.ToString()?.Trim();
                                    if (string.IsNullOrWhiteSpace(subgroup))
                                        break; // end of outcome

                                    var nStr = sheet.Cells[innerRow, column + 2].Value?.ToString()?.Trim();
                                    if (string.IsNullOrWhiteSpace(nStr))
                                        continue; // no data provided

                                    var n = int.Parse(nStr);

                                    var NStr = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                    if (string.IsNullOrWhiteSpace(NStr)) // take the previous one
                                        NStr = sheet.Cells[innerRow, column + 3 - 2].Value?.ToString()?.Trim();

                                    var N = int.Parse(NStr);

                                    Console.WriteLine("Study {0} Outcome {1} / {2} n/N {3}/{4}",
                                        study.Name,
                                        outcome,
                                        comparison,
                                        n,
                                        N);

                                    outputFile.WriteLine(
                                        JsonConvert.SerializeObject(
                                            new
                                            {
                                                Type="Dico",
                                                study.RevManStudyId,
                                                study.WorksheetIndex,
                                                StudyName = study.Name,
                                                Outcome = outcome,
                                                Comparison = comparison,
                                                Subgroup = subgroup,
                                                n,
                                                N
                                            }
                                        ));
                                }
                            }
                            else if (header1 == "mean")
                            {
                                string header2 = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                string header3 = sheet.Cells[innerRow, column + 4].Value?.ToString()?.Trim();

                                if (header2 == "SD" || header3 == "N")
                                {
                                    innerRow++;

                                    for (; innerRow < studyIdentificationRow - 4; innerRow++)
                                    {
                                        var subgroup = sheet.Cells[innerRow, 1].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(subgroup))
                                            break; // end of outcome

                                        var meanStr = sheet.Cells[innerRow, column + 2].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(meanStr))
                                            continue; // no data provided

                                        var mean = double.Parse(meanStr);

                                        var sdStr = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(sdStr))
                                            continue; // no data provided

                                        var sd = double.Parse(sdStr);

                                        var NStr = sheet.Cells[innerRow, column + 4].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(NStr)) // take the previous one
                                            NStr = sheet.Cells[innerRow, column + 4 - 2].Value?.ToString()?.Trim();

                                        try
                                        {
                                            var N = int.Parse(NStr);

                                            Console.WriteLine("Study {0} Outcome {1} / {2} mean/SD/N {3} +/- {4} ({5})",
                                                study.Name,
                                                outcome,
                                                comparison,
                                                mean,
                                                sd,
                                                N);

                                            outputFile.WriteLine(
                                            JsonConvert.SerializeObject(
                                                new
                                                {
                                                    Type = "Cont",
                                                    study.RevManStudyId,
                                                    study.WorksheetIndex,
                                                    StudyName = study.Name,
                                                    Outcome = outcome,
                                                    Comparison = comparison,
                                                    Subgroup = subgroup,
                                                    mean,
                                                    sd,
                                                    N
                                                }
                                            ));
                                        }
                                        catch (Exception e)
                                        {
                                            Console.WriteLine("Error Study {0} Outcome {1} / {2} mean/SD/N {3} +/- {4} ({5}) {6}",
                                                study.Name,
                                                outcome,
                                                comparison,
                                                mean,
                                                sd,
                                                NStr,
                                                e.Message);
                                        }
                                    }
                                }
                                else if (header2 == "SE" || header3 == "N")
                                {
                                    innerRow++;

                                    for (; innerRow < studyIdentificationRow - 4; innerRow++)
                                    {
                                        var subgroup = sheet.Cells[innerRow, 1].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(subgroup))
                                            break; // end of outcome

                                        var meanStr = sheet.Cells[innerRow, column + 2].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(meanStr))
                                            continue; // no data provided

                                        var mean = double.Parse(meanStr);

                                        var seStr = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(seStr))
                                            continue; // no data provided

                                        var se = double.Parse(seStr);

                                        var NStr = sheet.Cells[innerRow, column + 4].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(NStr)) // take the previous one
                                            NStr = sheet.Cells[innerRow, column + 4 - 2].Value?.ToString()?.Trim();

                                        var N = int.Parse(NStr);

                                        Console.WriteLine("Study {0} Outcome {1} / {2} mean/SE/N {3} +/- {4} ({5})",
                                            study.Name,
                                            outcome,
                                            comparison,
                                            mean,
                                            se,
                                            N);

                                        outputFile.WriteLine(
                                            JsonConvert.SerializeObject(
                                                new
                                                {
                                                    Type = "Cont",
                                                    study.RevManStudyId,
                                                    study.WorksheetIndex,
                                                    StudyName = study.Name,
                                                    Outcome = outcome,
                                                    Comparison = comparison,
                                                    Subgroup = subgroup,
                                                    mean,
                                                    sd = se * Math.Pow(N, 2),
                                                    N,
                                                    se
                                                }
                                            ));
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Skipping mean {0} / {1} / {2}", header1, header2, header3);
                                }
                            }
                            else if (!string.IsNullOrWhiteSpace(header1) && header1.ToLower() == "median")
                            {
                                // median, what else?
                                string header2 = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                string header3 = sheet.Cells[innerRow, column + 4].Value?.ToString()?.Trim();
                                string header4 = sheet.Cells[innerRow, column + 5].Value?.ToString()?.Trim();

                                if (((header2 == "25th" && header3 == "75th") || 
                                     (header2.ToLower() == "range (low)" && header3.ToLower() == "range (high)"))
                                    && header4 == "N" )
                                {
                                    innerRow++;

                                    for (; innerRow < studyIdentificationRow - 4; innerRow++)
                                    {
                                        var subgroup = sheet.Cells[innerRow, 1].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(subgroup))
                                            break; // end of outcome

                                        var medianStr= sheet.Cells[innerRow, column + 2].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(medianStr))
                                            continue; // no data provided

                                        var median = double.Parse(medianStr);

                                        var lowerStr = sheet.Cells[innerRow, column + 3].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(lowerStr))
                                            continue; // no data provided

                                        var lower = double.Parse(lowerStr);

                                        var upperStr = sheet.Cells[innerRow, column + 4].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(upperStr))
                                            continue; // no data provided

                                        var upper = double.Parse(upperStr);

                                        var NStr = sheet.Cells[innerRow, column + 5].Value?.ToString()?.Trim();
                                        if (string.IsNullOrWhiteSpace(NStr)) // take the previous one
                                            NStr = sheet.Cells[innerRow, column + 5 - 2].Value?.ToString()?.Trim();

                                        var N = int.Parse(NStr);

                                        Console.WriteLine("Study {0} Outcome {1} / {2} median/25th/75th/N {3} {4}-{5} ({6})",
                                            study.Name,
                                            outcome,
                                            comparison,
                                            median,
                                            lower,
                                            upper,
                                            N);

                                        // Hozo: http://vassarstats.net/median_range.html
                                        double sd;
                                        if (N <= 15)
                                            sd = Math.Sqrt((Math.Pow(lower - 2 * median + upper, 2) / 4 + Math.Pow(upper - lower, 2)) / 12.0);
                                        else if (N <= 70)
                                            sd = (upper - lower) / 4;
                                        else
                                            sd = (upper - lower) / 6;

                                        outputFile.WriteLine(
                                            JsonConvert.SerializeObject(
                                                new
                                                {
                                                    Type = "Cont",
                                                    study.RevManStudyId,
                                                    study.WorksheetIndex,
                                                    StudyName = study.Name,
                                                    Outcome = outcome,
                                                    Comparison = comparison,
                                                    Subgroup = subgroup,
                                                    mean = (lower + 2*median + upper) / 4.0,
                                                    sd,
                                                    N,
                                                    median,
                                                    lower,
                                                    upper
                                                }
                                            ));
                                    }
                                }
                                else 
                                    Console.WriteLine("Found median unknown {0} / {1} / {2} / {3}", header1, header2, header3, header4);
                            }
                            else
                            { 
                                if (!string.IsNullOrWhiteSpace(header1))
                                    Console.WriteLine("Skipping '{0}' for {1}", outcome, header1);

                                // skip for now
                                innerRow++;

                                // skip until whitespace in column 1
                                for (; innerRow < studyIdentificationRow - 4 && !string.IsNullOrWhiteSpace(sheet.Cells[innerRow, 1].Value); innerRow++) ;
                            }

                            innerRowEnd = innerRow;
                        }

                        row = innerRowEnd;
                    }
                }


            // until "Study Identification"
            }
        }
    }
}
