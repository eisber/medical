using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.XPath;
using F23.StringSimilarity;

namespace RevManCovidenceValidation
{
    partial class Program
    {
        
        static void Main(string[] args)
        {
            // RevManSummaryToCSV.ExtractOutcomes();

            // 2020/08/30 #1
            // EnrichedRevManToSubgroup.Process_Observational_RCT();

            // 2020/08/30 #2
            //EnrichedRevManToSubgroup.Process_THA_TKA();

            // 2020/08/30 #1
            //EnrichedRevManToSubgroup.Process_GRADE();

            EnrichedRevManToSubgroup.Process_CSV();

            // Note: re-number all to get re-calculation
            // Note: if RevMan crashes on print (or can't re-open after save). Copy&paste the subtree into another revman file
            // RevManSummaryToCSV.Process();

            //EnrichedRevManToSubgroup.Process_Infiltration();


            //if (args.Length != 4)
            //{
            //    Console.Error.WriteLine("Usage: RevManCovidenceValidation.exe <revman5-input.rm5> <covidence-excel-export.xls> <mapping.txt>");
            //    return;
            //}

            //var revmanFile = args[0];
            //var covidenceFile = args[1];
            //var mappingFileOutcomes = args[2];
            //var mappingFileComparison = args[3];

            //using (var validator = new DataValidator())
            //{
            //    validator.Step1_Load(revmanFile, covidenceFile, mappingFileOutcomes, mappingFileComparison);

            //    if (!File.Exists(validator.CovidenceDataPath))
            //    {
            //        validator.Step2_MatchRevmanAndCovidenceStudies();
            //        validator.Step3_EnrichOutcomes();
            //    }
            //    else
            //    {
            //        validator.Step4_ValidateRevMan();
            //    }
            //}


            Console.WriteLine("Press any key");
            Console.ReadKey();
        }
    }
}
