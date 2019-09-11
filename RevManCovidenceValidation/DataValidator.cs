using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RevManCovidenceValidation
{
    public partial class DataValidator : IDisposable
    {
        private Application excel;
        private Workbook covidenceExcel;
        private XDocument revmanXml;
        private string basePath;
        public string CovidenceDataPath { get; private set; }
        private string mappingFileOutcomes;
        private string mappingFileComparison;

        public void Step1_Load(string revmanFile, string covidenceFile, string mappingFileOutcomes, string mappingFileComparison)
        {
            this.mappingFileOutcomes = mappingFileOutcomes;
            this.mappingFileComparison = mappingFileComparison;

            basePath = Path.GetDirectoryName(revmanFile);
            CovidenceDataPath = basePath + "\\covidence_data.json";

            revmanXml = XDocument.Parse(File.ReadAllText(revmanFile));

            if (!File.Exists(CovidenceDataPath))
            {
                excel = new Application();
                covidenceExcel = excel.Workbooks.Open(covidenceFile);
            }
        }

        public void Dispose()
        {
            covidenceExcel?.Close(0);
            excel?.Quit();
        }
    }
}
