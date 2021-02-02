using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevManCovidenceValidation
{
    public class CsvOutcome
    {
        public string Study { get; set; }

        public string EndNoteStudy { get; set; }
        public string Outcome { get;set; }

        public int Year { get; set; }
        public string StudyType { get; set; }

        public bool THA { get; set; }
        public bool TKA { get; set; }
        public bool GA { get; set; }
        public bool NA { get; set; }

        public int Total1 { get; set; }
        public int Total2 { get; set; }
    }

    public class CsvOutcomeDich : CsvOutcome
    {
        public int Events1 { get; set; }
        public int Events2 { get; set; }
    }

    public class CsvOutcomeCont : CsvOutcome
    {
        public double Mean1 { get; set; }
        public double Mean2 { get; set; }
        public double SD1 { get; set; }
        public double SD2 { get; set; }
    }
}
