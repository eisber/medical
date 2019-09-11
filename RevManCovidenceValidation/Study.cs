namespace RevManCovidenceValidation
{
    public class Study
    {
        public string Name { get; set; }

        public string Title { get; set; }

        public int WorksheetIndex { get; set; }

        public string RevManStudyId { get; set; }

        public override string ToString()
        {
            return string.Format("{0} - {1}", Name, Title);
        }
    }
}
