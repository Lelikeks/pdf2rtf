namespace pdf2rtf
{
    class ReportData
    {
        public string PatientId { get; set; }
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string DateOfBirth { get; set; }
        public string Age { get; set; }
        public string Height { get; set; }
        public string History { get; set; }
        public string Physician { get; set; }
        public string Insurance { get; set; }
        public string VisitID { get; set; }
        public string Smoker { get; set; }
        public string PackYears { get; set; }
        public string BMI { get; set; }
        public string Gender { get; set; }
        public string Weight { get; set; }
        public string Ward { get; set; }

        public AmbientData AmbientPre { get; set; }
        public AmbientData AmbientPost { get; set; }

        public FirstData FEV1 { get; set; }
        public FirstData FVC { get; set; }
        public FirstData FEV1I { get; set; }
        public FirstData MMEF { get; set; }
        public FirstData MEF50 { get; set; }
        public FirstData PEF { get; set; }

        public SecondData DLCO { get; set; }
        public SecondData DLCOc { get; set; }
        public SecondData KCO { get; set; }
        public SecondData KCOc { get; set; }
        public SecondData VA { get; set; }
        public SecondData Hb { get; set; }

        public string TechnicianNotes { get; set; }
        public string Interpretation { get; set; }
    }

    class AmbientData
    {
        public string DateTime { get; set; }
        public string Ambient { get; set; }
    }

    class FirstData
    {
        public string Ref { get; set; }
        public string Pre { get; set; }
        public string PreRef { get; set; }
        public string Post { get; set; }
        public string RefPercent { get; set; }
        public string PrePost { get; set; }
    }

    class SecondData
    {
        public string Ref { get; set; }
        public string Post { get; set; }
        public string RefPercent { get; set; }
    }
}
