namespace pdf2rtf
{
    enum AmbientType
    {
        Unknown,
        Measured,
        PrePost
    }

    enum SpirometryType
    {
        Unknown,
        SixColumn,
        ThreeColumn
    }

    enum DiffusionType
    {
        Unknown,
        None,
        RefPost,
        RefPre
    }

    class ReportData
    {
        public AmbientType AmbientType { get; set; } = AmbientType.Unknown;
        public SpirometryType SpirometryType { get; set; } = SpirometryType.Unknown;
        public DiffusionType DiffusionType { get; set; } = DiffusionType.Unknown;

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

        public AmbientData[] AmbientData { get; set; }

        public string[] FEV1 { get; set; }
        public string[] FVC { get; set; }
        public string[] FEV1I { get; set; }
        public string[] MMEF { get; set; }
        public string[] MEF50 { get; set; }
        public string[] PEF { get; set; }

        public string[] DLCO { get; set; }
        public string[] DLCOc { get; set; }
        public string[] KCO { get; set; }
        public string[] KCOc { get; set; }
        public string[] VA { get; set; }
        public string[] Hb { get; set; }

        public string TechnicianNotes { get; set; }
        public string Interpretation { get; set; }
    }

    class AmbientData
    {
        public string DateTime { get; set; }
        public string Ambient { get; set; }
    }
}
