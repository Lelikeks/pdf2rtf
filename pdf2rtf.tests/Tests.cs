using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Schema;
using Xunit;
using Xunit.Abstractions;

namespace pdf2rtf.tests
{
	public class Tests
	{
		private readonly ITestOutputHelper _testOutputHelper;

		public Tests(ITestOutputHelper testOutputHelper)
		{
			_testOutputHelper = testOutputHelper;
		}

		public void WriteLine(object message)
		{
			_testOutputHelper.WriteLine(message == null ? "<null>" : message.ToString());
		}
		
		[Fact]
		public void Test1()
		{
			var data = PdfParser.Parse(@"..\..\..\SampleData\Awatef (Whsc).pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPre, data.DiffusionType);
			Assert.Equal("FALTAOUS_A18102017", data.PatientId);
			Assert.Equal("FALTAOUS", data.LastName);
			Assert.Equal("Awatef", data.FirstName);
			Assert.Equal("2/11/1966", data.DateOfBirth);
			Assert.Equal("50 years", data.Age);
			Assert.Equal("156.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("Dr R Abdelmalak", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("18/10/2017", data.VisitID);
			Assert.Equal("Never Smoked", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("44.4", data.BMI);
			Assert.Equal("Female", data.Gender);
			Assert.Equal("108.0 kg", data.Weight);
			Assert.Equal("FB", data.Technician);
			Assert.Equal("Silverton Medical Clinic", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("18/10/2017 9:43 AM", data.AmbientData[0].DateTime);
			Assert.Equal("25.3 \\'b0C   789.75 mmHg   37 %", data.AmbientData[0].Ambient);
			Assert.Equal("18/10/2017 9:38 AM", data.AmbientData[1].DateTime);
			Assert.Equal("25.1 \\'b0C   789.89 mmHg   37 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("2.53", data.SpirometryData[0][0]);
			Assert.Equal("1.82", data.SpirometryData[0][1]);
			Assert.Equal("72 %", data.SpirometryData[0][2]);
			Assert.Equal("1.79", data.SpirometryData[0][3]);
			Assert.Equal("70 %", data.SpirometryData[0][4]);
			Assert.Equal("-2 %", data.SpirometryData[0][5]);
			Assert.Equal("3.15", data.SpirometryData[1][0]);
			Assert.Equal("2.36", data.SpirometryData[1][1]);
			Assert.Equal("75 %", data.SpirometryData[1][2]);
			Assert.Equal("2.37", data.SpirometryData[1][3]);
			Assert.Equal("75 %", data.SpirometryData[1][4]);
			Assert.Equal("1 %", data.SpirometryData[1][5]);
			Assert.Equal(" 80.90", data.SpirometryData[2][0]);
			Assert.Equal("76.95", data.SpirometryData[2][1]);
			Assert.Equal("95 %", data.SpirometryData[2][2]);
			Assert.Equal("75.30", data.SpirometryData[2][3]);
			Assert.Equal("93 %", data.SpirometryData[2][4]);
			Assert.Equal("-2 %", data.SpirometryData[2][5]);
			Assert.Equal("2.58", data.SpirometryData[3][0]);
			Assert.Equal("1.63", data.SpirometryData[3][1]);
			Assert.Equal("63 %", data.SpirometryData[3][2]);
			Assert.Equal("1.62", data.SpirometryData[3][3]);
			Assert.Equal("63 %", data.SpirometryData[3][4]);
			Assert.Equal(" 0 %", data.SpirometryData[3][5]);
			Assert.Equal("3.73", data.SpirometryData[4][0]);
			Assert.Equal("2.45", data.SpirometryData[4][1]);
			Assert.Equal("66 %", data.SpirometryData[4][2]);
			Assert.Equal("1.80", data.SpirometryData[4][3]);
			Assert.Equal("48 %", data.SpirometryData[4][4]);
			Assert.Equal("-27 %", data.SpirometryData[4][5]);
			Assert.Equal("5.97", data.SpirometryData[5][0]);
			Assert.Equal("3.51", data.SpirometryData[5][1]);
			Assert.Equal("59 %", data.SpirometryData[5][2]);
			Assert.Equal("2.28", data.SpirometryData[5][3]);
			Assert.Equal("38 %", data.SpirometryData[5][4]);
			Assert.Equal("-35 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("22.61", data.DiffusionData[0][0]);
			Assert.Equal("21.41", data.DiffusionData[0][1]);
			Assert.Equal("95 %", data.DiffusionData[0][2]);
			Assert.Equal("22.61", data.DiffusionData[1][0]);
			Assert.Equal("21.41", data.DiffusionData[1][1]);
			Assert.Equal("95 %", data.DiffusionData[1][2]);
			Assert.Equal("5.02", data.DiffusionData[2][0]);
			Assert.Equal("6.52", data.DiffusionData[2][1]);
			Assert.Equal("130 %", data.DiffusionData[2][2]);
			Assert.Equal("5.02", data.DiffusionData[3][0]);
			Assert.Equal("6.52", data.DiffusionData[3][1]);
			Assert.Equal("130 %", data.DiffusionData[3][2]);
			Assert.Equal("4.36", data.DiffusionData[4][0]);
			Assert.Equal("3.28", data.DiffusionData[4][1]);
			Assert.Equal("75 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("Patient performed test to their best ability. ", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 9:30 AM: Spirometry results indicate a restrictive ventilatory defect which was mild in severity. /*/* Automatic Interpretation - Forced Spirometry - 9:38 AM: Spirometry results indicate a restrictive ventilatory defect which was mild in severity.There is no significant change following inhaled bronchodilator on this occasion. /*/* Automatic Interpretation - Single Breath Diffusion - 9:43 AM: CO Diffusion was within normal limits. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\Awatef (Whsc).rtf"), File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test2()
		{
			var data = PdfParser.Parse(@"..\..\..\SampleData\ian (TCR).pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPost, data.DiffusionType);
			Assert.Equal("Busuttil_I18102017", data.PatientId);
			Assert.Equal("Busuttil", data.LastName);
			Assert.Equal("Ian", data.FirstName);
			Assert.Equal("8/09/1968", data.DateOfBirth);
			Assert.Equal("49 years", data.Age);
			Assert.Equal("178.0 cm", data.Height);
			Assert.Equal("Asthma", data.History);
			Assert.Equal("Dr Wally Hosn", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("18/10/2017", data.VisitID);
			Assert.Equal("Ex Smoker", data.Smoker);
			Assert.Equal("2.50", data.PackYears);
			Assert.Equal("31.6", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("100.0 kg", data.Weight);
			Assert.Equal("Allie", data.Technician);
			Assert.Equal("Settlement Road Medical Centre", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("18/10/2017 9:14 AM", data.AmbientData[0].DateTime);
			Assert.Equal("27.7 \\'b0C   790.65 mmHg   34 %", data.AmbientData[0].Ambient);
			Assert.Equal("18/10/2017 9:23 AM", data.AmbientData[1].DateTime);
			Assert.Equal("28.1 \\'b0C   790.65 mmHg   34 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("3.97", data.SpirometryData[0][0]);
			Assert.Equal("3.80", data.SpirometryData[0][1]);
			Assert.Equal("96 %", data.SpirometryData[0][2]);
			Assert.Equal("4.07", data.SpirometryData[0][3]);
			Assert.Equal("102 %", data.SpirometryData[0][4]);
			Assert.Equal("7 %", data.SpirometryData[0][5]);
			Assert.Equal("5.03", data.SpirometryData[1][0]);
			Assert.Equal("4.64", data.SpirometryData[1][1]);
			Assert.Equal("92 %", data.SpirometryData[1][2]);
			Assert.Equal("4.81", data.SpirometryData[1][3]);
			Assert.Equal("96 %", data.SpirometryData[1][4]);
			Assert.Equal("4 %", data.SpirometryData[1][5]);
			Assert.Equal(" 79.24", data.SpirometryData[2][0]);
			Assert.Equal("81.86", data.SpirometryData[2][1]);
			Assert.Equal("103 %", data.SpirometryData[2][2]);
			Assert.Equal("84.51", data.SpirometryData[2][3]);
			Assert.Equal("107 %", data.SpirometryData[2][4]);
			Assert.Equal("3 %", data.SpirometryData[2][5]);
			Assert.Equal("3.65", data.SpirometryData[3][0]);
			Assert.Equal("4.04", data.SpirometryData[3][1]);
			Assert.Equal("111 %", data.SpirometryData[3][2]);
			Assert.Equal("5.00", data.SpirometryData[3][3]);
			Assert.Equal("137 %", data.SpirometryData[3][4]);
			Assert.Equal("24 %", data.SpirometryData[3][5]);
			Assert.Equal("4.88", data.SpirometryData[4][0]);
			Assert.Equal("4.86", data.SpirometryData[4][1]);
			Assert.Equal("100 %", data.SpirometryData[4][2]);
			Assert.Equal("6.63", data.SpirometryData[4][3]);
			Assert.Equal("136 %", data.SpirometryData[4][4]);
			Assert.Equal("37 %", data.SpirometryData[4][5]);
			Assert.Equal("8.97", data.SpirometryData[5][0]);
			Assert.Equal("9.01", data.SpirometryData[5][1]);
			Assert.Equal("100 %", data.SpirometryData[5][2]);
			Assert.Equal("9.28", data.SpirometryData[5][3]);
			Assert.Equal("103 %", data.SpirometryData[5][4]);
			Assert.Equal("3 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("31.39", data.DiffusionData[0][0]);
			Assert.Equal("36.55", data.DiffusionData[0][1]);
			Assert.Equal("116 %", data.DiffusionData[0][2]);
			Assert.Equal("31.39", data.DiffusionData[1][0]);
			Assert.Equal("36.55", data.DiffusionData[1][1]);
			Assert.Equal("116 %", data.DiffusionData[1][2]);
			Assert.Equal("4.39", data.DiffusionData[2][0]);
			Assert.Equal("5.38", data.DiffusionData[2][1]);
			Assert.Equal("122 %", data.DiffusionData[2][2]);
			Assert.Equal("4.39", data.DiffusionData[3][0]);
			Assert.Equal("5.38", data.DiffusionData[3][1]);
			Assert.Equal("122 %", data.DiffusionData[3][2]);
			Assert.Equal("6.99", data.DiffusionData[4][0]);
			Assert.Equal("6.80", data.DiffusionData[4][1]);
			Assert.Equal("97 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Null(data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 9:14 AM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Single Breath Diffusion - 9:22 AM: CO Diffusion was within normal limits. /*/* Automatic Interpretation - Forced Spirometry - 9:23 AM: Spirometry results are within normal limits.There is no significant change following inhaled bronchodilator on this occasion. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\ian (TCR).rtf"), File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test3()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-01-9_11-46-12 - Heartscope3Report Dimitra Fagridas - spiro only.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPre, data.DiffusionType);
			Assert.Equal("Fagridas_D010917", data.PatientId);
			Assert.Equal("Fagridas", data.LastName);
			Assert.Equal("Dimitra", data.FirstName);
			Assert.Equal("27/04/1968", data.DateOfBirth);
			Assert.Equal("49 years", data.Age);
			Assert.Equal("167.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("Dr Ali Al-Fiadh", data.Physician);
			Assert.Equal("Dr Carol Lawson", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("39.4", data.BMI);
			Assert.Equal("Female", data.Gender);
			Assert.Equal("110.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("Moone Ponds Specialist Centre", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("1/09/2017 11:25 AM", data.AmbientData[0].DateTime);
			Assert.Equal("25.3 \\'b0C   792 mmHg   34 %", data.AmbientData[0].Ambient);
			Assert.Equal("1/09/2017 11:15 AM", data.AmbientData[1].DateTime);
			Assert.Equal("24.8 \\'b0C   792 mmHg   34 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("2.96", data.SpirometryData[0][0]);
			Assert.Equal("2.43", data.SpirometryData[0][1]);
			Assert.Equal("82 %", data.SpirometryData[0][2]);
			Assert.Equal("2.51", data.SpirometryData[0][3]);
			Assert.Equal("85 %", data.SpirometryData[0][4]);
			Assert.Equal("4 %", data.SpirometryData[0][5]);
			Assert.Equal("3.70", data.SpirometryData[1][0]);
			Assert.Equal("3.00", data.SpirometryData[1][1]);
			Assert.Equal("81 %", data.SpirometryData[1][2]);
			Assert.Equal("3.12", data.SpirometryData[1][3]);
			Assert.Equal("84 %", data.SpirometryData[1][4]);
			Assert.Equal("4 %", data.SpirometryData[1][5]);
			Assert.Equal(" 79.79", data.SpirometryData[2][0]);
			Assert.Equal("91.71", data.SpirometryData[2][1]);
			Assert.Equal("115 %", data.SpirometryData[2][2]);
			Assert.Equal("91.68", data.SpirometryData[2][3]);
			Assert.Equal("115 %", data.SpirometryData[2][4]);
			Assert.Equal(" 0 %", data.SpirometryData[2][5]);
			Assert.Equal("2.88", data.SpirometryData[3][0]);
			Assert.Equal("2.30", data.SpirometryData[3][1]);
			Assert.Equal("80 %", data.SpirometryData[3][2]);
			Assert.Equal("2.53", data.SpirometryData[3][3]);
			Assert.Equal("88 %", data.SpirometryData[3][4]);
			Assert.Equal("10 %", data.SpirometryData[3][5]);
			Assert.Equal("4.03", data.SpirometryData[4][0]);
			Assert.Equal("2.89", data.SpirometryData[4][1]);
			Assert.Equal("72 %", data.SpirometryData[4][2]);
			Assert.Equal("3.25", data.SpirometryData[4][3]);
			Assert.Equal("81 %", data.SpirometryData[4][4]);
			Assert.Equal("13 %", data.SpirometryData[4][5]);
			Assert.Equal("6.61", data.SpirometryData[5][0]);
			Assert.Equal("5.33", data.SpirometryData[5][1]);
			Assert.Equal("81 %", data.SpirometryData[5][2]);
			Assert.Equal("5.81", data.SpirometryData[5][3]);
			Assert.Equal("88 %", data.SpirometryData[5][4]);
			Assert.Equal("9 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("25.44", data.DiffusionData[0][0]);
			Assert.Equal("- ", data.DiffusionData[0][1]);
			Assert.Equal("-", data.DiffusionData[0][2]);
			Assert.Equal("25.44", data.DiffusionData[1][0]);
			Assert.Equal("- ", data.DiffusionData[1][1]);
			Assert.Equal("-", data.DiffusionData[1][2]);
			Assert.Equal("4.86", data.DiffusionData[2][0]);
			Assert.Equal("- ", data.DiffusionData[2][1]);
			Assert.Equal("-", data.DiffusionData[2][2]);
			Assert.Equal("4.86", data.DiffusionData[3][0]);
			Assert.Equal("- ", data.DiffusionData[3][1]);
			Assert.Equal("-", data.DiffusionData[3][2]);
			Assert.Equal("5.08", data.DiffusionData[4][0]);
			Assert.Equal("3.64", data.DiffusionData[4][1]);
			Assert.Equal("72 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("the patient complaints of SOB, cough and wheezing. She was given 6 courses of antibiotics with no improvement the cough still persists. She is a smoker, smokes 20 cigarettes a day for 30yrs, last had a cigarette, 12hrs prior to the test. Is currently on ventolin which she last had it 12hrs prior to the test.The patient had difficulty on fully inhaling and holding on DLCO she only attempted it once despite sevra; attempts she mentioned that she was unable to do the DLCO.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 11:15 AM: Normal SpirometryThere is no significant change following inhaled bronchodilator on this occasion. /*/* Automatic Interpretation - Single Breath Diffusion - 11:25 AM: - /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-01-9_11-46-12 - Heartscope3Report Dimitra Fagridas - spiro only.rtf"),
				File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test4()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-11-8_15-51-52 - Heartscope3Report.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.None, data.DiffusionType);
			Assert.Equal("DBTest", data.PatientId);
			Assert.Equal("TestDB", data.LastName);
			Assert.Equal("DB", data.FirstName);
			Assert.Equal("1/01/1980", data.DateOfBirth);
			Assert.Equal("37 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("SOB", data.History);
			Assert.Equal("Dr Alex Smith", data.Physician);
			Assert.Equal("Spring Hill Medical Centre", data.Insurance);
			Assert.Equal("4/8/17", data.VisitID);
			Assert.Equal("Ex Smoker", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("27.8", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("90.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("Spring Hill Medical Centre", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("4/08/2017 2:21 PM", data.AmbientData[0].DateTime);
			Assert.Equal("18.5 \\'b0C   727 mmHg   54 %", data.AmbientData[0].Ambient);
			Assert.Equal("4/08/2017 2:23 PM", data.AmbientData[1].DateTime);
			Assert.Equal("18.9 \\'b0C   726 mmHg   55 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.43", data.SpirometryData[0][0]);
			Assert.Equal("3.71", data.SpirometryData[0][1]);
			Assert.Equal("84 %", data.SpirometryData[0][2]);
			Assert.Equal("3.90", data.SpirometryData[0][3]);
			Assert.Equal("88 %", data.SpirometryData[0][4]);
			Assert.Equal("5 %", data.SpirometryData[0][5]);
			Assert.Equal("5.47", data.SpirometryData[1][0]);
			Assert.Equal("4.77", data.SpirometryData[1][1]);
			Assert.Equal("87 %", data.SpirometryData[1][2]);
			Assert.Equal("5.03", data.SpirometryData[1][3]);
			Assert.Equal("92 %", data.SpirometryData[1][4]);
			Assert.Equal("6 %", data.SpirometryData[1][5]);
			Assert.Equal(" 80.55", data.SpirometryData[2][0]);
			Assert.Equal("91.27", data.SpirometryData[2][1]);
			Assert.Equal("113 %", data.SpirometryData[2][2]);
			Assert.Equal("98.67", data.SpirometryData[2][3]);
			Assert.Equal("122 %", data.SpirometryData[2][4]);
			Assert.Equal("8 %", data.SpirometryData[2][5]);
			Assert.Equal("4.34", data.SpirometryData[3][0]);
			Assert.Equal("3.29", data.SpirometryData[3][1]);
			Assert.Equal("76 %", data.SpirometryData[3][2]);
			Assert.Equal("3.33", data.SpirometryData[3][3]);
			Assert.Equal("77 %", data.SpirometryData[3][4]);
			Assert.Equal("1 %", data.SpirometryData[3][5]);
			Assert.Equal("5.33", data.SpirometryData[4][0]);
			Assert.Equal("3.85", data.SpirometryData[4][1]);
			Assert.Equal("72 %", data.SpirometryData[4][2]);
			Assert.Equal("3.82", data.SpirometryData[4][3]);
			Assert.Equal("72 %", data.SpirometryData[4][4]);
			Assert.Equal("-1 %", data.SpirometryData[4][5]);
			Assert.Equal("9.61", data.SpirometryData[5][0]);
			Assert.Equal("8.26", data.SpirometryData[5][1]);
			Assert.Equal("86 %", data.SpirometryData[5][2]);
			Assert.Equal("9.64", data.SpirometryData[5][3]);
			Assert.Equal("100 %", data.SpirometryData[5][4]);
			Assert.Equal("17 %", data.SpirometryData[5][5]);
			Assert.Null(data.DiffusionData);
			Assert.Equal("Good patient co-operation.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 2:21 PM: Normal Spirometry /*/* Automatic Interpretation - Forced Spirometry - 2:30 PM: Normal SpirometryThere is no significant change following inhaled bronchodilator on this occasion. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-11-8_15-51-52 - Heartscope3Report.rtf"),
				File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test5()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-11-8_15-53-03 - Heartscope3Report.pdf");
			#region TestData
			Assert.Equal(AmbientType.Measured, data.AmbientType);
			Assert.Equal(SpirometryType.ThreeColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPre, data.DiffusionType);
			Assert.Equal("BioCal", data.PatientId);
			Assert.Equal("Justin", data.LastName);
			Assert.Equal("Pilmore", data.FirstName);
			Assert.Equal("15/10/1981", data.DateOfBirth);
			Assert.Equal("35 years", data.Age);
			Assert.Equal("181.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("0.00", data.PackYears);
			Assert.Equal("24.4", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("80.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(1, data.AmbientData.Length);
			Assert.Equal("1/06/2017 3:32 PM", data.AmbientData[0].DateTime);
			Assert.Equal("23.0 \\'b0C   806 mmHg   50 %", data.AmbientData[0].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.54", data.SpirometryData[0][0]);
			Assert.Equal("3.83", data.SpirometryData[0][1]);
			Assert.Equal("84 %", data.SpirometryData[0][2]);
			Assert.Equal("5.59", data.SpirometryData[1][0]);
			Assert.Equal("4.87", data.SpirometryData[1][1]);
			Assert.Equal("87 %", data.SpirometryData[1][2]);
			Assert.Equal(" 80.91", data.SpirometryData[2][0]);
			Assert.Equal("76.45", data.SpirometryData[2][1]);
			Assert.Equal("94 %", data.SpirometryData[2][2]);
			Assert.Equal("4.47", data.SpirometryData[3][0]);
			Assert.Equal("3.33", data.SpirometryData[3][1]);
			Assert.Equal("75 %", data.SpirometryData[3][2]);
			Assert.Equal("5.42", data.SpirometryData[4][0]);
			Assert.Equal("3.95", data.SpirometryData[4][1]);
			Assert.Equal("73 %", data.SpirometryData[4][2]);
			Assert.Equal("9.76", data.SpirometryData[5][0]);
			Assert.Equal("10.93", data.SpirometryData[5][1]);
			Assert.Equal("112 %", data.SpirometryData[5][2]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("35.14", data.DiffusionData[0][0]);
			Assert.Equal("30.26", data.DiffusionData[0][1]);
			Assert.Equal("86 %", data.DiffusionData[0][2]);
			Assert.Equal("35.14", data.DiffusionData[1][0]);
			Assert.Equal("30.26", data.DiffusionData[1][1]);
			Assert.Equal("86 %", data.DiffusionData[1][2]);
			Assert.Equal("4.76", data.DiffusionData[2][0]);
			Assert.Equal("4.45", data.DiffusionData[2][1]);
			Assert.Equal("93 %", data.DiffusionData[2][2]);
			Assert.Equal("4.76", data.DiffusionData[3][0]);
			Assert.Equal("4.45", data.DiffusionData[3][1]);
			Assert.Equal("93 %", data.DiffusionData[3][2]);
			Assert.Equal("7.23", data.DiffusionData[4][0]);
			Assert.Equal("6.80", data.DiffusionData[4][1]);
			Assert.Equal("94 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Null(data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 3:26 PM: Normal Spirometry /*/* Automatic Interpretation - Single Breath Diffusion - 3:44 PM: Normal Diffusion /*", data.Interpretation);
			#endregion
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-11-8_15-53-03 - Heartscope3Report.rtf"),
				File.ReadAllText(temp));
		}

		[Fact]
		public void Test6()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-11-8_15-53-34 - Heartscope3Report.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPre, data.DiffusionType);
			Assert.Equal("_Ganshorn_TP", data.PatientId);
			Assert.Equal("_Muster", data.LastName);
			Assert.Equal("Patient", data.FirstName);
			Assert.Equal("7/03/1986", data.DateOfBirth);
			Assert.Equal("31 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("SOB", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("Spring Hill Medical Centre", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("23.1", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("75.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("12/04/2016 8:42 PM", data.AmbientData[0].DateTime);
			Assert.Equal("21.5 \\'b0C   738 mmHg   50 %", data.AmbientData[0].Ambient);
			Assert.Equal("12/04/2016 5:26 PM", data.AmbientData[1].DateTime);
			Assert.Equal("21.0 \\'b0C   738 mmHg   50 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.38", data.SpirometryData[0][0]);
			Assert.Equal("4.78", data.SpirometryData[0][1]);
			Assert.Equal("109 %", data.SpirometryData[0][2]);
			Assert.Equal("4.73", data.SpirometryData[0][3]);
			Assert.Equal("108 %", data.SpirometryData[0][4]);
			Assert.Equal("-1 %", data.SpirometryData[0][5]);
			Assert.Equal("5.25", data.SpirometryData[1][0]);
			Assert.Equal("6.02", data.SpirometryData[1][1]);
			Assert.Equal("115 %", data.SpirometryData[1][2]);
			Assert.Equal("6.15", data.SpirometryData[1][3]);
			Assert.Equal("117 %", data.SpirometryData[1][4]);
			Assert.Equal("2 %", data.SpirometryData[1][5]);
			Assert.Equal(" 81.81", data.SpirometryData[2][0]);
			Assert.Equal("78.18", data.SpirometryData[2][1]);
			Assert.Equal("96 %", data.SpirometryData[2][2]);
			Assert.Equal("76.99", data.SpirometryData[2][3]);
			Assert.Equal("94 %", data.SpirometryData[2][4]);
			Assert.Equal("-2 %", data.SpirometryData[2][5]);
			Assert.Equal("4.90", data.SpirometryData[3][0]);
			Assert.Equal("4.33", data.SpirometryData[3][1]);
			Assert.Equal("88 %", data.SpirometryData[3][2]);
			Assert.Equal("4.33", data.SpirometryData[3][3]);
			Assert.Equal("88 %", data.SpirometryData[3][4]);
			Assert.Equal(" 0 %", data.SpirometryData[3][5]);
			Assert.Equal("5.54", data.SpirometryData[4][0]);
			Assert.Equal("4.86", data.SpirometryData[4][1]);
			Assert.Equal("88 %", data.SpirometryData[4][2]);
			Assert.Equal("4.90", data.SpirometryData[4][3]);
			Assert.Equal("88 %", data.SpirometryData[4][4]);
			Assert.Equal("1 %", data.SpirometryData[4][5]);
			Assert.Equal("9.91", data.SpirometryData[5][0]);
			Assert.Equal("11.43", data.SpirometryData[5][1]);
			Assert.Equal("115 %", data.SpirometryData[5][2]);
			Assert.Equal("11.42", data.SpirometryData[5][3]);
			Assert.Equal("115 %", data.SpirometryData[5][4]);
			Assert.Equal(" 0 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("35.80", data.DiffusionData[0][0]);
			Assert.Equal("65.49", data.DiffusionData[0][1]);
			Assert.Equal("183 %", data.DiffusionData[0][2]);
			Assert.Equal("35.80", data.DiffusionData[1][0]);
			Assert.Equal("65.49", data.DiffusionData[1][1]);
			Assert.Equal("183 %", data.DiffusionData[1][2]);
			Assert.Equal("4.90", data.DiffusionData[2][0]);
			Assert.Equal("8.96", data.DiffusionData[2][1]);
			Assert.Equal("183 %", data.DiffusionData[2][2]);
			Assert.Equal("4.90", data.DiffusionData[3][0]);
			Assert.Equal("8.96", data.DiffusionData[3][1]);
			Assert.Equal("183 %", data.DiffusionData[3][2]);
			Assert.Equal("7.15", data.DiffusionData[4][0]);
			Assert.Equal("7.31", data.DiffusionData[4][1]);
			Assert.Equal("102 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("Patient co-operation difficult due to language barrier - possible mouth leak during testing procedures.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Single Breath Diffusion - 8:42 PM: Increased DLCO (Polycythemia, Asthma, increased cardiac output or capillary blood volume) /*/* Automatic Interpretation - Forced Spirometry - 5:26 PM: Normal SpirometryThere is no significant change following inhaled bronchodilator on this occasion. /*", data.Interpretation);
			#endregion
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-11-8_15-53-34 - Heartscope3Report.rtf"),
				File.ReadAllText(temp));
		}

		[Fact]
		public void Test7()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-14-8_16-26-12 - Heartscope3Report.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPost, data.DiffusionType);
			Assert.Equal("schillertest", data.PatientId);
			Assert.Equal("Test", data.LastName);
			Assert.Equal("Schiller", data.FirstName);
			Assert.Equal("16/11/1976", data.DateOfBirth);
			Assert.Equal("40 years", data.Age);
			Assert.Equal("183.0 cm", data.Height);
			Assert.Equal("COPD", data.History);
			Assert.Equal("Dr Vipul Kapadia", data.Physician);
			Assert.Equal("Dr Penny Wong", data.Insurance);
			Assert.Equal("14/8/2017", data.VisitID);
			Assert.Equal("Non Smoker", data.Smoker);
			Assert.Equal("0.00", data.PackYears);
			Assert.Equal("34.3", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("115.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("Heartscope Wheelers Hill", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("14/08/2017 4:13 PM", data.AmbientData[0].DateTime);
			Assert.Equal("27.0 \\'b0C   783 mmHg   37 %", data.AmbientData[0].Ambient);
			Assert.Equal("14/08/2017 4:16 PM", data.AmbientData[1].DateTime);
			Assert.Equal("27.2 \\'b0C   783 mmHg   37 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.50", data.SpirometryData[0][0]);
			Assert.Equal("3.03", data.SpirometryData[0][1]);
			Assert.Equal("67 %", data.SpirometryData[0][2]);
			Assert.Equal("3.28", data.SpirometryData[0][3]);
			Assert.Equal("73 %", data.SpirometryData[0][4]);
			Assert.Equal("8 %", data.SpirometryData[0][5]);
			Assert.Equal("5.63", data.SpirometryData[1][0]);
			Assert.Equal("3.03", data.SpirometryData[1][1]);
			Assert.Equal("54 %", data.SpirometryData[1][2]);
			Assert.Equal("4.04", data.SpirometryData[1][3]);
			Assert.Equal("72 %", data.SpirometryData[1][4]);
			Assert.Equal("33 %", data.SpirometryData[1][5]);
			Assert.Equal(" 80.01", data.SpirometryData[2][0]);
			Assert.Equal("101.04", data.SpirometryData[2][1]);
			Assert.Equal("126 %", data.SpirometryData[2][2]);
			Assert.Equal("84.10", data.SpirometryData[2][3]);
			Assert.Equal("105 %", data.SpirometryData[2][4]);
			Assert.Equal("-17 %", data.SpirometryData[2][5]);
			Assert.Equal("4.31", data.SpirometryData[3][0]);
			Assert.Equal("8.49", data.SpirometryData[3][1]);
			Assert.Equal("197 %", data.SpirometryData[3][2]);
			Assert.Equal("3.10", data.SpirometryData[3][3]);
			Assert.Equal("72 %", data.SpirometryData[3][4]);
			Assert.Equal("-63 %", data.SpirometryData[3][5]);
			Assert.Equal("5.35", data.SpirometryData[4][0]);
			Assert.Equal("9.01", data.SpirometryData[4][1]);
			Assert.Equal("169 %", data.SpirometryData[4][2]);
			Assert.Equal("3.15", data.SpirometryData[4][3]);
			Assert.Equal("59 %", data.SpirometryData[4][4]);
			Assert.Equal("-65 %", data.SpirometryData[4][5]);
			Assert.Equal("9.67", data.SpirometryData[5][0]);
			Assert.Equal("9.51", data.SpirometryData[5][1]);
			Assert.Equal("98 %", data.SpirometryData[5][2]);
			Assert.Equal("8.20", data.SpirometryData[5][3]);
			Assert.Equal("85 %", data.SpirometryData[5][4]);
			Assert.Equal("-14 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("34.82", data.DiffusionData[0][0]);
			Assert.Equal("31.94", data.DiffusionData[0][1]);
			Assert.Equal("92 %", data.DiffusionData[0][2]);
			Assert.Equal("34.82", data.DiffusionData[1][0]);
			Assert.Equal("34.79", data.DiffusionData[1][1]);
			Assert.Equal("100 %", data.DiffusionData[1][2]);
			Assert.Equal("4.62", data.DiffusionData[2][0]);
			Assert.Equal("5.45", data.DiffusionData[2][1]);
			Assert.Equal("118 %", data.DiffusionData[2][2]);
			Assert.Equal("4.62", data.DiffusionData[3][0]);
			Assert.Equal("5.93", data.DiffusionData[3][1]);
			Assert.Equal("128 %", data.DiffusionData[3][2]);
			Assert.Equal("7.39", data.DiffusionData[4][0]);
			Assert.Equal("5.87", data.DiffusionData[4][1]);
			Assert.Equal("79 %", data.DiffusionData[4][2]);
			Assert.Equal("-", data.DiffusionData[5][0]);
			Assert.Equal("7.45", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("Pre testing was with 3L syringe. Biological control with post test.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 4:15 PM: Mild RestrictionThere is no significant change following inhaled bronchodilator on this occasion. /*/* Automatic Interpretation - Single Breath Diffusion - 4:16 PM: Normal Diffusion /*", data.Interpretation);
			#endregion
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-14-8_16-26-12 - Heartscope3Report.rtf"),
				File.ReadAllText(temp));
		}

		[Fact]
		public void Test8()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-23-10 Kirk ONEILL.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPre, data.DiffusionType);
			Assert.Equal("O'NEILL_K23102017", data.PatientId);
			Assert.Equal("O'NEILL", data.LastName);
			Assert.Equal("Kirk", data.FirstName);
			Assert.Equal("2/12/1969", data.DateOfBirth);
			Assert.Equal("47 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("Asthma", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("23/10/2017", data.VisitID);
			Assert.Equal("Current Smoker", data.Smoker);
			Assert.Equal("31.25", data.PackYears);
			Assert.Equal("23.8", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("77.0 kg", data.Weight);
			Assert.Equal("SR", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("23/10/2017 9:44 AM", data.AmbientData[0].DateTime);
			Assert.Equal("21.1 \\'b0C   787.50 mmHg   37 %", data.AmbientData[0].Ambient);
			Assert.Equal("23/10/2017 9:40 AM", data.AmbientData[1].DateTime);
			Assert.Equal("20.9 \\'b0C   787.36 mmHg   37 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.13", data.SpirometryData[0][0]);
			Assert.Equal("2.34", data.SpirometryData[0][1]);
			Assert.Equal("57 %", data.SpirometryData[0][2]);
			Assert.Equal("2.45", data.SpirometryData[0][3]);
			Assert.Equal("59 %", data.SpirometryData[0][4]);
			Assert.Equal("5 %", data.SpirometryData[0][5]);
			Assert.Equal("5.23", data.SpirometryData[1][0]);
			Assert.Equal("3.68", data.SpirometryData[1][1]);
			Assert.Equal("70 %", data.SpirometryData[1][2]);
			Assert.Equal("3.61", data.SpirometryData[1][3]);
			Assert.Equal("69 %", data.SpirometryData[1][4]);
			Assert.Equal("-2 %", data.SpirometryData[1][5]);
			Assert.Equal(" 79.40", data.SpirometryData[2][0]);
			Assert.Equal("63.59", data.SpirometryData[2][1]);
			Assert.Equal("80 %", data.SpirometryData[2][2]);
			Assert.Equal("68.01", data.SpirometryData[2][3]);
			Assert.Equal("86 %", data.SpirometryData[2][4]);
			Assert.Equal("7 %", data.SpirometryData[2][5]);
			Assert.Equal("3.82", data.SpirometryData[3][0]);
			Assert.Equal("1.46", data.SpirometryData[3][1]);
			Assert.Equal("38 %", data.SpirometryData[3][2]);
			Assert.Equal("1.49", data.SpirometryData[3][3]);
			Assert.Equal("39 %", data.SpirometryData[3][4]);
			Assert.Equal("2 %", data.SpirometryData[3][5]);
			Assert.Equal("5.02", data.SpirometryData[4][0]);
			Assert.Equal("1.91", data.SpirometryData[4][1]);
			Assert.Equal("38 %", data.SpirometryData[4][2]);
			Assert.Equal("1.53", data.SpirometryData[4][3]);
			Assert.Equal("30 %", data.SpirometryData[4][4]);
			Assert.Equal("-20 %", data.SpirometryData[4][5]);
			Assert.Equal("9.18", data.SpirometryData[5][0]);
			Assert.Equal("6.59", data.SpirometryData[5][1]);
			Assert.Equal("72 %", data.SpirometryData[5][2]);
			Assert.Equal("5.29", data.SpirometryData[5][3]);
			Assert.Equal("58 %", data.SpirometryData[5][4]);
			Assert.Equal("-20 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("32.45", data.DiffusionData[0][0]);
			Assert.Equal("27.68", data.DiffusionData[0][1]);
			Assert.Equal("85 %", data.DiffusionData[0][2]);
			Assert.Equal("32.45", data.DiffusionData[1][0]);
			Assert.Equal("27.68", data.DiffusionData[1][1]);
			Assert.Equal("85 %", data.DiffusionData[1][2]);
			Assert.Equal("4.44", data.DiffusionData[2][0]);
			Assert.Equal("4.67", data.DiffusionData[2][1]);
			Assert.Equal("105 %", data.DiffusionData[2][2]);
			Assert.Equal("4.44", data.DiffusionData[3][0]);
			Assert.Equal("4.67", data.DiffusionData[3][1]);
			Assert.Equal("105 %", data.DiffusionData[3][2]);
			Assert.Equal("7.15", data.DiffusionData[4][0]);
			Assert.Equal("5.93", data.DiffusionData[4][1]);
			Assert.Equal("83 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("Patient mentioned he is a smoker, smoked for 25-30 years, last use 9 hours prior to test. Patient also mentioned he is on Asmol, Seretide and Seebreeze, last use for all more than 12 hours prior to test.Patient struggled slightly with the technique during the test but was able to perform to the best of his ability after further explanation and demonstration.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 9:30 AM: Spirometry results indicate a restrictive ventilatory defect which was mild in severity. /*/* Automatic Interpretation - Forced Spirometry - 9:40 AM: Spirometry results indicate a restrictive ventilatory defect which was moderate in severity.There is no significant change following inhaled bronchodilator on this occasion. /*/* Automatic Interpretation - Single Breath Diffusion - 9:44 AM: CO Diffusion was within normal limits. /*", data.Interpretation);
			#endregion
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-23-10 Kirk ONEILL.rtf"),
				File.ReadAllText(temp));
		}

		[Fact]
		public void Test9()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-23-10_10 John MCVEA.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPre, data.DiffusionType);
			Assert.Equal("MCVEA_J23102017", data.PatientId);
			Assert.Equal("MCVEA", data.LastName);
			Assert.Equal("John", data.FirstName);
			Assert.Equal("3/01/1943", data.DateOfBirth);
			Assert.Equal("74 years", data.Age);
			Assert.Equal("172.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("23/10/2017", data.VisitID);
			Assert.Equal("Ex Smoker", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("25.4", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("75.0 kg", data.Weight);
			Assert.Equal("SR", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("23/10/2017 10:38 AM", data.AmbientData[0].DateTime);
			Assert.Equal("23.6 \\'b0C   787.12 mmHg   37 %", data.AmbientData[0].Ambient);
			Assert.Equal("23/10/2017 10:35 AM", data.AmbientData[1].DateTime);
			Assert.Equal("23.5 \\'b0C   787.41 mmHg   37 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("2.86", data.SpirometryData[0][0]);
			Assert.Equal("2.89", data.SpirometryData[0][1]);
			Assert.Equal("101 %", data.SpirometryData[0][2]);
			Assert.Equal("2.29", data.SpirometryData[0][3]);
			Assert.Equal("80 %", data.SpirometryData[0][4]);
			Assert.Equal("-21 %", data.SpirometryData[0][5]);
			Assert.Equal("3.79", data.SpirometryData[1][0]);
			Assert.Equal("3.75", data.SpirometryData[1][1]);
			Assert.Equal("99 %", data.SpirometryData[1][2]);
			Assert.Equal("3.69", data.SpirometryData[1][3]);
			Assert.Equal("97 %", data.SpirometryData[1][4]);
			Assert.Equal("-2 %", data.SpirometryData[1][5]);
			Assert.Equal(" 75.82", data.SpirometryData[2][0]);
			Assert.Equal("77.21", data.SpirometryData[2][1]);
			Assert.Equal("102 %", data.SpirometryData[2][2]);
			Assert.Equal("62.00", data.SpirometryData[2][3]);
			Assert.Equal("82 %", data.SpirometryData[2][4]);
			Assert.Equal("-20 %", data.SpirometryData[2][5]);
			Assert.Equal("2.14", data.SpirometryData[3][0]);
			Assert.Equal("2.66", data.SpirometryData[3][1]);
			Assert.Equal("124 %", data.SpirometryData[3][2]);
			Assert.Equal("1.81", data.SpirometryData[3][3]);
			Assert.Equal("84 %", data.SpirometryData[3][4]);
			Assert.Equal("-32 %", data.SpirometryData[3][5]);
			Assert.Equal("3.87", data.SpirometryData[4][0]);
			Assert.Equal("2.97", data.SpirometryData[4][1]);
			Assert.Equal("77 %", data.SpirometryData[4][2]);
			Assert.Equal("1.56", data.SpirometryData[4][3]);
			Assert.Equal("40 %", data.SpirometryData[4][4]);
			Assert.Equal("-48 %", data.SpirometryData[4][5]);
			Assert.Equal("7.53", data.SpirometryData[5][0]);
			Assert.Equal("4.44", data.SpirometryData[5][1]);
			Assert.Equal("59 %", data.SpirometryData[5][2]);
			Assert.Equal("3.49", data.SpirometryData[5][3]);
			Assert.Equal("46 %", data.SpirometryData[5][4]);
			Assert.Equal("-21 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("24.47", data.DiffusionData[0][0]);
			Assert.Equal("18.02", data.DiffusionData[0][1]);
			Assert.Equal("74 %", data.DiffusionData[0][2]);
			Assert.Equal("24.47", data.DiffusionData[1][0]);
			Assert.Equal("18.02", data.DiffusionData[1][1]);
			Assert.Equal("74 %", data.DiffusionData[1][2]);
			Assert.Equal("3.67", data.DiffusionData[2][0]);
			Assert.Equal("3.22", data.DiffusionData[2][1]);
			Assert.Equal("88 %", data.DiffusionData[2][2]);
			Assert.Equal("3.67", data.DiffusionData[3][0]);
			Assert.Equal("3.22", data.DiffusionData[3][1]);
			Assert.Equal("88 %", data.DiffusionData[3][2]);
			Assert.Equal("6.51", data.DiffusionData[4][0]);
			Assert.Equal("5.60", data.DiffusionData[4][1]);
			Assert.Equal("86 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("Patient mentioned he is an ex-smoker, smoked for approximately 15 years, quit 45 years ago. Patient also mentioned he is not on any medications.Patient did not have any technical difficulty during the test. Result obtained were to the patients best ability.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 10:26 AM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Forced Spirometry - 10:35 AM: Spirometry results are within normal limits.There is no significant change following inhaled bronchodilator on this occasion. /*/* Automatic Interpretation - Single Breath Diffusion - 10:38 AM: Decreased DLCO, normal KCO; indicating normal alveolocapillary membrane (small lung syndrome, bullous emphysema) and indicative of lung parenchymal and/or pulmonary vascular dysfunction. /*", data.Interpretation);
			#endregion
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-23-10_10 John MCVEA.rtf"),
				File.ReadAllText(temp));
		}

		[Fact]
		public void Test10()
		{
			var data = PdfParser.Parse(
				@"..\..\..\SampleData\LFX-Report - 2017-31-8_11-34-15 - Heartscope3Report Margaret RADFORD.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.SixColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.RefPost, data.DiffusionType);
			Assert.Equal("RADFORD_M31082017", data.PatientId);
			Assert.Equal("RADFORD", data.LastName);
			Assert.Equal("Margaret", data.FirstName);
			Assert.Equal("1/24/1949", data.DateOfBirth);
			Assert.Equal("68 years", data.Age);
			Assert.Equal("167.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("Dr J Thomas", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("31/08/2017", data.VisitID);
			Assert.Equal("Ex Smoker", data.Smoker);
			Assert.Equal("22.00", data.PackYears);
			Assert.Equal("36.9", data.BMI);
			Assert.Equal("Female", data.Gender);
			Assert.Equal("103.0 kg", data.Weight);
			Assert.Equal("MT", data.Technician);
			Assert.Equal("Heartscope Victoria", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("8/31/2017 11:11 AM", data.AmbientData[0].DateTime);
			Assert.Equal("25.8 \\'b0C   795.36 mmHg   37 %", data.AmbientData[0].Ambient);
			Assert.Equal("8/31/2017 11:27 AM", data.AmbientData[1].DateTime);
			Assert.Equal("26.2 \\'b0C   795.01 mmHg   37 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("2.40", data.SpirometryData[0][0]);
			Assert.Equal("2.72", data.SpirometryData[0][1]);
			Assert.Equal("113 %", data.SpirometryData[0][2]);
			Assert.Equal("2.86", data.SpirometryData[0][3]);
			Assert.Equal("119 %", data.SpirometryData[0][4]);
			Assert.Equal("5 %", data.SpirometryData[0][5]);
			Assert.Equal("3.10", data.SpirometryData[1][0]);
			Assert.Equal("3.47", data.SpirometryData[1][1]);
			Assert.Equal("112 %", data.SpirometryData[1][2]);
			Assert.Equal("3.58", data.SpirometryData[1][3]);
			Assert.Equal("116 %", data.SpirometryData[1][4]);
			Assert.Equal("3 %", data.SpirometryData[1][5]);
			Assert.Equal(" 76.18", data.SpirometryData[2][0]);
			Assert.Equal("77.19", data.SpirometryData[2][1]);
			Assert.Equal("101 %", data.SpirometryData[2][2]);
			Assert.Equal("84.31", data.SpirometryData[2][3]);
			Assert.Equal("111 %", data.SpirometryData[2][4]);
			Assert.Equal("9 %", data.SpirometryData[2][5]);
			Assert.Equal("2.02", data.SpirometryData[3][0]);
			Assert.Equal("2.49", data.SpirometryData[3][1]);
			Assert.Equal("123 %", data.SpirometryData[3][2]);
			Assert.Equal("2.76", data.SpirometryData[3][3]);
			Assert.Equal("137 %", data.SpirometryData[3][4]);
			Assert.Equal("11 %", data.SpirometryData[3][5]);
			Assert.Equal("3.55", data.SpirometryData[4][0]);
			Assert.Equal("3.46", data.SpirometryData[4][1]);
			Assert.Equal("97 %", data.SpirometryData[4][2]);
			Assert.Equal("4.03", data.SpirometryData[4][3]);
			Assert.Equal("114 %", data.SpirometryData[4][4]);
			Assert.Equal("17 %", data.SpirometryData[4][5]);
			Assert.Equal("6.04", data.SpirometryData[5][0]);
			Assert.Equal("7.25", data.SpirometryData[5][1]);
			Assert.Equal("120 %", data.SpirometryData[5][2]);
			Assert.Equal("8.45", data.SpirometryData[5][3]);
			Assert.Equal("140 %", data.SpirometryData[5][4]);
			Assert.Equal("17 %", data.SpirometryData[5][5]);
			Assert.Equal(6, data.DiffusionData.Length);
			Assert.Equal("22.66", data.DiffusionData[0][0]);
			Assert.Equal("26.92", data.DiffusionData[0][1]);
			Assert.Equal("119 %", data.DiffusionData[0][2]);
			Assert.Equal("22.66", data.DiffusionData[1][0]);
			Assert.Equal("26.92", data.DiffusionData[1][1]);
			Assert.Equal("119 %", data.DiffusionData[1][2]);
			Assert.Equal("4.33", data.DiffusionData[2][0]);
			Assert.Equal("5.76", data.DiffusionData[2][1]);
			Assert.Equal("133 %", data.DiffusionData[2][2]);
			Assert.Equal("4.33", data.DiffusionData[3][0]);
			Assert.Equal("5.76", data.DiffusionData[3][1]);
			Assert.Equal("133 %", data.DiffusionData[3][2]);
			Assert.Equal("5.08", data.DiffusionData[4][0]);
			Assert.Equal("4.67", data.DiffusionData[4][1]);
			Assert.Equal("92 %", data.DiffusionData[4][2]);
			Assert.Equal("- ", data.DiffusionData[5][0]);
			Assert.Equal("- ", data.DiffusionData[5][1]);
			Assert.Equal("-", data.DiffusionData[5][2]);
			Assert.Equal("The patient had no technical difficulties performing the test. Please note 300 micrograms of Ventolin was given for post forced spirometry.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 11:11 AM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Forced Spirometry - 11:21 AM: Spirometry results are within normal limits.There is no significant change following inhaled bronchodilator on this occasion. /*/* Automatic Interpretation - Single Breath Diffusion - 11:27 AM: CO Diffusion was within normal limits. /*", data.Interpretation);
			#endregion
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(
				File.ReadAllText(
					@"..\..\ReferenceRtfs\LFX-Report - 2017-31-8_11-34-15 - Heartscope3Report Margaret RADFORD.rtf"),
				File.ReadAllText(temp));
		}
	}
}