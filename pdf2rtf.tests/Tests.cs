using System;
using System.IO;
using System.Text;
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
			var data = PdfParser.Parse(@"..\..\..\SampleData\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_17-10.pdf");
			#region TestData
			Assert.Equal(AmbientType.Measured, data.AmbientType);
			Assert.Equal(SpirometryType.FiveColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.None, data.DiffusionType);
			Assert.Equal("_Ganshorn_TP", data.PatientId);
			Assert.Equal("_Muster", data.LastName);
			Assert.Equal("Patient", data.FirstName);
			Assert.Equal("7/03/1986", data.DateOfBirth);
			Assert.Equal("32 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("23.1", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("75.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(1, data.AmbientData.Length);
			Assert.Equal("12/04/2016 5:10 PM", data.AmbientData[0].DateTime);
			Assert.Equal("19.8 \\'b0C   756.01 mmHg   50 %", data.AmbientData[0].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.61", data.SpirometryData[0][0]);
			Assert.Equal("4.78", data.SpirometryData[0][1]);
			Assert.Equal("104 %", data.SpirometryData[0][2]);
			Assert.Equal("3.68", data.SpirometryData[0][3]);
			Assert.Equal("0.3", data.SpirometryData[0][4]);
			Assert.Equal("5.61", data.SpirometryData[1][0]);
			Assert.Equal("6.02", data.SpirometryData[1][1]);
			Assert.Equal("107 %", data.SpirometryData[1][2]);
			Assert.Equal("4.52", data.SpirometryData[1][3]);
			Assert.Equal("0.6", data.SpirometryData[1][4]);
			Assert.Equal(" 81.81", data.SpirometryData[2][0]);
			Assert.Equal("78.18", data.SpirometryData[2][1]);
			Assert.Equal("96 %", data.SpirometryData[2][2]);
			Assert.Equal("70.05", data.SpirometryData[2][3]);
			Assert.Equal("-0.5", data.SpirometryData[2][4]);
			Assert.Equal("4.66", data.SpirometryData[3][0]);
			Assert.Equal("4.33", data.SpirometryData[3][1]);
			Assert.Equal("93 %", data.SpirometryData[3][2]);
			Assert.Equal("2.90", data.SpirometryData[3][3]);
			Assert.Equal("-0.3", data.SpirometryData[3][4]);
			Assert.Equal("5.54", data.SpirometryData[4][0]);
			Assert.Equal("4.86", data.SpirometryData[4][1]);
			Assert.Equal("88 %", data.SpirometryData[4][2]);
			Assert.Equal("3.38", data.SpirometryData[4][3]);
			Assert.Equal("-0.5", data.SpirometryData[4][4]);
			Assert.Equal("9.91", data.SpirometryData[5][0]);
			Assert.Equal("11.43", data.SpirometryData[5][1]);
			Assert.Equal("115 %", data.SpirometryData[5][2]);
			Assert.Equal("7.93", data.SpirometryData[5][3]);
			Assert.Equal("1.3", data.SpirometryData[5][4]);
			Assert.Null(data.DiffusionData);
			Assert.Equal("/* Automatic trial quality determination - Single Breath Diffusion - 8:42 PM: Trial2: Score Undefined; Trial1: Score Undefined /*", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 5:10 PM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Single Breath Diffusion - 8:42 PM: Increased DLCO indicating Polycythemia, Asthma, increased cardiac output or capillary blood volume. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_17-10.rtf"), File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test2()
		{
			var data = PdfParser.Parse(@"..\..\..\SampleData\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_17-26.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.EightColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.None, data.DiffusionType);
			Assert.Equal("_Ganshorn_TP", data.PatientId);
			Assert.Equal("_Muster", data.LastName);
			Assert.Equal("Patient", data.FirstName);
			Assert.Equal("7/03/1986", data.DateOfBirth);
			Assert.Equal("32 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("23.1", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("75.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("12/04/2016 5:10 PM", data.AmbientData[0].DateTime);
			Assert.Equal("19.8 \\'b0C   756.01 mmHg   50 %", data.AmbientData[0].Ambient);
			Assert.Equal("12/04/2016 5:26 PM", data.AmbientData[1].DateTime);
			Assert.Equal("21 \\'b0C   755.46 mmHg   50 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.61", data.SpirometryData[0][0]);
			Assert.Equal("4.78", data.SpirometryData[0][1]);
			Assert.Equal("104 %", data.SpirometryData[0][2]);
			Assert.Equal("3.68", data.SpirometryData[0][3]);
			Assert.Equal("0.2", data.SpirometryData[0][4]);
			Assert.Equal("4.73", data.SpirometryData[0][5]);
			Assert.Equal("103 %", data.SpirometryData[0][6]);
			Assert.Equal("-1 %", data.SpirometryData[0][7]);
			Assert.Equal("5.61", data.SpirometryData[1][0]);
			Assert.Equal("6.02", data.SpirometryData[1][1]);
			Assert.Equal("107 %", data.SpirometryData[1][2]);
			Assert.Equal("4.52", data.SpirometryData[1][3]);
			Assert.Equal("0.5", data.SpirometryData[1][4]);
			Assert.Equal("5.93", data.SpirometryData[1][5]);
			Assert.Equal("106 %", data.SpirometryData[1][6]);
			Assert.Equal("-1 %", data.SpirometryData[1][7]);
			Assert.Equal(" 81.81", data.SpirometryData[2][0]);
			Assert.Equal("78.18", data.SpirometryData[2][1]);
			Assert.Equal("96 %", data.SpirometryData[2][2]);
			Assert.Equal("70.05", data.SpirometryData[2][3]);
			Assert.Equal("-0.5", data.SpirometryData[2][4]);
			Assert.Equal("77.98", data.SpirometryData[2][5]);
			Assert.Equal("95 %", data.SpirometryData[2][6]);
			Assert.Equal("0 %", data.SpirometryData[2][7]);
			Assert.Equal("4.66", data.SpirometryData[3][0]);
			Assert.Equal("4.33", data.SpirometryData[3][1]);
			Assert.Equal("93 %", data.SpirometryData[3][2]);
			Assert.Equal("2.90", data.SpirometryData[3][3]);
			Assert.Equal("-0.3", data.SpirometryData[3][4]);
			Assert.Equal("4.33", data.SpirometryData[3][5]);
			Assert.Equal("93 %", data.SpirometryData[3][6]);
			Assert.Equal("0 %", data.SpirometryData[3][7]);
			Assert.Equal("5.54", data.SpirometryData[4][0]);
			Assert.Equal("4.86", data.SpirometryData[4][1]);
			Assert.Equal("88 %", data.SpirometryData[4][2]);
			Assert.Equal("3.38", data.SpirometryData[4][3]);
			Assert.Equal("-0.5", data.SpirometryData[4][4]);
			Assert.Equal("4.90", data.SpirometryData[4][5]);
			Assert.Equal("88 %", data.SpirometryData[4][6]);
			Assert.Equal("1 %", data.SpirometryData[4][7]);
			Assert.Equal("9.91", data.SpirometryData[5][0]);
			Assert.Equal("11.43", data.SpirometryData[5][1]);
			Assert.Equal("115 %", data.SpirometryData[5][2]);
			Assert.Equal("7.93", data.SpirometryData[5][3]);
			Assert.Equal("1.2", data.SpirometryData[5][4]);
			Assert.Equal("11.42", data.SpirometryData[5][5]);
			Assert.Equal("115 %", data.SpirometryData[5][6]);
			Assert.Equal("0 %", data.SpirometryData[5][7]);
			Assert.Null(data.DiffusionData);
			Assert.Equal("/* Automatic trial quality determination - Single Breath Diffusion - 8:42 PM: Trial2: Score Undefined; Trial1: Score Undefined /*", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 5:10 PM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Single Breath Diffusion - 8:42 PM: Increased DLCO indicating Polycythemia, Asthma, increased cardiac output or capillary blood volume. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_17-26.rtf"), File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test3()
		{
			var data = PdfParser.Parse(@"..\..\..\SampleData\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_20-42.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.EightColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.Present, data.DiffusionType);
			Assert.Equal("_Ganshorn_TP", data.PatientId);
			Assert.Equal("_Muster", data.LastName);
			Assert.Equal("Patient", data.FirstName);
			Assert.Equal("7/03/1986", data.DateOfBirth);
			Assert.Equal("32 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
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
			Assert.Equal("21.3 \\'b0C   755.16 mmHg   50 %", data.AmbientData[0].Ambient);
			Assert.Equal("12/04/2016 5:26 PM", data.AmbientData[1].DateTime);
			Assert.Equal("21 \\'b0C   755.46 mmHg   50 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.61", data.SpirometryData[0][0]);
			Assert.Equal("4.78", data.SpirometryData[0][1]);
			Assert.Equal("104 %", data.SpirometryData[0][2]);
			Assert.Equal("3.68", data.SpirometryData[0][3]);
			Assert.Equal("0.2", data.SpirometryData[0][4]);
			Assert.Equal("4.73", data.SpirometryData[0][5]);
			Assert.Equal("103 %", data.SpirometryData[0][6]);
			Assert.Equal("-1 %", data.SpirometryData[0][7]);
			Assert.Equal("5.61", data.SpirometryData[1][0]);
			Assert.Equal("6.02", data.SpirometryData[1][1]);
			Assert.Equal("107 %", data.SpirometryData[1][2]);
			Assert.Equal("4.52", data.SpirometryData[1][3]);
			Assert.Equal("0.5", data.SpirometryData[1][4]);
			Assert.Equal("5.93", data.SpirometryData[1][5]);
			Assert.Equal("106 %", data.SpirometryData[1][6]);
			Assert.Equal("-1 %", data.SpirometryData[1][7]);
			Assert.Equal(" 81.81", data.SpirometryData[2][0]);
			Assert.Equal("78.18", data.SpirometryData[2][1]);
			Assert.Equal("96 %", data.SpirometryData[2][2]);
			Assert.Equal("70.05", data.SpirometryData[2][3]);
			Assert.Equal("-0.5", data.SpirometryData[2][4]);
			Assert.Equal("77.98", data.SpirometryData[2][5]);
			Assert.Equal("95 %", data.SpirometryData[2][6]);
			Assert.Equal("0 %", data.SpirometryData[2][7]);
			Assert.Equal("4.66", data.SpirometryData[3][0]);
			Assert.Equal("4.33", data.SpirometryData[3][1]);
			Assert.Equal("93 %", data.SpirometryData[3][2]);
			Assert.Equal("2.90", data.SpirometryData[3][3]);
			Assert.Equal("-0.3", data.SpirometryData[3][4]);
			Assert.Equal("4.33", data.SpirometryData[3][5]);
			Assert.Equal("93 %", data.SpirometryData[3][6]);
			Assert.Equal("0 %", data.SpirometryData[3][7]);
			Assert.Equal("5.54", data.SpirometryData[4][0]);
			Assert.Equal("4.86", data.SpirometryData[4][1]);
			Assert.Equal("88 %", data.SpirometryData[4][2]);
			Assert.Equal("3.38", data.SpirometryData[4][3]);
			Assert.Equal("-0.5", data.SpirometryData[4][4]);
			Assert.Equal("4.90", data.SpirometryData[4][5]);
			Assert.Equal("88 %", data.SpirometryData[4][6]);
			Assert.Equal("1 %", data.SpirometryData[4][7]);
			Assert.Equal("9.91", data.SpirometryData[5][0]);
			Assert.Equal("11.43", data.SpirometryData[5][1]);
			Assert.Equal("115 %", data.SpirometryData[5][2]);
			Assert.Equal("7.93", data.SpirometryData[5][3]);
			Assert.Equal("1.2", data.SpirometryData[5][4]);
			Assert.Equal("11.42", data.SpirometryData[5][5]);
			Assert.Equal("115 %", data.SpirometryData[5][6]);
			Assert.Equal("0 %", data.SpirometryData[5][7]);
			Assert.Equal(3, data.DiffusionData.Length);
			Assert.Equal("33.55", data.DiffusionData[0][0]);
			Assert.Equal("63.84", data.DiffusionData[0][1]);
			Assert.Equal("190 %", data.DiffusionData[0][2]);
			Assert.Equal("26.43", data.DiffusionData[0][3]);
			Assert.Equal("5.3", data.DiffusionData[0][4]);
			Assert.Equal("4.93", data.DiffusionData[1][0]);
			Assert.Equal("8.74", data.DiffusionData[1][1]);
			Assert.Equal("177 %", data.DiffusionData[1][2]);
			Assert.Equal("3.92", data.DiffusionData[1][3]);
			Assert.Equal("5.3", data.DiffusionData[1][4]);
			Assert.Equal("6.85", data.DiffusionData[2][0]);
			Assert.Equal("7.31", data.DiffusionData[2][1]);
			Assert.Equal("107 %", data.DiffusionData[2][2]);
			Assert.Equal("5.62", data.DiffusionData[2][3]);
			Assert.Equal("0.6", data.DiffusionData[2][4]);
			Assert.Equal("/* Automatic trial quality determination - Single Breath Diffusion - 8:42 PM: Trial2: Score Undefined; Trial1: Score Undefined /*", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 5:10 PM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Single Breath Diffusion - 8:42 PM: Increased DLCO indicating Polycythemia, Asthma, increased cardiac output or capillary blood volume. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_20-42.rtf"), File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test4()
		{
			var data = PdfParser.Parse(@"..\..\..\SampleData\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_20-42_2.pdf");
			#region TestData
			Assert.Equal(AmbientType.Measured, data.AmbientType);
			Assert.Equal(SpirometryType.FiveColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.Present, data.DiffusionType);
			Assert.Equal("_Ganshorn_TP", data.PatientId);
			Assert.Equal("_Muster", data.LastName);
			Assert.Equal("Patient", data.FirstName);
			Assert.Equal("7/03/1986", data.DateOfBirth);
			Assert.Equal("32 years", data.Age);
			Assert.Equal("180.0 cm", data.Height);
			Assert.Equal("", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("23.1", data.BMI);
			Assert.Equal("Male", data.Gender);
			Assert.Equal("75.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(1, data.AmbientData.Length);
			Assert.Equal("12/04/2016 8:42 PM", data.AmbientData[0].DateTime);
			Assert.Equal("21.3 \\'b0C   755.16 mmHg   50 %", data.AmbientData[0].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("4.61", data.SpirometryData[0][0]);
			Assert.Equal("4.78", data.SpirometryData[0][1]);
			Assert.Equal("104 %", data.SpirometryData[0][2]);
			Assert.Equal("3.68", data.SpirometryData[0][3]);
			Assert.Equal("0.3", data.SpirometryData[0][4]);
			Assert.Equal("5.61", data.SpirometryData[1][0]);
			Assert.Equal("6.02", data.SpirometryData[1][1]);
			Assert.Equal("107 %", data.SpirometryData[1][2]);
			Assert.Equal("4.52", data.SpirometryData[1][3]);
			Assert.Equal("0.6", data.SpirometryData[1][4]);
			Assert.Equal(" 81.81", data.SpirometryData[2][0]);
			Assert.Equal("78.18", data.SpirometryData[2][1]);
			Assert.Equal("96 %", data.SpirometryData[2][2]);
			Assert.Equal("70.05", data.SpirometryData[2][3]);
			Assert.Equal("-0.5", data.SpirometryData[2][4]);
			Assert.Equal("4.66", data.SpirometryData[3][0]);
			Assert.Equal("4.33", data.SpirometryData[3][1]);
			Assert.Equal("93 %", data.SpirometryData[3][2]);
			Assert.Equal("2.90", data.SpirometryData[3][3]);
			Assert.Equal("-0.3", data.SpirometryData[3][4]);
			Assert.Equal("5.54", data.SpirometryData[4][0]);
			Assert.Equal("4.86", data.SpirometryData[4][1]);
			Assert.Equal("88 %", data.SpirometryData[4][2]);
			Assert.Equal("3.38", data.SpirometryData[4][3]);
			Assert.Equal("-0.5", data.SpirometryData[4][4]);
			Assert.Equal("9.91", data.SpirometryData[5][0]);
			Assert.Equal("11.43", data.SpirometryData[5][1]);
			Assert.Equal("115 %", data.SpirometryData[5][2]);
			Assert.Equal("7.93", data.SpirometryData[5][3]);
			Assert.Equal("1.3", data.SpirometryData[5][4]);
			Assert.Equal(3, data.DiffusionData.Length);
			Assert.Equal("33.55", data.DiffusionData[0][0]);
			Assert.Equal("63.84", data.DiffusionData[0][1]);
			Assert.Equal("190 %", data.DiffusionData[0][2]);
			Assert.Equal("26.43", data.DiffusionData[0][3]);
			Assert.Equal("5.3", data.DiffusionData[0][4]);
			Assert.Equal("4.93", data.DiffusionData[1][0]);
			Assert.Equal("8.74", data.DiffusionData[1][1]);
			Assert.Equal("177 %", data.DiffusionData[1][2]);
			Assert.Equal("3.92", data.DiffusionData[1][3]);
			Assert.Equal("5.3", data.DiffusionData[1][4]);
			Assert.Equal("6.85", data.DiffusionData[2][0]);
			Assert.Equal("7.31", data.DiffusionData[2][1]);
			Assert.Equal("107 %", data.DiffusionData[2][2]);
			Assert.Equal("5.62", data.DiffusionData[2][3]);
			Assert.Equal("0.6", data.DiffusionData[2][4]);
			Assert.Equal("/* Automatic trial quality determination - Single Breath Diffusion - 8:42 PM: Trial2: Score Undefined; Trial1: Score Undefined /*", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Forced Spirometry - 5:10 PM: Spirometry results are within normal limits. /*/* Automatic Interpretation - Single Breath Diffusion - 8:42 PM: Increased DLCO indicating Polycythemia, Asthma, increased cardiac output or capillary blood volume. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\_Muster__Ganshorn_TP_HeartscopeLFT_2016-12-4_20-42_2.rtf"), File.ReadAllText(temp));
			#endregion
		}

		[Fact]
		public void Test5()
		{
			var data = PdfParser.Parse(@"..\..\..\SampleData\Pattni_26072018_HeartscopeLFT_2018-26-7_14-58.pdf");
			#region TestData
			Assert.Equal(AmbientType.PrePost, data.AmbientType);
			Assert.Equal(SpirometryType.EightColumn, data.SpirometryType);
			Assert.Equal(DiffusionType.Present, data.DiffusionType);
			Assert.Equal("26072018", data.PatientId);
			Assert.Equal("Pattni", data.LastName);
			Assert.Equal("Priya", data.FirstName);
			Assert.Equal("4/02/1980", data.DateOfBirth);
			Assert.Equal("38 years", data.Age);
			Assert.Equal("168.0 cm", data.Height);
			Assert.Equal("COPD", data.History);
			Assert.Equal("", data.Physician);
			Assert.Equal("West", data.Insurance);
			Assert.Equal("", data.VisitID);
			Assert.Equal("", data.Smoker);
			Assert.Equal("", data.PackYears);
			Assert.Equal("28.0", data.BMI);
			Assert.Equal("Female", data.Gender);
			Assert.Equal("79.0 kg", data.Weight);
			Assert.Equal("", data.Technician);
			Assert.Equal("", data.Ward);
			Assert.Equal(2, data.AmbientData.Length);
			Assert.Equal("26/07/2018 2:21 PM", data.AmbientData[0].DateTime);
			Assert.Equal("19.4 \\'b0C   780.36 mmHg   40 %", data.AmbientData[0].Ambient);
			Assert.Equal("26/07/2018 2:58 PM", data.AmbientData[1].DateTime);
			Assert.Equal("20 \\'b0C   780.14 mmHg   40 %", data.AmbientData[1].Ambient);
			Assert.Equal(6, data.SpirometryData.Length);
			Assert.Equal("3.06", data.SpirometryData[0][0]);
			Assert.Equal("2.65", data.SpirometryData[0][1]);
			Assert.Equal("86 %", data.SpirometryData[0][2]);
			Assert.Equal("2.44", data.SpirometryData[0][3]);
			Assert.Equal("-1.5", data.SpirometryData[0][4]);
			Assert.Equal("2.51", data.SpirometryData[0][5]);
			Assert.Equal("82 %", data.SpirometryData[0][6]);
			Assert.Equal("-5 %", data.SpirometryData[0][7]);
			Assert.Equal("3.68", data.SpirometryData[1][0]);
			Assert.Equal("3.22", data.SpirometryData[1][1]);
			Assert.Equal("87 %", data.SpirometryData[1][2]);
			Assert.Equal("2.98", data.SpirometryData[1][3]);
			Assert.Equal("-1.2", data.SpirometryData[1][4]);
			Assert.Equal("3.15", data.SpirometryData[1][5]);
			Assert.Equal("85 %", data.SpirometryData[1][6]);
			Assert.Equal("-2 %", data.SpirometryData[1][7]);
			Assert.Equal(" 81.88", data.SpirometryData[2][0]);
			Assert.Equal("79.04", data.SpirometryData[2][1]);
			Assert.Equal("97 %", data.SpirometryData[2][2]);
			Assert.Equal("71.20", data.SpirometryData[2][3]);
			Assert.Equal("-0.3", data.SpirometryData[2][4]);
			Assert.Equal("79.68", data.SpirometryData[2][5]);
			Assert.Equal("97 %", data.SpirometryData[2][6]);
			Assert.Equal("1 %", data.SpirometryData[2][7]);
			Assert.Equal("3.24", data.SpirometryData[3][0]);
			Assert.Equal("2.65", data.SpirometryData[3][1]);
			Assert.Equal("82 %", data.SpirometryData[3][2]);
			Assert.Equal("1.99", data.SpirometryData[3][3]);
			Assert.Equal("-1.1", data.SpirometryData[3][4]);
			Assert.Equal("2.37", data.SpirometryData[3][5]);
			Assert.Equal("73 %", data.SpirometryData[3][6]);
			Assert.Equal("-11 %", data.SpirometryData[3][7]);
			Assert.Equal("4.33", data.SpirometryData[4][0]);
			Assert.Equal("3.60", data.SpirometryData[4][1]);
			Assert.Equal("83 %", data.SpirometryData[4][2]);
			Assert.Equal("2.52", data.SpirometryData[4][3]);
			Assert.Equal("-1.0", data.SpirometryData[4][4]);
			Assert.Equal("3.25", data.SpirometryData[4][5]);
			Assert.Equal("75 %", data.SpirometryData[4][6]);
			Assert.Equal("-10 %", data.SpirometryData[4][7]);
			Assert.Equal("6.99", data.SpirometryData[5][0]);
			Assert.Equal("7.06", data.SpirometryData[5][1]);
			Assert.Equal("101 %", data.SpirometryData[5][2]);
			Assert.Equal("5.51", data.SpirometryData[5][3]);
			Assert.Equal("0.5", data.SpirometryData[5][4]);
			Assert.Equal("7.47", data.SpirometryData[5][5]);
			Assert.Equal("107 %", data.SpirometryData[5][6]);
			Assert.Equal("6 %", data.SpirometryData[5][7]);
			Assert.Equal(3, data.DiffusionData.Length);
			Assert.Equal("27.29", data.DiffusionData[0][0]);
			Assert.Equal("22.60", data.DiffusionData[0][1]);
			Assert.Equal("83 %", data.DiffusionData[0][2]);
			Assert.Equal("21.56", data.DiffusionData[0][3]);
			Assert.Equal("-1.3", data.DiffusionData[0][4]);
			Assert.Equal("5.15", data.DiffusionData[1][0]);
			Assert.Equal("5.07", data.DiffusionData[1][1]);
			Assert.Equal("98 %", data.DiffusionData[1][2]);
			Assert.Equal("3.70", data.DiffusionData[1][3]);
			Assert.Equal("-0.1", data.DiffusionData[1][4]);
			Assert.Equal("5.15", data.DiffusionData[2][0]);
			Assert.Equal("4.45", data.DiffusionData[2][1]);
			Assert.Equal("87 %", data.DiffusionData[2][2]);
			Assert.Equal("5.15", data.DiffusionData[2][3]);
			Assert.Equal("", data.DiffusionData[2][4]);
			Assert.Equal("Good patient compliance.", data.TechnicianNotes);
			Assert.Equal("/* Automatic Interpretation - Single Breath Diffusion - 2:21 PM: CO Diffusion was within normal limits. /*/* Automatic Interpretation - Forced Spirometry - 2:58 PM: Spirometry results are within normal limits.There is no significant change following inhaled bronchodilator on this occasion. /*", data.Interpretation);
			#endregion
			#region TestRtf
			var temp = Path.GetTempFileName();
			RtfExporter.Export(data, temp);
			Assert.Equal(File.ReadAllText(@"..\..\ReferenceRtfs\Pattni_26072018_HeartscopeLFT_2018-26-7_14-58.rtf"), File.ReadAllText(temp));
			#endregion
		}

		private void GenerateCode(ReportData data)
		{
			var code = new StringBuilder();
			foreach (var prop in data.GetType().GetProperties())
			{
				switch (prop.GetValue(data))
				{
					case null:
						code.AppendLine($"Assert.Null(data.{prop.Name});");
						break;
					case Array array when array.Length > 0:
						code.AppendLine($"Assert.Equal({array.Length}, data.{prop.Name}.Length);");
						for (var i = 0; i < array.Length; i++)
						{
							switch (array.GetValue(i))
							{
								case null:
									code.AppendLine($"Assert.Null(data.{prop.Name}[{i}]);");
									break;
								case string str:
									code.AppendLine($"Assert.Equal(\"{str}\", data.{prop.Name}[{i}]);");
									break;
								case Array innerArr when innerArr.Length > 1:
									for (var j = 0; j < innerArr.Length; j++)
										code.AppendLine(
											$"Assert.Equal(\"{innerArr.GetValue(j)}\", data.{prop.Name}[{i}][{j}]);");
									break;
								case Array innerArr:
									code.AppendLine($"Assert.Equal(0, data.{prop.Name}[{i}].Length);");
									break;
								case object val:
									foreach (var arrProp in val.GetType().GetProperties())
										code.AppendLine(
											$"Assert.Equal(\"{arrProp.GetValue(val)}\", data.{prop.Name}[{i}].{arrProp.Name});");
									break;
							}
						}

						break;
					case Array _:
						code.AppendLine($"Assert.Equal(0, data.{prop.Name}.Length);");
						break;
					case Enum enm:
						code.AppendLine($"Assert.Equal({enm.GetType().Name}.{enm}, data.{prop.Name});");
						break;
					case string str:
						code.AppendLine($"Assert.Equal(\"{str}\", data.{prop.Name});");
						break;
					default:
						code.AppendLine($"{prop.Name} ???");
						break;
				}
			}

			WriteLine(code);
		}
	}
}