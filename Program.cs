using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace pdf2rtf
{
    class Program
    {
        static TextCapture[] simpleCaptures =
        {
            new TextCapture("Patient Id (.+) Visit ID (.+)", nameof(ReportData.PatientId), nameof(ReportData.VisitID)),
            new TextCapture("Last name (.+) Smoker (.+)", nameof(ReportData.LastName), nameof(ReportData.Smoker)),
            new TextCapture("First name (.+) Pack years (.+)", nameof(ReportData.FirstName), nameof(ReportData.PackYears)),
            new TextCapture("Date of birth (.+) BMI (.+)", nameof(ReportData.DateOfBirth), nameof(ReportData.BMI)),
            new TextCapture("Age (.+) Gender (.+)", nameof(ReportData.Age), nameof(ReportData.Gender)),
            new TextCapture("Height (.+) Weight (.+)", nameof(ReportData.Height), nameof(ReportData.Weight)),
            new TextCapture("History (.+)", nameof(ReportData.History)),
            new TextCapture("Physician (.+) Ward (.+)", nameof(ReportData.Physician), nameof(ReportData.Ward)),
            new TextCapture("Insurance (.+)", nameof(ReportData.Insurance)),
        };

        static TextCapture[] ambientCaptures =
        {
            new TextCapture("Pre: (.+) Ambient: (.+%)", nameof(ReportData.AmbientPre)),
            new TextCapture("Post: (.+) Ambient: (.+%)", nameof(ReportData.AmbientPost)),
        };

        static TextCapture[] firstCaptures =
        {
            new TextCapture(@"FEV 1 \[L\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)", nameof(ReportData.FEV1)),
            new TextCapture(@"FVC \[L\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)", nameof(ReportData.FVC)),
            new TextCapture(@"FEV1%I \[%\](.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)", nameof(ReportData.FEV1I)),
            new TextCapture(@"MMEF \[L/s\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)", nameof(ReportData.MMEF)),
            new TextCapture(@"MEF 50 \[L/s\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)", nameof(ReportData.MEF50)),
            new TextCapture(@"PEF \[L/s\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)", nameof(ReportData.PEF)),
        };

        static TextCapture[] secondCaptures =
        {
            new TextCapture(@"DLCO \[ml/min/mmHg\] (.+) (.+) (.+%|-)", nameof(ReportData.DLCO)),
            new TextCapture(@"DLCOc \[ml/min/mmHg\] (.+) (.+) (.+%|-)", nameof(ReportData.DLCOc)),
            new TextCapture(@"KCO \[ml/min/mmHg/L\] (.+) (.+) (.+%|-)", nameof(ReportData.KCO)),
            new TextCapture(@"KCOc \[ml/min/mmHg/L\] (.+) (.+) (.+%|-)", nameof(ReportData.KCOc)),
            new TextCapture(@"VA \[L\] (.+) (.+) (.+%|-)", nameof(ReportData.VA)),
            new TextCapture(@"Hb \[mmol/L\] (.+) (.+) (.+%|-)", nameof(ReportData.Hb)),
        };

        static void Main(string[] args)
        {
            var reader = new PdfReader(@"D:\000\LFX-Report - 2017-14-8_16-26-12 - Heartscope3Report.pdf");
            var text = PdfTextExtractor.GetTextFromPage(reader, 1);

            var data = new ReportData();
            var lines = text.Split('\n').ToList();

            foreach (var capture in simpleCaptures)
            {
                var match = lines.Select(l => capture.Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    for (int i = 0; i < capture.Properties.Length; i++)
                    {
                        var prop = typeof(ReportData).GetProperty(capture.Properties[i]);
                        prop.SetValue(data, match.Groups[i + 1].Value);
                    }
                    lines.Remove(match.Value);
                }
            }

            foreach (var capture in ambientCaptures)
            {
                var match = lines.Select(l => capture.Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    var ambientData = new FirstData();

                    var prop = typeof(ReportData).GetProperty(capture.Properties[0]);
                    prop.SetValue(data, new AmbientData
                    {
                        DateTime = match.Groups[1].Value,
                        Ambient = match.Groups[2].Value
                    });
                    lines.Remove(match.Value);
                }
            }

            SetPropertyValues<FirstData>(data, lines, firstCaptures);
            SetPropertyValues<SecondData>(data, lines, secondCaptures);

            Console.WriteLine(JsonConvert.SerializeObject(data, new JsonSerializerSettings { Formatting = Formatting.Indented, NullValueHandling = NullValueHandling.Ignore }));
            Console.Read();
        }

        private static void SetPropertyValues<TInner>(ReportData data, List<string> lines, TextCapture[] captures) where TInner : new()
        {
            foreach (var capture in captures)
            {
                var match = lines.Select(l => capture.Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    var innerData = new TInner();
                    var props = typeof(TInner).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    for (int i = 0; i < props.Length; i++)
                    {
                        props[i].SetValue(innerData, match.Groups[i + 1].Value);
                    }

                    var prop = typeof(ReportData).GetProperty(capture.Properties[0]);
                    prop.SetValue(data, innerData);

                    lines.Remove(match.Value);
                }
            }
        }
    }
}
