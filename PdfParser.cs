using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace pdf2rtf
{
    class PdfParser
    {
        static TextCapture[] _simpleCaptures =
        {
            new TextCapture("Patient Id *(.*) Visit ID *(.*)", nameof(ReportData.PatientId), nameof(ReportData.VisitID)),
            new TextCapture("Last name *(.*?) Smoker *(.*)", nameof(ReportData.LastName), nameof(ReportData.Smoker)),
            new TextCapture("First name *(.*) Pack years *(.*)", nameof(ReportData.FirstName), nameof(ReportData.PackYears)),
            new TextCapture("Date of birth *(.*) BMI *(.*)", nameof(ReportData.DateOfBirth), nameof(ReportData.BMI)),
            new TextCapture("Age *(.*) Gender *(.*)", nameof(ReportData.Age), nameof(ReportData.Gender)),
            new TextCapture("Height *(.*) Weight *(.*)", nameof(ReportData.Height), nameof(ReportData.Weight)),
            new TextCapture(@"History\s*(.*?)\s*(?:Technician\s*(.*))?$", nameof(ReportData.History), nameof(ReportData.Technician)),
            new TextCapture("Physician *(.*) Ward *(.*)", nameof(ReportData.Physician), nameof(ReportData.Ward)),
            new TextCapture("Insurance *(.*)", nameof(ReportData.Insurance)),
        };

        static TextCapture[] _ambientCaptures =
        {
            new TextCapture("Pre: (.+(?:AM|PM)).*Ambient: (.+%)"),
            new TextCapture("Post: (.+(?:AM|PM)).*Ambient: (.+%)"),
        };

        static Regex _ambientMeasured = new Regex("Measured: (.+(?:AM|PM)).*Ambient: (.+%)");

        static TextCapture[] _spirometrySixCaptures =
        {
            new TextCapture(@"FEV 1 \[L\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)"),
            new TextCapture(@"FVC \[L\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)"),
            new TextCapture(@"FEV1%I \[%\](.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)"),
            new TextCapture(@"MMEF \[L/s\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)"),
            new TextCapture(@"MEF 50 \[L/s\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)"),
            new TextCapture(@"PEF \[L/s\] (.+) (.+) (.+%|-) (.+) (.+%|-) (.+%|-)"),
        };

        static TextCapture[] _spirometryThreeCaptures =
{
            new TextCapture(@"FEV 1 \[L\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"FVC \[L\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"FEV1%I \[%\](.+) (.+) (.+%|-)"),
            new TextCapture(@"MMEF \[L/s\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"MEF 50 \[L/s\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"PEF \[L/s\] (.+) (.+) (.+%|-)"),
        };

        static TextCapture[] _diffusionCaptures =
        {
            new TextCapture(@"DLCO \[ml/min/mmHg\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"DLCOc \[ml/min/mmHg\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"KCO \[ml/min/mmHg/L\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"KCOc \[ml/min/mmHg/L\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"VA \[L\] (.+) (.+) (.+%|-)"),
            new TextCapture(@"Hb \[mmol/L\] (.+) (.+) (.+%|-)"),
        };

        public static ReportData Parse(Stream fileStream)
        {
            var reader = new PdfReader(fileStream);
            var text = PdfTextExtractor.GetTextFromPage(reader, 1);
            var lines = text.Split('\n').ToList();

            var data = new ReportData();

            ParseTitle(data, lines);
            ParseAmbient(data, lines);
            ParseSpirometry(data, lines);
            ParseDiffusion(data, lines);
            ParseNotes(data, lines);

            CapitalizeGender(data);

            return data;
        }

        private static void CapitalizeGender(ReportData data)
        {
            if (data.Gender != null && data.Gender.Length > 0)
            {
                if (data.Gender.Length == 1)
                {
                    data.Gender = data.Gender.ToUpper();
                }
                else
                {
                    data.Gender = char.ToUpper(data.Gender[0]) + data.Gender.Substring(1, data.Gender.Length - 1);
                }
            }
        }

        private static void ParseNotes(ReportData data, List<string> lines)
        {
            var notes = lines.IndexOf("Technician notes");
            var interpretation = lines.IndexOf("Interpretation", notes > -1 ? notes : 0);

            if (notes > -1 && interpretation > notes + 1 && notes < lines.Count - 1)
            {
                var end = interpretation > -1 ? interpretation : lines.Count;
                data.TechnicianNotes = lines.GetRange(notes + 1, end - notes - 1).Aggregate((o, n) => o + n);
            }

            if (interpretation > -1 && interpretation < lines.Count - 3)
            {
                data.Interpretation = lines.GetRange(interpretation + 1, lines.Count - interpretation - 3).Aggregate((o, n) => o + n);
            }
        }

        private static void ParseAmbient(ReportData data, List<string> lines)
        {
            var match = lines.Select(l => _ambientMeasured.Match(l)).FirstOrDefault(m => m.Success);
            if (match == null)
            {
                var num = 0;
                foreach (var capture in _ambientCaptures)
                {
                    match = lines.Select(l => capture.Match(l)).FirstOrDefault(m => m.Success);
                    if (match != null)
                    {
                        data.AmbientType = AmbientType.PrePost;
                        if (data.AmbientData == null)
                        {
                            data.AmbientData = new AmbientData[2];
                        }
                        data.AmbientData[num] = new AmbientData
                        {
                            DateTime = match.Groups[1].Value,
                            Ambient = match.Groups[2].Value.Replace("°", "\\'b0"),
                        };
                        lines.Remove(match.Value);
                        num++;
                    }
                }
            }
            else
            {
                data.AmbientType = AmbientType.Measured;
                data.AmbientData = new[]
                {
                    new AmbientData
                    {
                        DateTime = match.Groups[1].Value,
                        Ambient = match.Groups[2].Value.Replace("°", "\\'b0"),
                    }
                };
                lines.Remove(match.Value);
            }
        }

        private static void ParseTitle(ReportData data, List<string> lines)
        {
            foreach (var capture in _simpleCaptures)
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
        }

        private static void ParseDiffusion(ReportData data, List<string> lines)
        {
            data.DiffusionData = new string[6][];
            for (int i = 0; i < 6; i++)
            {
                var match = lines.Select(l => _diffusionCaptures[i].Match(l)).FirstOrDefault(m => m.Success);
                if (match == null)
                {
                    data.DiffusionType = DiffusionType.None;
                    data.DiffusionData = null;
                    return;
                }
                else
                {
                    data.DiffusionData[i] = new string[3];
                    for (int j = 0; j < 3; j++)
                    {
                        data.DiffusionData[i][j] = match.Groups[j + 1].Value;
                    }

                    lines.Remove(match.Value);
                }
            }

            if (lines.Any(l => l == "Ref Post % Ref"))
            {
                data.DiffusionType = DiffusionType.RefPost;
            }
            else if (lines.Any(l => l == "Ref Pre % Ref"))
            {
                data.DiffusionType = DiffusionType.RefPre;
            }
        }

        private static void ParseSpirometry(ReportData data, List<string> lines)
        {
            data.SpirometryData = new string[6][];
            for (int i = 0; i < 6; i++)
            {
                var match = lines.Select(l => _spirometrySixCaptures[i].Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    data.SpirometryType = SpirometryType.SixColumn;

                    data.SpirometryData[i] = new string[6];
                    for (int j = 0; j < 6; j++)
                    {
                        data.SpirometryData[i][j] = match.Groups[j + 1].Value;
                    }

                    lines.Remove(match.Value);
                }
            }
            if (data.SpirometryType == SpirometryType.SixColumn)
            {
                return;
            }

            for (int i = 0; i < 6; i++)
            {
                var match = lines.Select(l => _spirometryThreeCaptures[i].Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    data.SpirometryType = SpirometryType.ThreeColumn;

                    data.SpirometryData[i] = new string[3];
                    for (int j = 0; j < 3; j++)
                    {
                        data.SpirometryData[i][j] = match.Groups[j + 1].Value;
                    }

                    lines.Remove(match.Value);
                }
            }
        }
    }
}
