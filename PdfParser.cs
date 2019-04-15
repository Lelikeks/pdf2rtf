using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace pdf2rtf
{
    public class PdfParser
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

        static TextCapture[] _spirometryEightCaptures =
        {
            new TextCapture(@"FEV 1 \[L\] (.+?|-) (.+?|-) (.+?|-) (.+%|-) *(.*?) (.+?|-) (.+%|-) (.+%|-)"),
            new TextCapture(@"FVC \[L\] (.+?|-) (.+?|-) (.+?|-) (.+%|-) *(.*?) (.+?|-) (.+%|-) (.+%|-)"),
            new TextCapture(@"(?:FEV1%I|FEV1%F|FEV1%VCmax|FEV1%FVC) *\[%\] (.+?|-) (.+?|-) (.+?|-) (.+%|-) *(.*?) (.+?|-) (.+%|-) (.+%|-)"),
            new TextCapture(@"MMEF \[L/s\] (.+?|-) (.+?|-) (.+?|-) (.+%|-) *(.*?) (.+?|-) (.+%|-) (.+%|-)"),
            new TextCapture(@"MEF 50 \[L/s\] (.+?|-) (.+?|-) (.+?|-) (.+%|-) *(.*?) (.+?|-) (.+%|-) (.+%|-)"),
            new TextCapture(@"PEF \[L/s\] (.+?|-) (.+?|-) (.+?|-) (.+%|-) *(.*?) (.+?|-) (.+%|-) (.+%|-)"),
        };

        static TextCapture[] _spirometryFiveCaptures =
{
            new TextCapture(@"FEV 1 \[L\] (.+) (.+) (.+) (.+%|-) (.+)"),
            new TextCapture(@"FVC \[L\] (.+) (.+) (.+) (.+%|-) (.+)"),
            new TextCapture(@"(?:FEV1%I|FEV1%F|FEV1%VCmax|FEV1%FVC) *\[%\](.+) (.+) (.+) (.+%|-) (.+)"),
            new TextCapture(@"MMEF \[L/s\] (.+) (.+) (.+) (.+%|-) (.+)"),
            new TextCapture(@"MEF 50 \[L/s\] (.+) (.+) (.+) (.+%|-) (.+)"),
            new TextCapture(@"PEF \[L/s\] (.+) (.+) (.+) (.+%|-) (.+)"),
        };

        static TextCapture[] _diffusionCaptures =
        {
            new TextCapture(@"DLCO \[ml/min/mmHg\] (.+) (.+) (.+) (.+%|-)(?:$| ([^ ]+)| +)"),
            new TextCapture(@"KCO \[ml/min/mmHg/L\] (.+) (.+) (.+) (.+%|-)(?:$| ([^ ]+)| +)"),
            new TextCapture(@"VA \[L\] (.+) (.+) (.+) (.+%|-)(?:$| ([^ ]+)| +)"),
        };

        public static ReportData Parse(string file)
        {
            using (var stream = File.OpenRead(file))
            {
                return Parse(stream);
            }
        }

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

            if (interpretation > -1 && interpretation < lines.Count - 1)
            {
                var ganshorn = lines.LastOrDefault(x => x.Contains("GANSHORN"));
                if (ganshorn == null)
                {
                    if (lines.Count - interpretation > 1)
                        data.Interpretation = lines.GetRange(interpretation + 1, lines.Count - interpretation - 1).Aggregate((o, n) => o + n);
                }
                else
                {
                    var footer = lines.IndexOf(ganshorn);
                    if (footer - interpretation  > 1)
                        data.Interpretation = lines.GetRange(interpretation + 1, footer - interpretation - 1).Aggregate((o, n) => o + n);
                }
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
                    for (var i = 0; i < capture.Properties.Length; i++)
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
            data.DiffusionData = new string[_diffusionCaptures.Length][];
            for (var i = 0; i < _diffusionCaptures.Length; i++)
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
                    data.DiffusionData[i] = new string[5];
                    for (var j = 0; j < 5; j++)
                    {
                        data.DiffusionData[i][j] = match.Groups[j + 1].Value;
                    }

                    lines.Remove(match.Value);
                }
            }

            data.DiffusionType = DiffusionType.Present;
        }

        private static void ParseSpirometry(ReportData data, List<string> lines)
        {
            data.SpirometryData = new string[_spirometryEightCaptures.Length][];
            for (var i = 0; i < _spirometryEightCaptures.Length; i++)
            {
                var match = lines.Select(l => _spirometryEightCaptures[i].Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    data.SpirometryType = SpirometryType.EightColumn;

                    data.SpirometryData[i] = new string[8];
                    for (var j = 0; j < 8; j++)
                    {
                        data.SpirometryData[i][j] = match.Groups[j + 1].Value;
                    }

                    lines.Remove(match.Value);
                }
            }
            if (data.SpirometryType == SpirometryType.EightColumn)
            {
                return;
            }

            for (var i = 0; i < _spirometryEightCaptures.Length; i++)
            {
                var match = lines.Select(l => _spirometryFiveCaptures[i].Match(l)).FirstOrDefault(m => m.Success);
                if (match != null)
                {
                    data.SpirometryType = SpirometryType.FiveColumn;

                    data.SpirometryData[i] = new string[5];
                    for (var j = 0; j < 5; j++)
                    {
                        data.SpirometryData[i][j] = match.Groups[j + 1].Value;
                    }

                    lines.Remove(match.Value);
                }
            }
        }
    }
}
