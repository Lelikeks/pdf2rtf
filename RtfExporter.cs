﻿using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace pdf2rtf
{
    public class RtfExporter
    {
        public static void Export(ReportData data, string filePath)
        {
            var templateName = GetTemplateName(data);
            var template = new StringBuilder(File.ReadAllText(Path.Combine("Templates", templateName)));

            var titleProps = typeof(ReportData).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var titleProp in titleProps)
            {
                template.Replace($"#{titleProp.Name}#", titleProp.GetValue(data)?.ToString());
            }
            for (var i = 0; i < data.AmbientData.Length; i++)
            {
                template.Replace($"#Ambient{i}#", data.AmbientData[i].Ambient);
                template.Replace($"#AmbientDate{i}#", data.AmbientData[i].DateTime);
            }
            for (var i = 0; i < 6; i++)
            {
                for (var j = 0; j < (data.SpirometryType == SpirometryType.EightColumn ? 8 : 5); j++)
                {
                    template.Replace($"#s{i}{j}#", data.SpirometryData[i][j]);
                }
            }

            if (data.DiffusionType != DiffusionType.None)
            {
                for (var i = 0; i < 3; i++)
                {
                    for (var j = 0; j < 5; j++)
                    {
                        template.Replace($"#d{i}{j}#", data.DiffusionData[i][j]);
                    }
                }
            }

            if (data.DiffusionType == DiffusionType.None)
            {
                var str = template.ToString();
                var startDiffusion = str.IndexOf(@"\ul\b DIFFUSION\b0\par");
                var finishDiffusion = str.IndexOf(@"\b Technician notes:");
                template.Remove(startDiffusion, finishDiffusion - startDiffusion);
            }

            File.WriteAllText(filePath, template.ToString());
        }

        private static string GetTemplateName(ReportData data)
        {
            switch (data.AmbientType)
            {
                case AmbientType.PrePost when data.SpirometryType == SpirometryType.EightColumn:
                    return "LFX Report_template_amb_8col.rtf";
                case AmbientType.PrePost when data.SpirometryType == SpirometryType.FiveColumn:
                    return "LFX Report_template_amb_5col.rtf";
                case AmbientType.Measured when data.SpirometryType == SpirometryType.EightColumn:
                    return "LFX Report_template_measured_8col.rtf";
                case AmbientType.Measured when data.SpirometryType == SpirometryType.FiveColumn:
                    return "LFX Report_template_measured_5col.rtf";
                default:
                    throw new Exception("Cannot find template for this file");
            }
        }
    }
}
