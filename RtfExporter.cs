using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace pdf2rtf
{
    class RtfExporter
    {
        public static void Export(ReportData data, string filePath)
        {
            var templateName = GetTemplateName(data);
            var template = new StringBuilder(File.ReadAllText(templateName));

            var titleProps = typeof(ReportData).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < titleProps.Length; i++)
            {
                template.Replace($"#{titleProps[i].Name}#", titleProps[i].GetValue(data)?.ToString());
            }
            for (int i = 0; i < data.AmbientData.Length; i++)
            {
                template.Replace($"#Ambient{i}#", data.AmbientData[i].Ambient);
                template.Replace($"#AmbientDate{i}#", data.AmbientData[i].DateTime);
            }
            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < (data.SpirometryType == SpirometryType.SixColumn ? 6 : 3); j++)
                {
                    template.Replace($"#s{i}{j}#", data.SpirometryData[i][j]);
                }
                if (data.DiffusionType != DiffusionType.None)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        template.Replace($"#d{i}{j}#", data.DiffusionData[i][j]);
                    }
                }
            }

            var dtype = "";
            if (data.DiffusionType == DiffusionType.RefPre)
            {
                dtype = "Pre";
            }
            else if (data.DiffusionType == DiffusionType.RefPost)
            {
                dtype = "Post";
            }
            template.Replace($"#DType#", dtype);

            File.WriteAllText(filePath, template.ToString());
            Process.Start(filePath);
        }

        private static string GetTemplateName(ReportData data)
        {
            if (data.AmbientType == AmbientType.PrePost)
            {
                if (data.SpirometryType == SpirometryType.SixColumn)
                {
                    return "LFX Report_template_amb_6col.rtf";
                }
                else if (data.SpirometryType == SpirometryType.ThreeColumn)
                {
                    return "LFX Report_template_amb_3col.rtf";
                }
            }
            else if (data.AmbientType == AmbientType.Measured)
            {
                if (data.SpirometryType == SpirometryType.SixColumn)
                {
                    return "LFX Report_template_measured_6col.rtf";
                }
                else if (data.SpirometryType == SpirometryType.ThreeColumn)
                {
                    return "LFX Report_template_measured_3col.rtf";
                }
            }
            throw new Exception("Cannot find template for this file");
        }
    }
}
