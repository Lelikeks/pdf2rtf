using Newtonsoft.Json;
using System;

namespace pdf2rtf
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = PdfParser.Parse(@"D:\000\LFX-Report - 2017-14-8_16-26-12 - Heartscope3Report.pdf"); // 6 col, normal, ref post
            //var data = PdfParser.Parse(@"D:\000\LFX-Report - 2017-11-8_15-53-03 - Heartscope3Report.pdf"); // 3 col, ref pre
            //var data = PdfParser.Parse(@"D:\000\LFX-Report - 2017-11-8_15-51-52 - Heartscope3Report.pdf"); // 6 col, none
            Console.WriteLine(JsonConvert.SerializeObject(data, new JsonSerializerSettings { Formatting = Formatting.Indented, NullValueHandling = NullValueHandling.Ignore }));
            Console.Read();
        }
    }
}
