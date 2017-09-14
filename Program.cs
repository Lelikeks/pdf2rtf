using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using pdf2rtf.Properties;
using System.Configuration;
using System.Text;
using System.Collections.Generic;

namespace pdf2rtf
{
    class Program
    {
        static ConcurrentDictionary<string, bool> FilesQueue = new ConcurrentDictionary<string, bool>();
        static ConcurrentDictionary<string, int> TryCounter = new ConcurrentDictionary<string, int>();

        static Settings Settings
        {
            get
            {
                return Settings.Default;
            }
        }

        static void Main(string[] args)
        {
            Trace.Listeners.Add(new ConsoleTraceListener());
            Trace.Listeners.Add(new TextWriterTraceListener("pdf2rtf.log"));
            Trace.AutoFlush = true;

            if (!CheckFolders())
            {
                Console.Read();
                return;
            }

            Console.WriteLine($"start monitoring folder {Settings.InputFolder}, press enter to exit ...");

            for (int i = 0; i < 3; i++)
            {
                Task.Run(async () =>
                {
                    while (true)
                    {
                        var ready = FilesQueue.FirstOrDefault(kv => kv.Value == false);
                        if (ready.Key != null && FilesQueue.TryUpdate(ready.Key, true, false))
                        {
                            if (await ProcessFile(ready.Key))
                            {
                                FilesQueue.TryRemove(ready.Key, out bool value);
                            }
                            else
                            {
                                var attempts = TryCounter.AddOrUpdate(ready.Key, 1, (s, v) => v + 1);
                                if (attempts > 4)
                                {
                                    FilesQueue.TryRemove(ready.Key, out bool value);
                                }
                                else
                                {
                                    FilesQueue.TryUpdate(ready.Key, false, true);
                                }
                            }
                        }
                        else
                        {
                            await Task.Delay(100);
                        }
                    }
                });
            }

            var watcher = new FileSystemWatcher
            {
                Path = Settings.InputFolder,
                Filter = "*.pdf"
            };
            watcher.Created += (sender, e) =>
            {
                ReadDirectory();
            };
            watcher.EnableRaisingEvents = true;

            ReadDirectory();

            Console.Read();
        }

        private static bool CheckFolders()
        {
            var empty = new List<string>();
            if (string.IsNullOrEmpty(Settings.InputFolder))
            {
                empty.Add("InputFolder");
            }
            if (string.IsNullOrEmpty(Settings.OutputFolder))
            {
                empty.Add("OutputFolder");
            }
            if (string.IsNullOrEmpty(Settings.ProcessedFolder))
            {
                empty.Add("ProcessedFolder");
            }
            if (empty.Count > 0)
            {
                Console.WriteLine($"please set the following configuration parameters: {empty.Aggregate((o, n) => o + ", " + n)}");
                return false;
            }

            if (!Directory.Exists(Settings.InputFolder))
            {
                Directory.CreateDirectory(Settings.InputFolder);
            }
            if (!Directory.Exists(Settings.OutputFolder))
            {
                Directory.CreateDirectory(Settings.OutputFolder);
            }
            if (!Directory.Exists(Settings.ProcessedFolder))
            {
                Directory.CreateDirectory(Settings.ProcessedFolder);
            }
            return true;
        }

        private static void ReadDirectory()
        {
            foreach (var file in Directory.GetFiles(Settings.InputFolder, "*.pdf"))
            {
                if (!TryCounter.TryGetValue(file, out int attempts) || attempts < 5)
                {
                    FilesQueue.TryAdd(file, false);
                }
            }
        }

        private async static Task<bool> ProcessFile(string filePath)
        {
            return await Task.Run(async () =>
            {
                try
                {
                    Trace.WriteLine($"start processing {Path.GetFileName(filePath)}");

                    ReportData data;
                    using (var fileStream = await GetFileStream(filePath))
                    {
                        data = PdfParser.Parse(fileStream);
                    }
                    var fileName = GetSafeFilename($"{data.PatientId}_{data.LastName}_{data.FirstName}.rtf");
                    RtfExporter.Export(data, $@"{Settings.OutputFolder}\{fileName}");
                    await MoveFile(filePath, Path.Combine(Settings.ProcessedFolder, Path.GetFileName(filePath)));

                    Trace.WriteLine($"results saved to {fileName}");
                    return true;
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"following error occured processing file {Path.GetFileName(filePath)}: {ex}");
                    return false;
                }
            });
        }

        public static string GetSafeFilename(string filename)
        {
            return string.Join("_", filename.Split(Path.GetInvalidFileNameChars()));
        }

        private static async Task MoveFile(string source, string destination)
        {
            for (int i = 0; i < 100; i++)
            {
                try
                {
                    if (File.Exists(destination))
                    {
                        File.Delete(destination);
                    }
                    File.Move(source, destination);
                    return;
                }
                catch
                {
                    if (i == 100)
                    {
                        throw;
                    }
                    await Task.Delay(100);
                }
            }
        }

        private static async Task<FileStream> GetFileStream(string filePath)
        {
            for (int i = 0; i < 100; i++)
            {
                try
                {
                    return File.OpenRead(filePath);
                }
                catch
                {
                    if (i == 99)
                    {
                        throw;
                    }
                    await Task.Delay(100);
                }
            }
            throw null;
        }
    }
}
