using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using pdf2rtf.Properties;

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
                return Properties.Settings.Default;
            }
        }

        static void Main(string[] args)
        {
            Trace.Listeners.Add(new ConsoleTraceListener());
            Trace.Listeners.Add(new TextWriterTraceListener("pdf2rtf.log"));
            Trace.AutoFlush = true;

            Console.WriteLine($"start monitoring directory {Settings.IncomingPath}, press enter to exit ...");

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
                Path = Settings.IncomingPath,
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

        private static void ReadDirectory()
        {
            foreach (var file in Directory.GetFiles(Settings.IncomingPath, "*.pdf"))
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
                    RtfExporter.Export(data, $@"{Settings.OutgoingPath}\{fileName}");
                    await MoveFile(filePath, Path.Combine(Settings.ProcessedPath, Path.GetFileName(filePath)));

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
