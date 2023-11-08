using System;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using Mono.Options;
using OfficeOpenXml;

namespace GameDataExporter
{
    public struct Schema
    {
        [JsonPropertyName("tables")]
        public List<Table> Tables { get; set; }
    }

    public struct Table
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }
        [JsonPropertyName("columns")]
        public List<string> Columns { get; set; }
    }

    class Program
    {
        static int Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string schemaDir = AppDomain.CurrentDomain.BaseDirectory;
            string excelDir = AppDomain.CurrentDomain.BaseDirectory;
            string jsonDir = AppDomain.CurrentDomain.BaseDirectory;
            bool showHelp = false;

            JsonDocumentOptions jsonDocumentOptions = new JsonDocumentOptions() { MaxDepth = 5 };

            var p = new OptionSet()
            {
                { "s|schemaDir=", "schema.json file directory", v => schemaDir = v },
                { "e|excelDir=", "game data excel files input directory", v => excelDir = v },
                { "j|jsonDir=", "json files output directory", v => jsonDir = v },
                { "h|help", "show this message and exit", v => showHelp = true },
            };

            try
            {
                p.Parse(args);
            }
            catch (OptionException)
            {
                Console.WriteLine("Try '--help' for more information.");
                return -1;
            }

            if (showHelp)
            {
                p.WriteOptionDescriptions(Console.Out);
                return 0;
            }


            try
            {
                string schemaPath = Path.Combine(schemaDir, "schema.json");

                using var reader = new StreamReader(schemaPath);
                string json = reader.ReadToEnd();
                var schema = JsonSerializer.Deserialize<Schema>(json);

                foreach (Table table in schema.Tables)
                {
                    string excelPath = Path.Combine(excelDir, $"{table.Name}.xlsx");
                    using var excelPackage = new ExcelPackage(excelPath);

                    foreach (var sheet in excelPackage.Workbook.Worksheets)
                    {
                        var startCell = sheet.Dimension.Start;
                        string outputFileName = sheet.GetValue<string>(startCell.Row, startCell.Column);
                        if (outputFileName != $"{table.Name}.json")
                        {
                            continue;
                        }

                        string outputPath = Path.Combine(jsonDir, $"GameData{table.Name}.json");

                        using var stream = new FileStream(outputPath, FileMode.Create);
                        using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions() { Indented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping });

                        writer.WriteStartObject();

                        var endCell = sheet.Dimension.End;
                        int keyInfoRow = startCell.Row + 1;

                        var keyInfoList = new List<(string keyName, string keyType, bool isArray)>();
                        for(int c = startCell.Column; c <= endCell.Column; c++)
                        {
                            var keyInfos = sheet.GetValue<string>(keyInfoRow, c).Split(new char[2] { ':', ' ' });
                            keyInfoList.Add((keyName: keyInfos[0],  keyType: keyInfos[1].ToLower(), isArray: keyInfos.Length > 2));
                        }

                        for (int r = startCell.Row + 2; r <= endCell.Row; r++)
                        {
                            bool isStartedObject = false;
                            for (int c = startCell.Column; c <= endCell.Column; c++)
                            {
                                var keyInfo = keyInfoList[c-1];
                                if (!table.Columns.Contains(keyInfo.keyName))
                                {
                                    continue;
                                }

                                if (keyInfo.keyName.Equals("Id"))
                                {
                                    var key = sheet.GetValue<string>(r, c);
                                    if (string.IsNullOrEmpty(key))
                                    {
                                        break;
                                    }

                                    writer.WriteStartObject(key);
                                    isStartedObject = true;
                                }

                                if (keyInfo.isArray)
                                {
                                    var arrayStrData = sheet.GetValue<string>(r, c);
                                    if (string.IsNullOrEmpty(arrayStrData))
                                    {
                                        continue;
                                    }

                                    writer.WritePropertyName(keyInfo.keyName);
                                    JsonNode.Parse(arrayStrData, null, jsonDocumentOptions)?.WriteTo(writer);
                                }
                                else
                                {
                                    switch (keyInfo.keyType)
                                    {
                                        case "string":
                                            writer.WriteString(keyInfo.keyName, sheet.GetValue<string>(r, c) ?? "");
                                            break;
                                        case "bool":
                                            writer.WriteBoolean(keyInfo.keyName, sheet.GetValue<bool>(r, c));
                                            break;
                                        case "int":
                                            writer.WriteNumber(keyInfo.keyName, sheet.GetValue<int>(r, c));
                                            break;
                                        case "float":
                                            writer.WriteNumber(keyInfo.keyName, sheet.GetValue<float>(r, c));
                                            break;
                                    }
                                }
                            }

                            if (isStartedObject)
                            {
                                writer.WriteEndObject();
                            }
                        }

                        writer.WriteEndObject();
                        writer.Flush();
                    }
                }
            }
            catch (Exception e)
            {
                PrintMsg(e.ToString(), ConsoleColor.Red);
                return -1;
            }

            PrintMsg("Export Success!", ConsoleColor.Yellow);
            return 0;
        }


        static void PrintMsg(string msg, ConsoleColor color = ConsoleColor.White)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(msg);
            Console.ResetColor();
        }
    }
}
