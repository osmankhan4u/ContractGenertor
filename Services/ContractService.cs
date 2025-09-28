using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Text.Json;
using ClosedXML.Excel;
using ContractGeneratorBlazor.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ContractGeneratorBlazor.Services
{
    public interface IContractService
    {
    List<string> GetContractTypes();
    List<ContractTemplate> LoadContracts();
    Task<byte[]> GenerateContractsAsync(Stream excelStream, string contractType, string outputFormat);
    Task AddContractAsync(string type, Dictionary<string, string> placeholders, Stream templateStream, string templateFileName);
    void SaveContracts(List<ContractTemplate> contracts);
    Task<byte[]> GenerateContractsFromFilesAsync(Stream templateStream, Stream excelStream);
    }

    public class ContractService : IContractService
    {
        private readonly IDocumentGenerator _docGenerator;
        private readonly ILogger<ContractService> _logger;
        private readonly string _contractsFolder;
        private readonly string _contractsJsonPath;

        public ContractService(
            IDocumentGenerator docGenerator,
            ILogger<ContractService> logger)
        {
            _docGenerator = docGenerator;
            _logger = logger;
            var rootDir = Directory.GetCurrentDirectory();
            _contractsFolder = Path.Combine(rootDir, "Contracts");
            _contractsJsonPath = Path.Combine(_contractsFolder, "contracts.json");
        }

        // Extract placeholders from Word template using regex
        private List<string> ExtractPlaceholdersFromTemplate(string templatePath)
        {
            var placeholders = new List<string>();
            using (var wordDoc = WordprocessingDocument.Open(templatePath, false))
            {
                var body = wordDoc.MainDocumentPart?.Document?.Body;
                if (body != null)
                {
                    foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                    {
                        var matches = System.Text.RegularExpressions.Regex.Matches(text.Text, "{[^}]+}");
                        foreach (System.Text.RegularExpressions.Match match in matches)
                        {
                            placeholders.Add(match.Value);
                        }
                    }
                }
            }
            return placeholders.Distinct().ToList();
        }

        public async Task<byte[]> GenerateContractsFromFilesAsync(Stream templateStream, Stream excelStream)
        {
            // Save uploaded template to temp file
            var tempTemplatePath = Path.Combine(Path.GetTempPath(), $"template_{Guid.NewGuid()}.docx");
            using (var fileStream = File.Create(tempTemplatePath))
            {
                await templateStream.CopyToAsync(fileStream);
            }

            // Extract placeholders from template
            var placeholders = ExtractPlaceholdersFromTemplate(tempTemplatePath);

            // Read Excel
            using var memoryStream = new MemoryStream();
            await excelStream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;
            using var workbook = new XLWorkbook(memoryStream);
            var worksheet = workbook.Worksheet(1);
            var headerRow = worksheet.FirstRowUsed();
            if (headerRow == null)
                throw new InvalidOperationException("No header row found in the Excel worksheet.");
            var headerMap = headerRow.Cells()
                .Select((cell, index) => new { Name = cell.Value.ToString().Trim(), Index = index + 1 })
                .ToDictionary(x => x.Name!, x => x.Index);
            var rangeUsed = worksheet.RangeUsed();
            if (rangeUsed == null)
                throw new InvalidOperationException("No data range found in the Excel worksheet.");
            var rows = rangeUsed.RowsUsed().Skip(1);

            var outputFolder = Path.Combine(Path.GetTempPath(), $"Contracts_{Guid.NewGuid()}");
            Directory.CreateDirectory(outputFolder);

            foreach (var row in rows)
            {
                var tempDocPath = Path.Combine(outputFolder, $"temp_{Guid.NewGuid()}.docx");
                File.Copy(tempTemplatePath, tempDocPath);
                using (var wordDoc = WordprocessingDocument.Open(tempDocPath, true))
                {
                    var body = wordDoc.MainDocumentPart?.Document?.Body;
                    if (body != null)
                    {
                        foreach (var placeholder in placeholders)
                        {
                            var colName = placeholder.Replace("{", "").Replace("}", "");
                            if (!headerMap.TryGetValue(colName, out int colIndex))
                                continue;
                            var cellValue = row.Cell(colIndex).GetValue<string>() ?? string.Empty;
                            foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                            {
                                if (text.Text.Contains(placeholder))
                                {
                                    text.Text = text.Text.Replace(placeholder, cellValue);
                                }
                            }
                        }
                        wordDoc.MainDocumentPart?.Document?.Save();
                    }
                }
                var fileName = $"Contract_{Guid.NewGuid()}";
                var filePath = Path.Combine(outputFolder, $"{fileName}.docx");
                File.Copy(tempDocPath, filePath);
                File.Delete(tempDocPath);
            }
            File.Delete(tempTemplatePath);

            // Package all generated files into a ZIP
            var zipPath = outputFolder + ".zip";
            ZipFile.CreateFromDirectory(outputFolder, zipPath);
            var result = await File.ReadAllBytesAsync(zipPath);
            Directory.Delete(outputFolder, true);
            File.Delete(zipPath);
            return result;
        }

        public List<string> GetContractTypes()
        {
            var contracts = LoadContracts();
            return contracts.Select(c => c.Type).ToList();
        }

        public List<ContractTemplate> LoadContracts()
        {
            if (!File.Exists(_contractsJsonPath)) return new List<ContractTemplate>();
            var json = File.ReadAllText(_contractsJsonPath);
            return JsonSerializer.Deserialize<List<ContractTemplate>>(json) ?? new List<ContractTemplate>();
        }

        public void SaveContracts(List<ContractTemplate> contracts)
        {
            var json = JsonSerializer.Serialize(contracts, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_contractsJsonPath, json);
        }

        public async Task AddContractAsync(string type, Dictionary<string, string> placeholders, Stream templateStream, string templateFileName)
        {
            Directory.CreateDirectory(_contractsFolder);
            var templatePath = Path.Combine(_contractsFolder, templateFileName);
            using (var fileStream = File.Create(templatePath))
            {
                await templateStream.CopyToAsync(fileStream);
            }
            var contracts = LoadContracts();
            contracts.Add(new ContractTemplate
            {
                Type = type,
                Placeholders = placeholders,
                TemplatePath = templatePath
            });
            SaveContracts(contracts);
        }

        public async Task<byte[]> GenerateContractsAsync(Stream excelStream, string contractType, string outputFormat)
        {
            // Load contract templates from configuration
            var contracts = LoadContracts();
            var templateConfig = contracts.FirstOrDefault(c => c.Type == contractType);
            _logger.LogInformation($"Selected contract type: {contractType}");
            if (templateConfig != null)
            {
                var phList = string.Join(", ", templateConfig.Placeholders.Keys);
                _logger.LogInformation($"Placeholders for {contractType}: {phList}");
            }
            if (templateConfig == null)
                throw new ArgumentException($"Contract type '{contractType}' not found in configuration.");

            if (!File.Exists(templateConfig.TemplatePath))
            {
                _logger.LogError($"Template file not found: {templateConfig.TemplatePath}");
                // Return empty zip or error message as bytes
                using var ms = new MemoryStream();
                using (var archive = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, true))
                {
                    var entry = archive.CreateEntry("ERROR.txt");
                    using var writer = new StreamWriter(entry.Open());
                    writer.Write($"Template file not found: {templateConfig.TemplatePath}");
                }
                return ms.ToArray();
            }

            using var memoryStream = new MemoryStream();
            await excelStream.CopyToAsync(memoryStream);
            memoryStream.Position = 0;

            using var workbook = new XLWorkbook(memoryStream);
            var worksheet = workbook.Worksheet(1);

            var headerRow = worksheet.FirstRowUsed();
            if (headerRow == null)
                throw new InvalidOperationException("No header row found in the Excel worksheet.");
            var headerMap = headerRow.Cells()
                .Select((cell, index) => new { Name = cell.Value.ToString().Trim(), Index = index + 1 })
                .ToDictionary(x => x.Name!, x => x.Index);

            var rangeUsed = worksheet.RangeUsed();
            if (rangeUsed == null)
                throw new InvalidOperationException("No data range found in the Excel worksheet.");
            var rows = rangeUsed.RowsUsed().Skip(1);

            var outputFolder = Path.Combine(Path.GetTempPath(), $"Contracts_{Guid.NewGuid()}");
            Directory.CreateDirectory(outputFolder);

            foreach (var row in rows)
            {
                // Replace placeholders in Word document using OpenXml
                var tempDocPath = Path.Combine(outputFolder, $"temp_{Guid.NewGuid()}.docx");
                File.Copy(templateConfig.TemplatePath, tempDocPath);

                using (var wordDoc = WordprocessingDocument.Open(tempDocPath, true))
                {
                    var body = wordDoc.MainDocumentPart?.Document?.Body;
                    if (body != null)
                    {
                        foreach (var placeholder in templateConfig.Placeholders)
                        {
                            // Remove curly braces for matching Excel column names
                            var excelColName = placeholder.Value.Replace("{", "").Replace("}", "");
                            if (!headerMap.TryGetValue(excelColName, out int colIndex))
                            {
                                _logger.LogWarning($"Column '{excelColName}' not found in Excel, skipping placeholder {placeholder.Key}");
                                continue;
                            }

                            var cellValue = row.Cell(colIndex).GetValue<string>() ?? string.Empty;
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                _logger.LogWarning($"Excel value for column '{excelColName}' is empty for this row, placeholder {placeholder.Key} will not be replaced.");
                            }
                            else
                            {
                                _logger.LogInformation($"Replacing {placeholder.Key} with value '{cellValue}'");
                            }

                            foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                            {
                                if (text.Text.Contains(placeholder.Key))
                                {
                                    text.Text = text.Text.Replace(placeholder.Key, cellValue);
                                }
                            }
                        }
                        wordDoc.MainDocumentPart?.Document?.Save();
                    }
                }

                var fileName = $"{contractType}_{Guid.NewGuid()}";
                var filePath = Path.Combine(outputFolder, $"{fileName}.docx");
                File.Copy(tempDocPath, filePath);
                File.Delete(tempDocPath);
            }

            // Package all generated files into a ZIP
            var zipPath = outputFolder + ".zip";
            ZipFile.CreateFromDirectory(outputFolder, zipPath);
            var result = await File.ReadAllBytesAsync(zipPath);

            Directory.Delete(outputFolder, true);
            File.Delete(zipPath);

            return result;
        }
        private byte[] GeneratePdfFromText(string text)
        {
            // PDF generation removed
            return Array.Empty<byte>();
        }
    }
}
