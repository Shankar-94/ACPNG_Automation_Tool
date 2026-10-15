using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace SimpleSapScraper.Core.Excel
{
    public sealed class ExcelMsnReader
    {
        private static readonly Regex FiveDigitMsnRegex = new(@"\b\d{5}\b", RegexOptions.Compiled);

        public IReadOnlyList<string> ReadMsnsFromExcel(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Excel file path is empty.", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Excel file was not found.", filePath);
            }

            var extractedMsns = new List<string>();

            using var workbook = new XLWorkbook(filePath);

            foreach (var worksheet in workbook.Worksheets)
            {
                var usedRange = worksheet.RangeUsed();

                if (usedRange == null)
                {
                    continue;
                }

                var msnColumnNumber = FindMsnColumnNumber(usedRange);

                if (msnColumnNumber == null)
                {
                    continue;
                }

                var firstDataRow = usedRange.FirstRowUsed().RowNumber() + 1;
                var lastDataRow = usedRange.LastRowUsed().RowNumber();

                for (var rowNumber = firstDataRow; rowNumber <= lastDataRow; rowNumber++)
                {
                    var cell = worksheet.Cell(rowNumber, msnColumnNumber.Value);

                    var rawValue = cell.GetFormattedString()?.Trim();

                    if (string.IsNullOrWhiteSpace(rawValue))
                    {
                        rawValue = cell.GetString()?.Trim();
                    }

                    if (string.IsNullOrWhiteSpace(rawValue))
                    {
                        continue;
                    }

                    foreach (Match match in FiveDigitMsnRegex.Matches(rawValue))
                    {
                        extractedMsns.Add(match.Value);
                    }
                }

                if (extractedMsns.Count > 0)
                {
                    break;
                }
            }

            return extractedMsns
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static int? FindMsnColumnNumber(IXLRange usedRange)
        {
            var headerRow = usedRange.FirstRowUsed();

            if (headerRow == null)
            {
                return null;
            }

            foreach (var cell in headerRow.CellsUsed())
            {
                var header = cell.GetFormattedString()?.Trim();

                if (string.Equals(header, "MSN", StringComparison.OrdinalIgnoreCase))
                {
                    return cell.Address.ColumnNumber;
                }
            }

            return null;
        }
    }
}