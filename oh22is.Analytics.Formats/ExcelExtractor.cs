using System;
using Microsoft.Analytics.Interfaces;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace oh22is.Analytics.Formats
{
    public class ExcelExtractor : IExtractor
    {

        private string _sheet;

        public ExcelExtractor(string sheet = null)
        {
            _sheet = sheet;
        }

        private static Stream mStream = new MemoryStream();

        public override IEnumerable<IRow> Extract(IUnstructuredReader input, IUpdatableRow output)
        {
            var stream = input.BaseStream;
            if (input.Length > 0)
            {

                var inputBuffer = new byte[input.Length];
                stream.Read(inputBuffer, 0, (int)input.Length);
                mStream.Write(inputBuffer, (int)mStream.Length, inputBuffer.Length);

                var document = SpreadsheetDocument.Open(mStream, false);
                var workbookPart = document.WorkbookPart;

                //var worksheetPart = workbookPart.WorksheetParts.First();

                WorksheetPart worksheetPart;

                if (string.IsNullOrEmpty(_sheet))
                {
                    worksheetPart = workbookPart.WorksheetParts.First();
                }
                else
                {
                    var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == _sheet);
                    worksheetPart = (WorksheetPart) workbookPart.GetPartById(sheet.Id);
                }

                var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                var rowCount = 1;
                string value = null;

                foreach (var row in sheetData.Elements<Row>())
                {
                    foreach (var queryColumn in output.Schema)
                    {
                        var addressName = $"{queryColumn.Name}{rowCount}";
                        var theCell = row.Descendants<Cell>().FirstOrDefault(c => c.CellReference == addressName);

                        if (theCell != null)
                        {
                            value = theCell.InnerText;

                            if (theCell.DataType != null)
                            {
                                switch (theCell.DataType.Value)
                                {
                                    case CellValues.SharedString:

                                        var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                        if (stringTable != null)
                                        {
                                            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                        }
                                        break;

                                    case CellValues.Boolean:
                                        switch (value)
                                        {
                                            case "0":
                                                value = "FALSE";
                                                break;
                                            default:
                                                value = "TRUE";
                                                break;
                                        }
                                        break;
                                }
                            }
                        }

                        output.Set<object>(queryColumn.Name, Convert.ChangeType(value, queryColumn.Type));
                        
                    }

                    rowCount++;
                    yield return output.AsReadOnly();
                }
            }
        }
    }
}