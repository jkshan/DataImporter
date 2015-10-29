using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DataImporter.Business.Interface;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataImporter.Business.Parser
{
    public class ExcelParser : IParser
    {
        private Stream _fileStream;
        public string ExcelSheetName { get; set; }

        public ExcelParser(Stream FileStream, string SheetName)
        {
            _fileStream = FileStream;
            ExcelSheetName = SheetName;
        }

        #region Excel Helper Functions
        ///<summary>returns an empty cell when a blank cell is encountered
        ///</summary>
        private static IEnumerable<Cell> GetRowCells(Row row)
        {
            int currentCount = 0;

            foreach(Cell cell in row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                int currentColumnIndex = GetColumnIndexFromName(cell.CellReference);

                for(; currentCount < currentColumnIndex; currentCount++)
                {
                    yield return new DocumentFormat.OpenXml.Spreadsheet.Cell();
                }

                yield return cell;
                currentCount++;
            }
        }

        /// <summary>
        /// Return Zero-based column index given a cell name or column name
        /// </summary>
        /// <param name="columnNameOrCellReference">Column Name (ie. A, AB3, or AB44)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
        public static int GetColumnIndexFromName(string columnNameOrCellReference)
        {
            int columnIndex = 0;
            int factor = 1;
            for(int pos = columnNameOrCellReference.Length - 1; pos >= 0; pos--)   // R to L
            {
                if(Char.IsLetter(columnNameOrCellReference[pos]))  // for letters (columnName)
                {
                    columnIndex += factor * (columnNameOrCellReference[pos] - 'A');
                    factor *= 26;
                }
            }
            return columnIndex;
        }

        /// <summary>
        /// Get the Cell values, Handles only String and Numbers
        /// </summary>
        /// <param name="SharedStringList">List of SharedString Elements of the Workbook to which the Cell belongs to.</param>
        /// <param name="c">Cell for which the Value to be retrived.</param>
        /// <returns>Returns the cell values</returns>
        private static string GetCellValue(Cell c, List<OpenXmlElement> SharedStringList)
        {
            string text;
            //Numeric will be stored in Cell itself.
            text = c.InnerText;
            if(c.DataType != null)
            {
                switch(c.DataType.Value)
                {
                    //If the cell Type is Shared String, then cell will contain the Index of the String value in the Shared String Collection.
                    case CellValues.SharedString:
                        int index = 0;
                        if(int.TryParse(text, out index) && index < SharedStringList.Count)
                        {
                            text = SharedStringList[index].InnerText;
                        }
                        break;
                }
            }
            return text;
        }
        #endregion

        public IEnumerable<string[]> Readline()
        {
            using(SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_fileStream, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                var theSheet = workbookPart.Workbook.Sheets.Cast<Sheet>().Where(i => i.Name == ExcelSheetName).FirstOrDefault();
                if(theSheet != null)
                {
                    var wsPart = workbookPart.GetPartById(theSheet.Id);

                    //Load the Shared Stirng Element into List.(For Performance)
                    var stringTableList = workbookPart.SharedStringTablePart.SharedStringTable.ChildElements.ToList();

                    OpenXmlReader reader = OpenXmlReader.Create(wsPart);

                    while(reader.Read())
                    {
                        if(reader.ElementType == typeof(Row))
                        {
                            for(var r = (Row)reader.LoadCurrentElement(); r != null; r = r.NextSibling<Row>())
                            {
                                var values = new List<string>();

                                var cells = GetRowCells(r);

                                foreach(var c in cells)
                                {
                                    values.Add(GetCellValue(c, stringTableList));
                                }
                                yield return values.ToArray();
                            }
                        }
                    }
                }

            }
        }
    }
}