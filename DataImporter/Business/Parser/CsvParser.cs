using System.Collections.Generic;
using System.IO;
using DataImporter.Business.Interface;
using Microsoft.VisualBasic.FileIO;

namespace DataImporter.Business.Parser
{
    public class CsvParser : IParser
    {
        private Stream _fileStream;

        public CsvParser(Stream FileStream)
        {
            _fileStream = FileStream;
        }

        public IEnumerable<string[]> Readline()
        {
            using(var textFieldParser = new TextFieldParser(_fileStream))
            {
                textFieldParser.TextFieldType = FieldType.Delimited;
                textFieldParser.Delimiters = new[] { "," };
                textFieldParser.HasFieldsEnclosedInQuotes = true;

                while(!textFieldParser.EndOfData)
                {
                    yield return textFieldParser.ReadFields();
                }
            }
        }
    }
}