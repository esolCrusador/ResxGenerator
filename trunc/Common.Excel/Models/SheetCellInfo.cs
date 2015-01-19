using System;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace Common.Excel.Models
{
    public class SheetCellInfo
    {
        public SheetCellInfo(int rowIndex, int colIndex, string value)
        {
            Column = (uint) colIndex + 1;
            Row = (uint) rowIndex + 1;

            Value = value;

            IdString = String.Format("R{0}C{1}", Row.ToString(), Column.ToString());
        }

        public string IdString { get; private set; }

        public string GetIdUri(string baseUri)
        {
            return String.Format("{0}/{1}", baseUri, IdString);
        }

        public uint Column { get; private set; }
        public uint Row { get; private set; }
        public string Value { get; set; }

        public CellEntry BatchCreateEntry(string baseUri)
        {
            return new CellEntry(Row, Column, Value)
            {
                Id = new AtomId(GetIdUri(baseUri)),
                BatchData = new GDataBatchEntryData(IdString, GDataBatchOperationType.query),
                InputValue = Value
            };
        }

        public CellEntry BatchUpdateEntry(CellEntry cellEntry)
        {
            cellEntry.InputValue = Value;
            cellEntry.BatchData = new GDataBatchEntryData(IdString, GDataBatchOperationType.update);

            return cellEntry;
        }
    }
}
