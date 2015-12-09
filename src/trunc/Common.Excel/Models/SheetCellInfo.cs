using System;
using System.Collections.Generic;
using System.Linq;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace Common.Excel.Models
{
    public class SheetCellInfo
    {
        private static readonly Dictionary<string, string> EscapeStrings = new List<string> {"true", "false", "True", "False"}.ToDictionary(k => k, k => String.Format("=\"{0}\"", k));
        private static readonly Dictionary<string, string> EscapedStrings = EscapeStrings.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

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
                BatchData = new GDataBatchEntryData(IdString, GDataBatchOperationType.query)
            };
        }

        public CellEntry BatchUpdateEntry(CellEntry cellEntry)
        {
            SetValue(cellEntry, Value);
            cellEntry.BatchData = new GDataBatchEntryData(IdString, GDataBatchOperationType.update);

            return cellEntry;
        }

        public static void SetValue(CellEntry cellEntry, string value)
        {
            cellEntry.InputValue = EscapeStrings.ContainsKey(value) ? EscapeStrings[value] : value;
        }

        public static string GetValue(CellEntry cellEntry)
        {
            if (EscapedStrings.ContainsKey(cellEntry.InputValue))
            {
                return EscapedStrings[cellEntry.InputValue];
            }
            else
            {
                return cellEntry.InputValue;
            }
        }
    }

}
