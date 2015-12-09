using System;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace Common.Excel.GoogleSheets
{
    public static class AtomEntryExtensions
    {
        public static string GetLink(this AtomEntry atomEntry, string linkType)
        {
            return atomEntry.Links.FindService(linkType, null).HRef.ToString();
        }

        public static string GetWorkSheetLink(this SpreadsheetEntry spreadsheetEntry)
        {
            return GetLink(spreadsheetEntry, GDataSpreadsheetsNameTable.WorksheetRel);
        }

        public static WorksheetQuery GetWorkSheetQuery(this SpreadsheetEntry spreadsheetEntry)
        {
            return new WorksheetQuery(GetWorkSheetLink(spreadsheetEntry));
        }

        public static string GetCellsLink(this WorksheetEntry spreadsheetEntry)
        {
            return GetLink(spreadsheetEntry, GDataSpreadsheetsNameTable.CellRel);
        }

        public static CellQuery GetCellsQuery(this WorksheetEntry spreadsheetEntry)
        {
            return new CellQuery(GetCellsLink(spreadsheetEntry));
        }
    }
}
