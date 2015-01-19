using System;
using System.Drawing;
using System.Windows.Forms;

namespace Common.Excel
{
    public static partial class Excel
    {
        public static double GetDefaultFontWidth(string text)
        {
            const string font = "Calibri";
            const int fontSize = 11;
            System.Drawing.Font stringFont = new System.Drawing.Font(font, fontSize);
            return GetWidth(stringFont, text) + 2.0;
        }

        private static double GetWidth(string text, string font, int fontSize)
        {
            Font stringFont = new Font(font, fontSize);
            return GetWidth(stringFont, text);
        }

        private static double GetWidth(Font stringFont, string text)
        {
            // This formula is based on this article plus a nudge ( + 0.2M )
            // http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.width.aspx
            // Truncate(((256 * Solve_For_This + Truncate(128 / 7)) / 256) * 7) = DeterminePixelsOfString

            Size textSize = TextRenderer.MeasureText(text, stringFont, new Size(int.MaxValue, int.MaxValue) , TextFormatFlags.SingleLine|TextFormatFlags.LeftAndRightPadding);
            double width = (double)(((textSize.Width / (double)7) * 256) - (128 / 7)) / 256;
            width = (double)decimal.Round((decimal)width + 0.2M, 2);

            return width;
        }
    }
}
