using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Reflection;

namespace WordAddInTab
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void BtnImport_Click(object sender, RibbonControlEventArgs e)
        {
            int i = 0;
            int j = 0;
            int numRows = 2;
            int numColumns = 4;
            int numRowsTable = 40;
            int numColumnsTable = 4;
            string strText;
            Word.Paragraph paragraph;
            object oMissing = Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word.Application application = Globals.ThisAddIn.Application;
            Word._Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Table table;

            try
            {
                Word.Range secondRange = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                table = document.Tables.Add(secondRange, numRows, numColumns, ref oMissing, ref oMissing);
                table.Range.ParagraphFormat.SpaceAfter = 0;
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Range.Font.Size = 12f;

                for (i = 1; i <= numRows; i++)
                    for (j = 1; j <= numColumns; j++)
                    {
                        strText = "T1 Cell (" + i + "," + j + ")";
                        table.Cell(i, j).Range.Text = strText;
                    }
                table.Rows[1].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGold;
                table.Rows[1].HeadingFormat = -1;
                table.ApplyStyleHeadingRows = true;

                object range = document.Bookmarks.get_Item(ref oEndOfDoc).Range; //go to end of the page
                paragraph = document.Content.Paragraphs.Add(ref range); //add paragraph at end of document
                paragraph.Range.Text = "Continue with our demo...  ";
                paragraph.Range.Font.Size = 18f;
                paragraph.Format.SpaceAfter = 0;

                table = null;
                Word.Range wrdRng = document.Bookmarks.get_Item(ref oEndOfDoc).Range;
                table = document.Tables.Add(wrdRng, numRowsTable, numColumnsTable, ref oMissing, ref oMissing);
                table.Range.ParagraphFormat.SpaceAfter = 0;
                table.Range.Font.Size = 12f;

                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                for (i = 1; i <= numRowsTable; i++)
                    for (j = 1; j <= numColumnsTable; j++)
                    {
                        strText = "T2 Cell (" + i + "," + j + ")";
                        table.Cell(i, j).Range.Text = strText;
                    }
                table.Rows[1].Range.Font.Bold = 1;
                table.Rows[1].Range.Font.Italic = 1;
                table.Rows[1].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorSkyBlue;
                table.Rows[1].HeadingFormat = -1;
                table.ApplyStyleHeadingRows = true;
                document.ExportAsFixedFormat(OutputFileName: "test.pdf",
                            ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                            OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen);

            }
            finally
            {
                object saveChanges = false;
                object originalFormat = Missing.Value;
                object routeDocument = Missing.Value;
                document.Close(ref saveChanges, ref originalFormat, ref routeDocument);
                application = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
