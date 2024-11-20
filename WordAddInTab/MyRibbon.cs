using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddInTab
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnImport_Click(object sender, RibbonControlEventArgs e)
        {
            int i = 0;
            int j = 0;
            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc";
            Word._Document objDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Table objTable2;
            Word.Range wrdRng2 = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            objTable2 = objDoc.Tables.Add(wrdRng2, 2, 4, ref oMissing, ref oMissing);
            objTable2.Range.ParagraphFormat.SpaceAfter = 0;
            // TH, 20Feb21, to set borders to the table
            objTable2.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            objTable2.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            objTable2.Range.Font.Size = 11f;
            string strText2;
            for (i = 1; i <= 2; i++)
                for (j = 1; j <= 4; j++)
                {
                    strText2 = "Row" + i + " Column" + j;
                    objTable2.Cell(i, j).Range.Text = strText2;
                }
            objTable2.Rows[1].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightGreen;
            objTable2.Rows[1].HeadingFormat = -1;
            objTable2.ApplyStyleHeadingRows = true;

            //Insert a paragraph at the end of the document.
            Word.Paragraph objPara3; //define paragraph object
            object oRng2 = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range; //go to end of the page
            objPara3 = objDoc.Content.Paragraphs.Add(ref oRng2); //add paragraph at end of document
            objPara3.Range.Text = " Continue with...  "; //add some text in paragraph
            objPara3.Range.Font.Size = 18f;
            objPara3.Format.SpaceAfter = 0; //defind some style

            Word.Table objTable;
            Word.Range wrdRng = objDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            objTable = objDoc.Tables.Add(wrdRng, 100, 4, ref oMissing, ref oMissing);
            objTable.Range.ParagraphFormat.SpaceAfter = 0;
            objTable.Range.Font.Size = 11f;
            // TH, 20Feb21, to set borders to the table
            objTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            objTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            string strText;
            for (i = 1; i <= 100; i++)
                for (j = 1; j <= 4; j++)
                {
                    strText = "Row" + i + " Column" + j;
                    objTable.Cell(i, j).Range.Text = strText;
                }
            objTable.Rows[1].Range.Font.Bold = 1;
            objTable.Rows[1].Range.Font.Italic = 1;
            // TH, 20Feb21, to set the background color of the Header of a table
            objTable.Rows[1].Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorLightBlue;

            objTable.Rows[1].HeadingFormat = -1;
            objTable.ApplyStyleHeadingRows = true;
        }
    }
}
