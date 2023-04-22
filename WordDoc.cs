using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MsWord = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace SNT_MMUnicode_Converter
{
    class WordDoc
    {
        MsWord._Document documentFrom = null; 
        object oMissing = System.Reflection.Missing.Value;



        public void change(string fileNameFrom,String OutputPath, string fileNameTo)
        {
            //var fileNameFrom = @"D:\Cardreader.docx";
            MsWord._Application wordApp = new MsWord.Application();
            wordApp.Visible = false;
            string input, output;


            try
            {
                documentFrom = wordApp.Documents.Add(fileNameFrom, Type.Missing, false);

                Microsoft.Office.Interop.Word.Paragraphs DocPar = documentFrom.Paragraphs;
                long parCount = DocPar.Count;

                //table range
                List<MsWord.Range> TablesRanges = new List<MsWord.Range>();
                for (int iCounter = 1; iCounter <= documentFrom.Tables.Count; iCounter++)
                {
                    MsWord.Range TRange = documentFrom.Tables[iCounter].Range;
                    TablesRanges.Add(TRange);
                }
                Boolean bInTable;




                // second doc
                MsWord._Document DocumentTo = wordApp.Documents.Add();
                MsWord.Paragraph objPara;
                objPara = DocumentTo.Paragraphs.Add();

                // Step through the paragraphs
                for (int i = 1; i < parCount; i++)
                {
                    bInTable = false;
                    MsWord.Range r = DocPar[i].Range;

                    foreach (MsWord.Range tbrange in TablesRanges)
                    {

                        if (r.Start >= tbrange.Start && r.Start <= tbrange.End)// tbrange.Start<=r.start&& r.Start<=tbrange.End
                        {
                            if (r.Start == tbrange.End)
                                doc_createtable(DocumentTo, objPara, tbrange);


                            bInTable = true;
                            break;
                        }

                        /*
                        foreach (Table tbl in documentFrom.Range(tbrange.Start, tbrange.End).Tables)
                        {

                            //doc_createtable(documentFrom, DocumentTo,tbrange);
                            int rowCount = tbl.Rows.Count;
                            int colCount = tbl.Columns.Count;
                            doc_createtable(DocumentTo, objPara, rowCount, colCount, tbl);
                            // MessageBox.Show("rowCount" + rowCount + " " + "colCount" + colCount);
                        }
                         * */
                    }

                    if (!bInTable)
                    {

                        input = DocPar[i].Range.Text;
                        output = Rabbit.Zg2Uni(input);


                        // objPara = DocumentTo.Paragraphs.Add();
                        objPara.Range.Text = output;
                        objPara.Range.Font.Name = "Pyidaungsu";
                    }



                }




                string dirTarget = System.IO.Path.Combine(OutputPath, fileNameTo);
                DocumentTo.SaveAs(dirTarget);
                DocumentTo.Close();
                documentFrom.Close();
                wordApp.Quit();

            }
            catch (Exception ex)
            {
                // Console.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
            }



        }
        private void doc_createtable(MsWord._Document documentTo, MsWord.Paragraph objParaTo, MsWord.Range tbrange)
        {




            //Microsoft.Office.Interop.Word.Paragraph para1 = documentTo.Content.Paragraphs.Add(ref oMissing);
            //object styleHeading1 = "Heading 1";
            //para1.Range.set_Style(ref styleHeading1);
            //objPara.Range.Text = "Para 1 text";
            //objParaTo.Range.InsertParagraphAfter();
            foreach (Table tbl in documentFrom.Range(tbrange.Start, tbrange.End).Tables)
            {

                //doc_createtable(documentFrom, DocumentTo,tbrange);
                int rowCount = tbl.Rows.Count;
                int colCount = tbl.Columns.Count;

                // MessageBox.Show("rowCount" + rowCount + " " + "colCount" + colCount);

                //MessageBox.Show("R"+rowCount+"C" + colCount);
                Table newtable = documentTo.Tables.Add(objParaTo.Range, rowCount, colCount, ref oMissing, ref oMissing);

                newtable.Borders.Enable = 1;
                /*
                foreach (Row row2 in newtable.Rows)
                {
                    foreach (Cell cell2 in row2.Cells)
                    {
                        //Header row  
                  
                        if (cell2.RowIndex == 1)
                        {
                            cell2.Range.Text = "Column " + cell2.ColumnIndex.ToString();
                            cell2.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell2.Range.Font.Name = "verdana";
                            cell2.Range.Font.Size = 10;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell2.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                            //Center alignment for the Header cells  
                            cell2.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        //Data row  
                        else
                        {
                            cell2.Range.Text = (cell2.RowIndex - 2 + cell2.ColumnIndex).ToString();
                        }
                    

                    }
                }
                 */
                string input, output;
                // MsWord.Cell cell;
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        var copyFrom = tbl.Cell(r, c).Range;
                        input = copyFrom.Text;
                        output = Rabbit.Zg2Uni(input);
                        newtable.Cell(r, c).Range.Text = output;


                    }
                }
            }









        }
        private void doc_createtable_foreach(MsWord.Document DocumentTo, MsWord.Paragraph objPara, MsWord.Range r)
        {
            string input, output;
            foreach (Table tbl in documentFrom.Range(r.Start, r.End).Tables)
            {

                Table newtable = DocumentTo.Tables.Add(objPara.Range, tbl.Rows.Count, tbl.Columns.Count, ref oMissing, ref oMissing);
                newtable.Borders.Enable = 1;
                MessageBox.Show("NEw Table");
                foreach (Row row in tbl.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        input = cell.Range.Text;
                        output = Rabbit.Zg2Uni(input);
                        newtable.Cell(row.Index, cell.ColumnIndex).Range.Text = output;
                    }
                }
            }

        }
        
        
        
        
        
        
        
        
        private static void MsWordCopy()
        {
            object nullobject = Type.Missing;
            //var wordApp = new MsWord.Application();
            MsWord._Application wordApp = new MsWord.Application();
            //wordApp.Visible = false;
            MsWord._Document documentFrom = null, documentTo = null;


            try
            {
                var fileNameFrom = @"D:\Cardreader.docx";

                wordApp.Visible = true;

                //documentFrom = wordApp.Documents.Open(fileNameFrom, Type.Missing, true);
                documentFrom = wordApp.Documents.Add(fileNameFrom, Type.Missing, true);

                MsWord.Range oRange = documentFrom.Content;
                oRange.Copy();

                var fileNameTo = @"D:\MyDocFile-Copy.docx";
                documentTo = wordApp.Documents.Add();
                documentTo.Content.PasteSpecial(DataType: MsWord.WdPasteOptions.wdKeepSourceFormatting);
                documentTo.SaveAs(fileNameTo);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            finally
            {
                if (documentFrom != null)
                    documentFrom.Close(MsWord.WdSaveOptions.wdDoNotSaveChanges);

                if (documentTo != null)
                    documentTo.Close();

                if (wordApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);


                wordApp = null;
                documentFrom = null;
                documentTo = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                //wordApp.Quit();


            }
        }




        public void removeStyle(Microsoft.Office.Interop.Word.Document document, string styleName)
        {
            Microsoft.Office.Interop.Word.Range rng = document.Range();

            Microsoft.Office.Interop.Word.Style style = null;

            foreach (Microsoft.Office.Interop.Word.Style currentStyle in document.Styles)
            {
                if (currentStyle.NameLocal == "Only CZ")
                    style = currentStyle;
            }

            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();

            rng.Find.set_Style(style);
            rng.Find.Forward = true;
            rng.Find.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop;

            //change this property to true as we want to replace format
            rng.Find.Format = true;

            rng.Find.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
        }
    }
}
