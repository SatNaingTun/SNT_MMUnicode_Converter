using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace SNT_MMUnicode_Converter
{
    class ExcelDoc
    {

        Microsoft.Office.Interop.Excel._Application excelApp = null;
        Microsoft.Office.Interop.Excel._Application excelApp2 = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel._Workbook wbSource = null;
        Microsoft.Office.Interop.Excel._Workbook wbTarget = null;

        Microsoft.Office.Interop.Excel._Worksheet wsSource = null;
       Microsoft.Office.Interop.Excel._Worksheet wsTarget = null;
      // List<Microsoft.Office.Interop.Excel.Range> mergerange = new List<Microsoft.Office.Interop.Excel.Range>();
       

        public void Change(string sourceFilePath, string OutputPath, string fileNameTo)
        {




            try
            {

                excelApp = new Microsoft.Office.Interop.Excel.Application();
                
                wbSource = excelApp.Workbooks.Add(sourceFilePath);
                //wbTarget = excelApp.Workbooks.Open(destinationFilePath);
                wbTarget = excelApp2.Workbooks.Add(Type.Missing);
                /*
                for (int i = 2; i <= wbTarget.Sheets.Count; i++)
                {
                    wbTarget.Sheets[i].Delete();
                }
                 */ 
              
                //wsSource = wbSource.Worksheets["Sheet1"];
                int wscount = wbSource.Sheets.Count;
                //MessageBox.Show("No of worksheet" + wscount);
               //List<string> worksheetName=new List<string>();
                for (int wsIndex = 1; wsIndex <= wscount; wsIndex++)
                {
                    wsSource = wbSource.Worksheets.get_Item(wsIndex);
                    

                   // worksheetName.Add(wsSource.Name);
                    string wsName = wsSource.Name;
                    Microsoft.Office.Interop.Excel.Range theRange = (Microsoft.Office.Interop.Excel.Range)wsSource.UsedRange;
                   // wsTarget = wbTarget.ActiveSheet;
                  
                        //wsTarget = wbTarget.ActiveSheet;
                        UpdateCellValue(theRange, wsIndex, "Pyidaungsu", OutputPath, fileNameTo);
                                      
                    wsSource = null;
                    
                    if (wsIndex == wscount)
                    {
                        string dirTarget = System.IO.Path.Combine(OutputPath, fileNameTo);
                        wbTarget.SaveAs(dirTarget);


                    }
                }
                
                /*
                wsSource = wbSource.Worksheets.get_Item(2);
                Microsoft.Office.Interop.Excel.Range theRange = (Microsoft.Office.Interop.Excel.Range)wsSource.UsedRange;
                UpdateCellValue(theRange, 2, "Pyidaungsu", OutputPath, fileNameTo);
                wsSource = null;
               */
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excelApp.Quit();
                wbSource = null;

                excelApp2.Quit();
                wbTarget = null;
                //wsTarget = null;

            }
        }
        private void UpdateCellValue(Microsoft.Office.Interop.Excel.Range excelRange,  int wsIndex, string FontName, string OutputPath, string fileNameTo)
        {

            int RowCount = excelRange.Rows.Count;
            int ColumnCount = excelRange.Columns.Count;
            string output;
            int mcount;

           // Microsoft.Office.Interop.Excel.Range mergeRange = excelRange.MergeArea;
          
           
           /*foreach (Microsoft.Office.Interop.Excel.Range m in wsSource.Range.MergeArea)
        {
            mergerange.Add(m);
        }
          

           /* string beforeMergeAddress = excelRange.MergeArea.get_Address(Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1);
            MessageBox.Show(beforeMergeAddress);
            */ 
          
          
            
      
           // Microsoft.Office.Interop.Excel.Range rng ;
            

            //MessageBox.Show(mergeRange.Address);
          
            //Microsoft.Office.Interop.Excel._Worksheet wsTarget;



            try
            {
                //wsTarget = wbTarget.Worksheets.get_Item(wsIndex);
                  
                 //
                wbTarget.Worksheets.Add();
                wsTarget = wbTarget.ActiveSheet;
               
                // wsTarget.Name = wsSource.Name;


                //wsTarget = (Microsoft.Office.Interop.Excel.Worksheet)wbTarget.Worksheets[wsIndex];
                        
                //wsTarget = wbTarget.Sheets[wsIndex];
                /*
                if (wsIndex == 1)
                {
                    wsTarget = wbTarget.ActiveSheet;
                }
                else
                {
                    wsTarget=wbTarget.Worksheets.Add(Type.Missing,
                        wbTarget.Worksheets[wsIndex], Type.Missing, Type.Missing);
                }
                 * */
                //int mergecellcount = 0;
                                    
               
                for (int r = 1; r <= RowCount; r++)
                {

                    for (int c = 1; c <= ColumnCount; c++)
                    {
                        dynamic cell = excelRange.Cells[r, c];


                        Microsoft.Office.Interop.Excel.Range Range = (Microsoft.Office.Interop.Excel.Range)wsSource.Cells[r, c];
                        
                        /*
                                                Microsoft.Office.Interop.Excel.Range R1 = (Microsoft.Office.Interop.Excel.Range)cell;
                                                R1.Copy(Type.Missing);
                         * 

    
                         * * */

                       // if (cell.MergeCells) {

                         //  string sth = excelRange.MergeArea.get_Address(Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1);
                          //MessageBox.Show(sth);
                        
                        //}
                        
                        
                                 
                      
                        
                        string content;

                        if (cell.Value2 != null)
                        {
                            content = cell.Value2.ToString();

                            //cell.copy(Type.Missing);

                            //string content = cell.Value;
                            Microsoft.Office.Interop.Excel.Borders xborder = cell.Borders;
                            //xborder.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                            

                            if (FontName == "Pyidaungsu")
                            {
                                output = Rabbit.Zg2Uni(content);
                                wsTarget.Cells[r, c].value = output;
                                //cell.Font.Name = FontName;
                                wsTarget.Cells[r, c].Font.Name = FontName;
                               
                                wsTarget.Cells[r, c].RowHeight = cell.RowHeight;
                                wsTarget.Cells[r, c].ColumnWidth = cell.ColumnWidth;
                                wsTarget.Cells[r, c].font.Underline = cell.font.Underline;
                                wsTarget.Cells[r, c].font.Italic = cell.font.Italic;
                                wsTarget.Cells[r, c].font.color = cell.font.color;
                                wsTarget.Cells[r, c].NumberFormat = cell.NumberFormat;
                                if (cell.font.Bold == true) { wsTarget.Cells[r, c].font.Bold = true; }
                                wsTarget.Cells[r, c].Borders.linestyle = cell.Borders.linestyle;


                                if (cell.MergeCells)
                                {

                                    int x = Int32.Parse(Range.MergeArea.Rows.Count.ToString());
                                    int y = Int32.Parse(Range.MergeArea.Columns.Count.ToString());
                                   // int mx = x;
                                    //int my = y;
                                    /*for (int i = r; i <= (r + x - 1); i++)
                                        for (int j = c; j <= (c + y - 1); j++)
                                            if (wsSource.Cells[i, j].value != null)
                                            {
                                                mx = i;
                                                my = j;
                                                break;
                                            }
                                     */ 

                                    /*     int mx = 1;
                                         int my = 1;
                                         for (int i = r; i <= RowCount; i++) 
                                             for(int j=c;j<=ColumnCount;j++)
                                                 if (wsSource.Cells[i, j].value != null)
                                                 {
                                                     mx = i;
                                                     my = j;
                                                     break;
                                                
                                                 }
                                     * */

                                    Microsoft.Office.Interop.Excel.Range c1 = wsTarget.Cells[r, c];
                                    Microsoft.Office.Interop.Excel.Range c2 = wsTarget.Cells[r + (x - 1), c + (y - 1)];


                                   

                                    Microsoft.Office.Interop.Excel.Range mrange = (Microsoft.Office.Interop.Excel.Range)wsTarget.get_Range(c1, c2);
                                   mrange.Merge(true);
                                   mrange.Borders.LineStyle = Range.Borders.LineStyle;
                                   mrange.HorizontalAlignment = Range.HorizontalAlignment;
                                   mrange.VerticalAlignment = Range.VerticalAlignment;
                                    

                                    //MergeAndCenter(mrange);


                                    // MessageBox.Show("Row" + x + "Column" + y);

                                    //Microsoft.Office.Interop.Excel.Range mrange = (Microsoft.Office.Interop.Excel.Range)wsSource.get_Range(c1, cr);




                                    /*   //mergecellcount++;
                                     if (excelRange.Cells[r + 1, c].MergeCells)
                                      {
                                          Microsoft.Office.Interop.Excel.Range c1 = wsTarget.Cells[r, c];
                                          Microsoft.Office.Interop.Excel.Range c2 = wsTarget.Cells[r + 1, c];
                                          Microsoft.Office.Interop.Excel.Range mrange = (Microsoft.Office.Interop.Excel.Range)wsTarget.get_Range(c1, c2);
                                          //Microsoft.Office.Interop.Excel.Range mrange = (Microsoft.Office.Interop.Excel.Range)wsTarget.get_Range(wsTarget.Cells[r, c], wsTarget.Cells[r + 1, c]);
                                           mrange.Merge(true);
                                          mrange.Borders.LineStyle = cell.Borders.linestyle;
                                      }
                                      if (excelRange.Cells[r, c + 1].MergeCells)
                                      {
                                          Microsoft.Office.Interop.Excel.Range c1 = wsTarget.Cells[r, c];
                                          Microsoft.Office.Interop.Excel.Range c2 = wsTarget.Cells[r, c + 1];
                                          Microsoft.Office.Interop.Excel.Range mrange = (Microsoft.Office.Interop.Excel.Range)wsTarget.get_Range(c1, c2);
                                          //Microsoft.Office.Interop.Excel.Range mrange = (Microsoft.Office.Interop.Excel.Range)wsTarget.get_Range(wsTarget.Cells[r, c], wsTarget.Cells[r, c + 1]);
                                           mrange.Merge();
                                      }
                          */


                                }
                                

                                //  wsTarget.Cells[r, c].Style.HorizontalAlignment = cell.Style.HorizontalAlignment;

                                /*
                                                           Microsoft.Office.Interop.Excel.Range R2 = (Microsoft.Office.Interop.Excel.Range)wsTarget.Cells[r, c];
                                                           R2.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats,
                                   Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                                 */


                                //wsTarget.Cells[r, c].Borders = xborder;
                                // wsTarget.Cells.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic);


                            }
                            else if (FontName == "Zawgyi-One")
                            {
                                output = Rabbit.Uni2Zg(content);
                                wsTarget.Cells[r, c].value = output;
                                //cell.Font.Name = FontName;
                                wsTarget.Cells[r, c].Font.Name = FontName;
                                //wsTarget.Cells[r, c].Borders = xborder;
                                wsTarget.Cells[r, c].RowHeight = cell.RowHeight;
                                wsTarget.Cells[r, c].ColumnWidth = cell.ColumnWidth;

                            }
                            
                            /*
                            if (cell.Merged) {

                                MessageBox.Show("Row:" + r + "Column" + c);
                            }
                             */ 

                            // cell.Value2 = output;
                            //xlWorkSheet.Range[xlWorkSheet.Cells[1, 1], xlWorkSheet.Cells[1, 4]].Merge();
                            //range.Merge(true);


                        }

                    }
                }
                if (wsSource.Shapes.Count > 0) {

                    foreach (Microsoft.Office.Interop.Excel.Shape o in wsSource.Shapes)
                    {


                        
                    }
                
                }

              
               

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                wsTarget = null;
                
            }
            





        }
        private static void CopyCharts(Microsoft.Office.Interop.Excel.Workbook wbIn, Microsoft.Office.Interop.Excel.Workbook wbOut)
        {
            Microsoft.Office.Interop.Excel.Worksheet wsOutAfter = (Microsoft.Office.Interop.Excel.Worksheet)wbOut.Sheets["Plot Items"];
            foreach (Microsoft.Office.Interop.Excel.Chart c in wbIn.Charts)
            {
                string chartName = c.Name;
                c.Copy(Type.Missing, wsOutAfter);
                Microsoft.Office.Interop.Excel.SeriesCollection sc = (Microsoft.Office.Interop.Excel.SeriesCollection)c.SeriesCollection();
                foreach (Microsoft.Office.Interop.Excel.Series s in sc)
                {
                    Microsoft.Office.Interop.Excel.Range r = (Microsoft.Office.Interop.Excel.Range)s.XValues;
                    // get string representing range, modify it and set corresponding
                    // series in wbOut.Charts.Item[chartName] to something appropriate
                }
            }
        }

        public void MergeAndCenter(Microsoft.Office.Interop.Excel.Range MergeRange)
        {
            MergeRange.Select();

            MergeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            MergeRange.VerticalAlignment = XlVAlign.xlVAlignBottom;
            MergeRange.WrapText = false;
            MergeRange.Orientation = 0;
            MergeRange.AddIndent = false;
            MergeRange.IndentLevel = 0;
            MergeRange.ShrinkToFit = false;
            MergeRange.ReadingOrder = (int)(Constants.xlContext);
            MergeRange.MergeCells = false;

            MergeRange.Merge(System.Type.Missing);
        }
        /*
        private void Excel_iterate_char(Microsoft.Office.Interop.Excel.Range cellRange)
        {
            // Create a dictionary to save the font settings for each character;
            Dictionary<int, ExcelFont> fontDictionary = new Dictionary<int, ExcelFont>();

            ExcelFont excelFont;

            // Iterate the characters and get their settings:
            for (int i = 0; i < cellRange.Characters.Count; i++)
            {
                excelFont = new ExcelFont();
                excelFont.characters = cellRange.Characters[i, 1];
                excelFont.Bold = excelFont.characters.Font.Bold;
                excelFont.Color = excelFont.characters.Font.Color;
                excelFont.FontStyle = excelFont.characters.Font.FontStyle;
                excelFont.Size = excelFont.characters.Font.Size;
                excelFont.Italic = excelFont.characters.Font.Italic;
                excelFont.Name = excelFont.characters.Font.Name;
                excelFont.Strikethrough = excelFont.characters.Font.Strikethrough;
                excelFont.Subscript = excelFont.characters.Font.Subscript;
                excelFont.Superscript = excelFont.characters.Font.Superscript;
                excelFont.ThemeFont = excelFont.characters.Font.ThemeFont;
                excelFont.Underline = excelFont.characters.Font.Underline;

                fontDictionary.Add(i, excelFont);
            }

            // Assign the text:
            cellRange.Value2 = cellRange.Value2 + "and some more text";

            // Re assign the font for each character:
            for (int i = 0; i < fontDictionary.Count; i++)
            {
                fontDictionary[i].characters.Font.Bold = fontDictionary[i].Bold;
                fontDictionary[i].characters.Font.Color = fontDictionary[i].Color;
                fontDictionary[i].characters.Font.FontStyle = fontDictionary[i].FontStyle;
                fontDictionary[i].characters.Font.Size = fontDictionary[i].Size;
                fontDictionary[i].characters.Font.Italic = fontDictionary[i].Italic;
                fontDictionary[i].characters.Font.Name = fontDictionary[i].Name;
                fontDictionary[i].characters.Font.Strikethrough = fontDictionary[i].Strikethrough;
                fontDictionary[i].characters.Font.Subscript = fontDictionary[i].Subscript;
                fontDictionary[i].characters.Font.Superscript = fontDictionary[i].Superscript;
                fontDictionary[i].characters.Font.ThemeFont = fontDictionary[i].ThemeFont;
                fontDictionary[i].characters.Font.Underline = fontDictionary[i].Underline;
            }
        
        }
         */ 
    }
}
