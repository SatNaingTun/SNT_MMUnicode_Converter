using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;




namespace SNT_MMUnicode_Converter
{
    public partial class Form1 : Form
    {
       
        public Form1()
        {
            InitializeComponent();
        }

        private void Convert_Click(object sender, EventArgs e)
        {
            WordDoc doc = new WordDoc();
            ExcelDoc exl = new ExcelDoc();

          
            for (int i = 0; i < dataGridView1.RowCount-1; i++)
            {
                label3.Visible = true;
                string filename=dataGridView1.Rows[i].Cells[0].Value.ToString();
                label3.Text = "Running :" + filename;
                string path = dataGridView1.Rows[i].Cells[1].Value.ToString();
                string fullname = System.IO.Path.Combine(path, filename);
                if (filename.Contains(".docx") || filename.Contains(".doc"))
                    doc.change(fullname, OutputPath.Text, filename);
               // Doc_change(fullname, filename);
                if (filename.Contains(".xlsx") || filename.Contains(".xls"))
                    exl.Change(fullname, OutputPath.Text, filename);

                //Excel_Change(fullname, filename);
                
                //MessageBox.Show(filename.Substring(filename.Length - 4,4));
            
                
            }
           
            label3.Visible = false;
           //Doc_change(@"D:\Cycle counter.docx","test3.docx");
            //Excel_Change(@"E:\office letter\Member Card List.xlsx", "Member Card List.xlsx");
          
            //Doc_change(@"E:\office letter\Cycle counter.docx", "Cycle counter");
           
            //doc.change(@"E:\office letter\Cycle counter.docx",OutputPath.Text, "Cycle counter");

           
         
           

        }
        
        
        

        private void InputBtn_Click(object sender, EventArgs e)
        {
            InputPath.Text = Browse_fun();
        }

        private void OutputBtn_Click(object sender, EventArgs e)
        {
            OutputPath.Text = Browse_fun();
        }
     
        private void add_List(string [] files) {
           // listFileName.Items.Clear();

            

            foreach (string file in files)
            {
                FileInfo info= new FileInfo(file);
              if (((info.Attributes & FileAttributes.Hidden) == 0) & ((info.Attributes & FileAttributes.System) == 0))
                {
                    string[] col = new string[3];

                    col[0] = Path.GetFileName(file);
                    col[1] = Path.GetDirectoryName(file) + @"\";
                    string[] row = new string[] { col[0], col[1] };
                    dataGridView1.Rows.Add(row);
                }

               //ListViewItem item = new ListViewItem(col[0]);
               //item.Tag = file;

                //listFileName.Items.Add(item);

            }
        
        }
     
        private string Browse_fun()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();

            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "Select a folder";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    //add_List(dlg.SelectedPath);
                   //string [] files = Directory.GetFiles(dlg.SelectedPath);

                    string []extension = {"*.docx","*.doc","*.xlsx","*.xls","*.txt"};
                    foreach(string searchPattern in extension)
                    {
                        //MessageBox.Show(searchPattern);
                        string[] files = Directory.GetFiles(dlg.SelectedPath, searchPattern);
                    add_List( files);
                    }

                    return dlg.SelectedPath + @"\";

                }
                else
                    return "D:\\";
            }
        }
    }
}
