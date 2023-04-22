namespace SNT_MMUnicode_Converter
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.Convert = new System.Windows.Forms.Button();
            this.InputBtn = new System.Windows.Forms.Button();
            this.InputPath = new System.Windows.Forms.TextBox();
            this.OutputPath = new System.Windows.Forms.TextBox();
            this.OutputBtn = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.FileName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FileDir = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // Convert
            // 
            this.Convert.Location = new System.Drawing.Point(89, 206);
            this.Convert.Name = "Convert";
            this.Convert.Size = new System.Drawing.Size(75, 23);
            this.Convert.TabIndex = 4;
            this.Convert.Text = "Convert";
            this.Convert.UseVisualStyleBackColor = true;
            this.Convert.Click += new System.EventHandler(this.Convert_Click);
            // 
            // InputBtn
            // 
            this.InputBtn.Location = new System.Drawing.Point(309, 26);
            this.InputBtn.Name = "InputBtn";
            this.InputBtn.Size = new System.Drawing.Size(75, 23);
            this.InputBtn.TabIndex = 1;
            this.InputBtn.Text = "Browse";
            this.InputBtn.UseVisualStyleBackColor = true;
            this.InputBtn.Click += new System.EventHandler(this.InputBtn_Click);
            // 
            // InputPath
            // 
            this.InputPath.Location = new System.Drawing.Point(93, 29);
            this.InputPath.Name = "InputPath";
            this.InputPath.Size = new System.Drawing.Size(199, 20);
            this.InputPath.TabIndex = 2;
            this.InputPath.Text = "D:\\";
            // 
            // OutputPath
            // 
            this.OutputPath.Location = new System.Drawing.Point(93, 79);
            this.OutputPath.Name = "OutputPath";
            this.OutputPath.Size = new System.Drawing.Size(199, 20);
            this.OutputPath.TabIndex = 3;
            this.OutputPath.Text = "D:\\Uni Output\\";
            // 
            // OutputBtn
            // 
            this.OutputBtn.Location = new System.Drawing.Point(309, 76);
            this.OutputBtn.Name = "OutputBtn";
            this.OutputBtn.Size = new System.Drawing.Size(75, 23);
            this.OutputBtn.TabIndex = 3;
            this.OutputBtn.Text = "Browse";
            this.OutputBtn.UseVisualStyleBackColor = true;
            this.OutputBtn.Click += new System.EventHandler(this.OutputBtn_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Input Directory";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 81);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Output Directory";
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FileName,
            this.FileDir});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Right;
            this.dataGridView1.GridColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.Location = new System.Drawing.Point(509, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(246, 265);
            this.dataGridView1.TabIndex = 5;
            // 
            // FileName
            // 
            this.FileName.HeaderText = "FileName";
            this.FileName.Name = "FileName";
            this.FileName.ReadOnly = true;
            // 
            // FileDir
            // 
            this.FileDir.HeaderText = "FileDir";
            this.FileDir.Name = "FileDir";
            this.FileDir.ReadOnly = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(93, 143);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(199, 16);
            this.progressBar1.TabIndex = 6;
            this.progressBar1.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(93, 124);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 13);
            this.label3.TabIndex = 7;
            this.label3.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(755, 265);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OutputBtn);
            this.Controls.Add(this.OutputPath);
            this.Controls.Add(this.InputBtn);
            this.Controls.Add(this.InputPath);
            this.Controls.Add(this.Convert);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "SNT_MMUnicode_Converter";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Convert;
        private System.Windows.Forms.Button InputBtn;
        private System.Windows.Forms.TextBox InputPath;
        private System.Windows.Forms.TextBox OutputPath;
        private System.Windows.Forms.Button OutputBtn;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn FileName;
        private System.Windows.Forms.DataGridViewTextBoxColumn FileDir;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label3;
    }
}

