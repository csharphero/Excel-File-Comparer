using System.Windows.Forms;

namespace Gs.Excel.Comparer
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            this.gsfPanel1 = new Panel();
            this.btnClose = new Button();
            this.gsfLabel3 = new Label();
            this.btnExecute = new Button();
            this.btnSelectDest = new Button();
            this.cmbDestColIndex = new ComboBox();
            this.destItemBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.excelColnfoBindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.gsfLabel10 = new Label();
            this.gsfLabel9 = new Label();
            this.gsfPanel2 = new Panel();
            this.gsfLabel1 = new Label();
            this.btnSelectSourceFiles = new Button();
            this.txtSource = new System.Windows.Forms.TextBox();
            this.sourceItemBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gsfLabel5 = new Label();
            this.gsfLabel2 = new Label();
            this.cbSourceAddress = new ComboBox();
            this.excelColnfoBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.cmbSheetName = new ComboBox();
            this.txtDestFile = new System.Windows.Forms.TextBox();
            this.gsfLabel6 = new Label();
            this.gsfLabel7 = new Label();
            this.cmbDestSheetName = new ComboBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.gsfPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.destItemBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelColnfoBindingSource1)).BeginInit();
            this.gsfPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sourceItemBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelColnfoBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // gsfPanel1
            // 
            this.gsfPanel1.Controls.Add(this.btnClose);
            this.gsfPanel1.Controls.Add(this.gsfLabel3);
            this.gsfPanel1.Controls.Add(this.btnExecute);
            this.gsfPanel1.Controls.Add(this.btnSelectDest);
            this.gsfPanel1.Controls.Add(this.cmbDestColIndex);
            this.gsfPanel1.Controls.Add(this.gsfLabel10);
            this.gsfPanel1.Controls.Add(this.gsfLabel9);
            this.gsfPanel1.Controls.Add(this.gsfPanel2);
            this.gsfPanel1.Controls.Add(this.txtDestFile);
            this.gsfPanel1.Controls.Add(this.gsfLabel6);
            this.gsfPanel1.Controls.Add(this.gsfLabel7);
            this.gsfPanel1.Controls.Add(this.cmbDestSheetName);
            this.gsfPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gsfPanel1.Font = new System.Drawing.Font("Arial", 15.75F);
            this.gsfPanel1.Location = new System.Drawing.Point(0, 0);
            this.gsfPanel1.Margin = new System.Windows.Forms.Padding(4);
            this.gsfPanel1.Name = "gsfPanel1";
            this.gsfPanel1.Size = new System.Drawing.Size(732, 607);
            this.gsfPanel1.TabIndex = 0;
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClose.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btnClose.Location = new System.Drawing.Point(16, 550);
            this.btnClose.Margin = new System.Windows.Forms.Padding(4);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(120, 42);
            this.btnClose.TabIndex = 17;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // gsfLabel3
            // 
            this.gsfLabel3.AutoSize = true;
            this.gsfLabel3.Location = new System.Drawing.Point(19, 473);
            this.gsfLabel3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel3.Name = "gsfLabel3";
            this.gsfLabel3.Size = new System.Drawing.Size(108, 32);
            this.gsfLabel3.TabIndex = 16;
            this.gsfLabel3.Text = "Column";
            // 
            // btnExecute
            // 
            this.btnExecute.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExecute.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.btnExecute.Location = new System.Drawing.Point(596, 550);
            this.btnExecute.Margin = new System.Windows.Forms.Padding(4);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(120, 42);
            this.btnExecute.TabIndex = 16;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // btnSelectDest
            // 
            this.btnSelectDest.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelectDest.Location = new System.Drawing.Point(573, 350);
            this.btnSelectDest.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectDest.Name = "btnSelectDest";
            this.btnSelectDest.Size = new System.Drawing.Size(60, 38);
            this.btnSelectDest.TabIndex = 4;
            this.btnSelectDest.Text = "...";
            this.btnSelectDest.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSelectDest.UseVisualStyleBackColor = true;
            this.btnSelectDest.Click += new System.EventHandler(this.btnSelectDest_Click);
            // 
            // cmbDestColIndex
            // 
            this.cmbDestColIndex.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.destItemBindingSource, "PhoneNumberIndex", true));
            this.cmbDestColIndex.DataSource = this.excelColnfoBindingSource1;
            this.cmbDestColIndex.DisplayMember = "ColumnName";
            this.cmbDestColIndex.FormattingEnabled = true;
            this.cmbDestColIndex.Location = new System.Drawing.Point(193, 469);
            this.cmbDestColIndex.Margin = new System.Windows.Forms.Padding(4);
            this.cmbDestColIndex.Name = "cmbDestColIndex";
            this.cmbDestColIndex.Size = new System.Drawing.Size(257, 39);
            this.cmbDestColIndex.TabIndex = 14;
            this.cmbDestColIndex.ValueMember = "ColumnIndex";
            // 
            // destItemBindingSource
            // 
            this.destItemBindingSource.DataSource = typeof(Gs.Excel.Comparer.Structures.Classes.ExcelFileItem);
            // 
            // excelColnfoBindingSource1
            // 
            this.excelColnfoBindingSource1.DataSource = typeof(Gs.Utility.Excel.Structure.Classes.ExcelColnfo);
            // 
            // gsfLabel10
            // 
            this.gsfLabel10.AutoSize = true;
            this.gsfLabel10.Location = new System.Drawing.Point(19, 409);
            this.gsfLabel10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel10.Name = "gsfLabel10";
            this.gsfLabel10.Size = new System.Drawing.Size(165, 32);
            this.gsfLabel10.TabIndex = 10;
            this.gsfLabel10.Text = "Sheet Name";
            // 
            // gsfLabel9
            // 
            this.gsfLabel9.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.gsfLabel9.Dock = System.Windows.Forms.DockStyle.Top;
            this.gsfLabel9.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gsfLabel9.ForeColor = System.Drawing.Color.White;
            this.gsfLabel9.Location = new System.Drawing.Point(0, 289);
            this.gsfLabel9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel9.Name = "gsfLabel9";
            this.gsfLabel9.Size = new System.Drawing.Size(732, 42);
            this.gsfLabel9.TabIndex = 15;
            this.gsfLabel9.Text = "Destination";
            this.gsfLabel9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // gsfPanel2
            // 
            this.gsfPanel2.Controls.Add(this.gsfLabel1);
            this.gsfPanel2.Controls.Add(this.btnSelectSourceFiles);
            this.gsfPanel2.Controls.Add(this.txtSource);
            this.gsfPanel2.Controls.Add(this.gsfLabel5);
            this.gsfPanel2.Controls.Add(this.gsfLabel2);
            this.gsfPanel2.Controls.Add(this.cbSourceAddress);
            this.gsfPanel2.Controls.Add(this.cmbSheetName);
            this.gsfPanel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.gsfPanel2.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gsfPanel2.Location = new System.Drawing.Point(0, 42);
            this.gsfPanel2.Margin = new System.Windows.Forms.Padding(4);
            this.gsfPanel2.Name = "gsfPanel2";
            this.gsfPanel2.Size = new System.Drawing.Size(732, 247);
            this.gsfPanel2.TabIndex = 13;
            // 
            // gsfLabel1
            // 
            this.gsfLabel1.AutoSize = true;
            this.gsfLabel1.Location = new System.Drawing.Point(16, 36);
            this.gsfLabel1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel1.Name = "gsfLabel1";
            this.gsfLabel1.Size = new System.Drawing.Size(134, 27);
            this.gsfLabel1.TabIndex = 15;
            this.gsfLabel1.Text = "Source File";
            // 
            // btnSelectSourceFiles
            // 
            this.btnSelectSourceFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold);
            this.btnSelectSourceFiles.Location = new System.Drawing.Point(604, 31);
            this.btnSelectSourceFiles.Margin = new System.Windows.Forms.Padding(4);
            this.btnSelectSourceFiles.Name = "btnSelectSourceFiles";
            this.btnSelectSourceFiles.Size = new System.Drawing.Size(55, 37);
            this.btnSelectSourceFiles.TabIndex = 2;
            this.btnSelectSourceFiles.Text = "...";
            this.btnSelectSourceFiles.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSelectSourceFiles.UseVisualStyleBackColor = true;
            this.btnSelectSourceFiles.Click += new System.EventHandler(this.btnSelectSourceFiles_Click);
            // 
            // txtSource
            // 
            this.txtSource.BackColor = System.Drawing.SystemColors.Window;
            this.txtSource.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sourceItemBindingSource, "FilePath", true));
            this.txtSource.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSource.Location = new System.Drawing.Point(193, 31);
            this.txtSource.Margin = new System.Windows.Forms.Padding(4);
            this.txtSource.Name = "txtSource";
            this.txtSource.ReadOnly = true;
            this.txtSource.Size = new System.Drawing.Size(401, 36);
            this.txtSource.TabIndex = 14;
            // 
            // sourceItemBindingSource
            // 
            this.sourceItemBindingSource.DataSource = typeof(Gs.Excel.Comparer.Structures.Classes.ExcelFileItem);
            // 
            // gsfLabel5
            // 
            this.gsfLabel5.AutoSize = true;
            this.gsfLabel5.Location = new System.Drawing.Point(16, 102);
            this.gsfLabel5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel5.Name = "gsfLabel5";
            this.gsfLabel5.Size = new System.Drawing.Size(145, 27);
            this.gsfLabel5.TabIndex = 10;
            this.gsfLabel5.Text = "Sheet Name";
            // 
            // gsfLabel2
            // 
            this.gsfLabel2.AutoSize = true;
            this.gsfLabel2.Location = new System.Drawing.Point(19, 177);
            this.gsfLabel2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel2.Name = "gsfLabel2";
            this.gsfLabel2.Size = new System.Drawing.Size(96, 27);
            this.gsfLabel2.TabIndex = 6;
            this.gsfLabel2.Text = "Column";
            // 
            // cbSourceAddress
            // 
            this.cbSourceAddress.DataBindings.Add(new System.Windows.Forms.Binding("SelectedValue", this.sourceItemBindingSource, "PhoneNumberIndex", true));
            this.cbSourceAddress.DataSource = this.excelColnfoBindingSource;
            this.cbSourceAddress.DisplayMember = "ColumnName";
            this.cbSourceAddress.FormattingEnabled = true;
            this.cbSourceAddress.Location = new System.Drawing.Point(193, 174);
            this.cbSourceAddress.Margin = new System.Windows.Forms.Padding(4);
            this.cbSourceAddress.Name = "cbSourceAddress";
            this.cbSourceAddress.Size = new System.Drawing.Size(257, 35);
            this.cbSourceAddress.TabIndex = 11;
            this.cbSourceAddress.ValueMember = "ColumnIndex";
            // 
            // excelColnfoBindingSource
            // 
            this.excelColnfoBindingSource.DataSource = typeof(Gs.Utility.Excel.Structure.Classes.ExcelColnfo);
            // 
            // cmbSheetName
            // 
            this.cmbSheetName.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.sourceItemBindingSource, "SheetName", true));
            this.cmbSheetName.FormattingEnabled = true;
            this.cmbSheetName.Location = new System.Drawing.Point(193, 98);
            this.cmbSheetName.Margin = new System.Windows.Forms.Padding(4);
            this.cmbSheetName.Name = "cmbSheetName";
            this.cmbSheetName.Size = new System.Drawing.Size(401, 35);
            this.cmbSheetName.TabIndex = 9;
            this.cmbSheetName.SelectedValueChanged += new System.EventHandler(this.cmbSheetName_SelectedValueChanged);
            // 
            // txtDestFile
            // 
            this.txtDestFile.BackColor = System.Drawing.SystemColors.Window;
            this.txtDestFile.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.destItemBindingSource, "FilePath", true));
            this.txtDestFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDestFile.Location = new System.Drawing.Point(193, 350);
            this.txtDestFile.Margin = new System.Windows.Forms.Padding(4);
            this.txtDestFile.Name = "txtDestFile";
            this.txtDestFile.Size = new System.Drawing.Size(371, 36);
            this.txtDestFile.TabIndex = 12;
            // 
            // gsfLabel6
            // 
            this.gsfLabel6.AutoSize = true;
            this.gsfLabel6.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gsfLabel6.Location = new System.Drawing.Point(16, 350);
            this.gsfLabel6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel6.Name = "gsfLabel6";
            this.gsfLabel6.Size = new System.Drawing.Size(119, 29);
            this.gsfLabel6.TabIndex = 11;
            this.gsfLabel6.Text = "Excel File";
            // 
            // gsfLabel7
            // 
            this.gsfLabel7.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.gsfLabel7.Dock = System.Windows.Forms.DockStyle.Top;
            this.gsfLabel7.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.gsfLabel7.ForeColor = System.Drawing.Color.White;
            this.gsfLabel7.Location = new System.Drawing.Point(0, 0);
            this.gsfLabel7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.gsfLabel7.Name = "gsfLabel7";
            this.gsfLabel7.Size = new System.Drawing.Size(732, 42);
            this.gsfLabel7.TabIndex = 14;
            this.gsfLabel7.Text = "Source";
            this.gsfLabel7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cmbDestSheetName
            // 
            this.cmbDestSheetName.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.destItemBindingSource, "SheetName", true));
            this.cmbDestSheetName.FormattingEnabled = true;
            this.cmbDestSheetName.Location = new System.Drawing.Point(193, 405);
            this.cmbDestSheetName.Margin = new System.Windows.Forms.Padding(4);
            this.cmbDestSheetName.Name = "cmbDestSheetName";
            this.cmbDestSheetName.Size = new System.Drawing.Size(371, 39);
            this.cmbDestSheetName.TabIndex = 9;
            this.cmbDestSheetName.SelectedIndexChanged += new System.EventHandler(this.cmbDestSheetName_SelectedIndexChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excels Files|*.xlsx;*.xls";
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Please select the source excel files";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(732, 607);
            this.Controls.Add(this.gsfPanel1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MainForm";
            this.Text = "Excel Request Comparer";
            this.gsfPanel1.ResumeLayout(false);
            this.gsfPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.destItemBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelColnfoBindingSource1)).EndInit();
            this.gsfPanel2.ResumeLayout(false);
            this.gsfPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sourceItemBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.excelColnfoBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Panel gsfPanel1;
        private Button btnSelectSourceFiles;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private Label gsfLabel2;
        private System.Windows.Forms.BindingSource sourceItemBindingSource;
        private Label gsfLabel5;
        private ComboBox cmbSheetName;
        private Button btnSelectDest;
        private Label gsfLabel9;
        private Panel gsfPanel2;
        private TextBox txtDestFile;
        private Label gsfLabel6;
        private Label gsfLabel7;
        private Button btnExecute;
        private Label gsfLabel10;
        private ComboBox cmbDestSheetName;
        private ComboBox cmbDestColIndex;
        private Button btnClose;
        private Label gsfLabel3;
        private Label gsfLabel1;
        private TextBox txtSource;
        private ComboBox cbSourceAddress;
        private System.Windows.Forms.BindingSource destItemBindingSource;
        private System.Windows.Forms.BindingSource excelColnfoBindingSource1;
        private System.Windows.Forms.BindingSource excelColnfoBindingSource;
    }
}

