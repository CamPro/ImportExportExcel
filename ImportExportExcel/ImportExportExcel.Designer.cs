namespace ImportExportExcel
{
    partial class ImportExportExcel
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.buttonExportDataTableToExcel = new System.Windows.Forms.Button();
            this.buttonExportDataGridViewToExcel = new System.Windows.Forms.Button();
            this.buttonExportDataBaseToExcel = new System.Windows.Forms.Button();
            this.buttonAsposeImportExcel = new System.Windows.Forms.Button();
            this.buttonAsposeExportExcel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(291, 13);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(534, 388);
            this.dataGridView1.TabIndex = 0;
            // 
            // buttonExportDataTableToExcel
            // 
            this.buttonExportDataTableToExcel.Location = new System.Drawing.Point(13, 13);
            this.buttonExportDataTableToExcel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.buttonExportDataTableToExcel.Name = "buttonExportDataTableToExcel";
            this.buttonExportDataTableToExcel.Size = new System.Drawing.Size(270, 64);
            this.buttonExportDataTableToExcel.TabIndex = 1;
            this.buttonExportDataTableToExcel.Text = "Export DataTable To Excel";
            this.buttonExportDataTableToExcel.UseVisualStyleBackColor = true;
            this.buttonExportDataTableToExcel.Click += new System.EventHandler(this.buttonExportDataTableToExcel_Click);
            // 
            // buttonExportDataGridViewToExcel
            // 
            this.buttonExportDataGridViewToExcel.Location = new System.Drawing.Point(13, 85);
            this.buttonExportDataGridViewToExcel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonExportDataGridViewToExcel.Name = "buttonExportDataGridViewToExcel";
            this.buttonExportDataGridViewToExcel.Size = new System.Drawing.Size(270, 64);
            this.buttonExportDataGridViewToExcel.TabIndex = 2;
            this.buttonExportDataGridViewToExcel.Text = "Export DataGridView To Excel";
            this.buttonExportDataGridViewToExcel.UseVisualStyleBackColor = true;
            this.buttonExportDataGridViewToExcel.Click += new System.EventHandler(this.buttonExportDataGridViewToExcel_Click);
            // 
            // buttonExportDataBaseToExcel
            // 
            this.buttonExportDataBaseToExcel.Location = new System.Drawing.Point(13, 157);
            this.buttonExportDataBaseToExcel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonExportDataBaseToExcel.Name = "buttonExportDataBaseToExcel";
            this.buttonExportDataBaseToExcel.Size = new System.Drawing.Size(270, 64);
            this.buttonExportDataBaseToExcel.TabIndex = 3;
            this.buttonExportDataBaseToExcel.Text = "Export DataBase To Excel";
            this.buttonExportDataBaseToExcel.UseVisualStyleBackColor = true;
            this.buttonExportDataBaseToExcel.Click += new System.EventHandler(this.buttonExportDataBaseToExcel_Click);
            // 
            // buttonAsposeImportExcel
            // 
            this.buttonAsposeImportExcel.Location = new System.Drawing.Point(13, 229);
            this.buttonAsposeImportExcel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAsposeImportExcel.Name = "buttonAsposeImportExcel";
            this.buttonAsposeImportExcel.Size = new System.Drawing.Size(208, 64);
            this.buttonAsposeImportExcel.TabIndex = 4;
            this.buttonAsposeImportExcel.Text = "Aspose Import Excel";
            this.buttonAsposeImportExcel.UseVisualStyleBackColor = true;
            this.buttonAsposeImportExcel.Click += new System.EventHandler(this.buttonAsposeImportExcel_Click);
            // 
            // buttonAsposeExportExcel
            // 
            this.buttonAsposeExportExcel.Location = new System.Drawing.Point(13, 301);
            this.buttonAsposeExportExcel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAsposeExportExcel.Name = "buttonAsposeExportExcel";
            this.buttonAsposeExportExcel.Size = new System.Drawing.Size(208, 64);
            this.buttonAsposeExportExcel.TabIndex = 5;
            this.buttonAsposeExportExcel.Text = "Aspose Export Excel";
            this.buttonAsposeExportExcel.UseVisualStyleBackColor = true;
            this.buttonAsposeExportExcel.Click += new System.EventHandler(this.buttonAsposeExportExcel_Click);
            // 
            // ImportExportExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(839, 410);
            this.Controls.Add(this.buttonAsposeExportExcel);
            this.Controls.Add(this.buttonAsposeImportExcel);
            this.Controls.Add(this.buttonExportDataBaseToExcel);
            this.Controls.Add(this.buttonExportDataGridViewToExcel);
            this.Controls.Add(this.buttonExportDataTableToExcel);
            this.Controls.Add(this.dataGridView1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "ImportExportExcel";
            this.Text = "ImportExportExcel";
            this.Load += new System.EventHandler(this.ImportExportExcel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button buttonExportDataTableToExcel;
        private System.Windows.Forms.Button buttonExportDataGridViewToExcel;
        private System.Windows.Forms.Button buttonExportDataBaseToExcel;
        private System.Windows.Forms.Button buttonAsposeImportExcel;
        private System.Windows.Forms.Button buttonAsposeExportExcel;
    }
}

