using Aspose.Cells;
using Syncfusion.XlsIO;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ImportExportExcel
{
    public partial class ImportExportExcel : Form
    {
        public ImportExportExcel()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void ImportExportExcel_Load(object sender, EventArgs e)
        {
            #region Loading the data to Data Grid
            DataSet customersDataSet = new DataSet();
            //Get the path of the input file
            string inputXmlPath = Path.GetFullPath(@"../../Data/Employees.xml");
            customersDataSet.ReadXml(inputXmlPath);
            DataTable dataTable = new DataTable();
            //Copy the structure and data of the table
            dataTable = customersDataSet.Tables[1].Copy();
            //Removing unwanted columns
            dataTable.Columns.RemoveAt(0);
            dataTable.Columns.RemoveAt(10);
            this.dataGridView1.DataSource = dataTable;


            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.White;
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.LightBlue;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 9F, ((System.Drawing.FontStyle)(System.Drawing.FontStyle.Bold)));
            dataGridView1.ForeColor = Color.Black;
            dataGridView1.BorderStyle = BorderStyle.None;

            #endregion

        }

        private void ExportDataTableToExcel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a dataset from XML file
                DataSet customersDataSet = new DataSet();
                customersDataSet.ReadXml(Path.GetFullPath(@"../../Data/Employees.xml"));

                //Create datatable from the dataset
                DataTable dataTable = new DataTable();
                dataTable = customersDataSet.Tables[1];

                //Import data from the data table with column header, at first row and first column, and by its column type.
                sheet.ImportDataTable(dataTable, true, 1, 1, true);

                //Creating Excel table or list object and apply style to the table
                IListObject table = sheet.ListObjects.Create("Employee_PersonalDetails", sheet.UsedRange);

                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium14;

                //Autofit the columns
                sheet.UsedRange.AutofitColumns();

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(@"Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
                Process.Start(Path.GetFullPath(@"Output.xlsx"));
            }
        }

        private void ExportDataGridViewToExcel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                //Create a workbook with single worksheet
                IWorkbook workbook = application.Workbooks.Create(1);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Import from DataGridView to worksheet
                worksheet.ImportDataGridView(dataGridView1, 1, 1, true, true);

                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs("Output.xlsx");
                Process.Start("Output.xlsx");
            }
        }

        private void ExportDataBaseToExcel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                if (sheet.ListObjects.Count == 0)
                {
                    //Estabilishing the connection in the worksheet
                    string dBPath = Path.GetFullPath(@"../../Data/Employee.mdb");
                    string ConnectionString = "OLEDB;Provider=Microsoft.JET.OLEDB.4.0;Password=\"\";User ID=Admin;Data Source=" + dBPath;
                    string query = "SELECT EmployeeID,FirstName,LastName,Title,HireDate,Extension,ReportsTo FROM [Employees]";
                    IConnection Connection = workbook.Connections.Add("Connection1", "Sample connection with MsAccess", ConnectionString, query, ExcelCommandType.Sql);
                    sheet.ListObjects.AddEx(ExcelListObjectSourceType.SrcQuery, Connection, sheet.Range["A1"]);
                }

                //Refresh Excel table to get updated values from database
                sheet.ListObjects[0].Refresh();

                sheet.UsedRange.AutofitColumns();

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(@"Output.xlsx"));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
                Process.Start(Path.GetFullPath(@"Output.xlsx"));
            }
        }

        private void buttonExportDataTableToExcel_Click(object sender, EventArgs e)
        {
            ExportDataTableToExcel();
        }

        private void buttonExportDataGridViewToExcel_Click(object sender, EventArgs e)
        {
            ExportDataGridViewToExcel();
        }

        private void buttonExportDataBaseToExcel_Click(object sender, EventArgs e)
        {
            ExportDataBaseToExcel();
        }

        OpenFileDialog ofdExcel = new OpenFileDialog();
        SaveFileDialog sfdExcel = new SaveFileDialog();

        private void AsposeImportExcel()
        {
            ofdExcel.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            DialogResult fileExcel = ofdExcel.ShowDialog();

            if (fileExcel == DialogResult.OK)
            {
                string pathFileExcel = ofdExcel.FileName;
                //load file excel
                Workbook workbook = new Workbook(pathFileExcel);
                //chon sheet dau tien
                Worksheet worksheet = workbook.Worksheets[0];
                //lấy vung du lieu
                Range range = worksheet.Cells.MaxDisplayRange;
                //day du lieu vao datatable, ý nghĩa tham số (từ dòng, từ cột, tổng số dòng lấy, tổng số cột lấy, tự đặt tên cột theo dòng đầu tiên)
                //link tham khảo: https://docs.aspose.com/display/cellsnet/Export+Data+from+Worksheet
                DataTable dataTable = worksheet.Cells.ExportDataTable(0, 0, range.RowCount, range.ColumnCount, true);
                dataGridView1.DataSource = dataTable;
            }
        }

        private void AsposeExportExcel()
        {
            string template = AppDomain.CurrentDomain.BaseDirectory + @"..\..\Data\" + "TemplateExport2.xlsx";
            template = Path.GetFullPath(template);

            WorkbookDesigner wbDesigner = new WorkbookDesigner();
            //Lấy template
            wbDesigner.Workbook = new Workbook(template);
            Worksheet ws = wbDesigner.Workbook.Worksheets[0];
            //fill du lieu don le 
            //cấu hình cell co format: &=$<<ten_tu_dat>>
            wbDesigner.SetDataSource("NGAY_XUAT", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
            //fill table
            //cấu hình cột có dạng: &=<<ten_data_source_tu_dat>>.<<ten_cot_tuong_ung_trong_data_source>>
            wbDesigner.SetDataSource("DTS", (dataGridView1.DataSource as DataTable).DefaultView);
            //fill du lieu don le
            wbDesigner.SetDataSource("TOTAL_ROWS", (dataGridView1.DataSource as DataTable).Rows.Count);
            wbDesigner.Process();

            sfdExcel.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            sfdExcel.FileName = "Export";
            if (sfdExcel.ShowDialog() == DialogResult.OK)
            {
                wbDesigner.Workbook.Save(sfdExcel.FileName);
            }
        }

        private void buttonAsposeImportExcel_Click(object sender, EventArgs e)
        {
            AsposeImportExcel();
        }

        private void buttonAsposeExportExcel_Click(object sender, EventArgs e)
        {
            AsposeExportExcel();
        }
    }
}
