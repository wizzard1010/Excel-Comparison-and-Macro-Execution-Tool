using System;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Style;


namespace codetest1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        // Task 1: Excel file path and btnClick event
        private void btnCompare_Click(object sender, EventArgs e)
        {
            try
            {
                //Set the licenseContext to NonCommercial or Commercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                string filePath = txtFilePath.Text;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var workbook = package.Workbook;
                    var trueSheet = workbook.Worksheets["TRUE"];
                    var toBeCheckSheet = workbook.Worksheets["TobeCheck"];

                    // loop through cells in the ToBeChecked sheet
                    for (int row = 1; row <= toBeCheckSheet.Dimension.Rows; row++)
                    {
                        for (int col = 1; col <= toBeCheckSheet.Dimension.Columns; col++)
                        {
                            var trueCell = trueSheet.Cells[row, col];
                            var toBeCheckCell = toBeCheckSheet.Cells[row, col];

                            if (!trueCell.Text.Equals(toBeCheckCell.Text))
                            {
                                //Set the background color to yellow for incorrect cells
                                toBeCheckCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                toBeCheckCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);

                            }
                        }
                    }
                    //Save the modified Excel file
                    package.Save();
                }

                MessageBox.Show("Comparison completed, and Excel file saved.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An Error Ocurred: {ex.Message}");
            }
        }
        private void txtFilePath_GotFocus(object sender, RoutedEventArgs e)
        {
            //Clear the initial text when the TextBox gets focus
            if (txtFilePath.Text == "Enter Excel File Path")
            {
                txtFilePath.Text = "";
            }
        }
        // Task2 : Excel Macro Execution Code
        private void btnRunMacro_Click(object sneder, RoutedEventArgs e)
        {
            try
            {
                // Set the path to excel file
                string filePath = txtFilePath.Text;

                // Create a new instance of Excel
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = true;

                // Open the excel file
                var workbooks = excelApp.Workbooks;
                var workbook = workbooks.Open(filePath);

                //Run the macro
                string macroName = txtMacroName.Text;
                excelApp.Run(macroName);

                //Close and save the workbook
                workbook.Close(true);

                //Release the Excel COM objects
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                MessageBox.Show("macro executed successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An Error Occured: {ex.Message}");
            }
        }
    }
}