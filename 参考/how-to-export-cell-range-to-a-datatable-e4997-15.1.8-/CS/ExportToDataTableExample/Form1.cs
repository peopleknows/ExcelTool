#region #usings
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
#endregion #usings
using System;
using System.Data;
using System.Windows.Forms;

namespace ExportToDataTableExample
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public Form1()
        {
            InitializeComponent();
            spreadsheetControl1.LoadDocument("TopTradingPartners.xlsx");
            ribbonControl1.SelectedPage = exportDataExampleRibbonPage;
        }

        private void barButtonItemRangeToDataTable_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (barCheckItemStopEmptyRow.Checked) {
                ExportSelectionStopOnEmptyRow();
                return;
            }
            #region #SimpleDataExport
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets.ActiveWorksheet;
            Range range = worksheet.Selection;
            bool rangeHasHeaders = this.barCheckItemHasHeaders1.Checked;
            
            // Create a data table with column names obtained from the first row in a range if it has headers.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, rangeHasHeaders);

            //Validate cell value types. If cell value types in a column are different, the column values are exported as text.
            for (int col = 0; col < range.ColumnCount; col++)
            {
                CellValueType cellType = range[0, col].Value.Type;
                for (int r = 1; r < range.RowCount; r++)
                {
                    if (cellType != range[r, col].Value.Type)
                    {
                        dataTable.Columns[col].DataType = typeof(string);
                        break;
                    }
                }
            }

            // Create the exporter that obtains data from the specified range, 
            // skips the header row (if required) and populates the previously created data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, rangeHasHeaders);
            // Handle value conversion errors.
            exporter.CellValueConversionError += exporter_CellValueConversionError;

            // Perform the export.
            exporter.Export();
            #endregion #SimpleDataExport
            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

        private void ExportSelectionStopOnEmptyRow() {
            #region #StopExportOnEmptyRow
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets.ActiveWorksheet;
            Range range = worksheet.Selection;
            // Determine whether the first row in a range contains headers.
            bool rangeHasHeaders = this.barCheckItemHasHeaders1.Checked;
            // Determine whether an empty row must stop conversion.
            bool stopOnEmptyRow = barCheckItemStopEmptyRow.Checked;

            // Create a data table with column names obtained from the first row in a range if it has headers.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, rangeHasHeaders);
            // Create the exporter that obtains data from the specified range, 
            // skips the header row (if required) and populates the previously created data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, rangeHasHeaders);
            // Handle value conversion errors.
            exporter.CellValueConversionError += (sender,args)=> {args.Action = DataTableExporterAction.Continue;};
            if (stopOnEmptyRow) {
                exporter.Options.SkipEmptyRows = false;
                // Handle empty row.
                exporter.ProcessEmptyRow += (sender, args) => { args.Action = DataTableExporterAction.Stop; };
            }
            // Perform the export.
            exporter.Export();
            #endregion #StopExportOnEmptyRow
            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

        private void barButtonItemUseExporterOptions_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            #region #DataExportWithOptions
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
            Range range = worksheet.Tables[0].Range;
            
            // Create a data table with column names obtained from the first row in a range.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, true);
            
            // Create the exporter that obtains data from the specified range which has a header row and populates the previously created data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
            // Handle value conversion errors.
            exporter.CellValueConversionError += exporter_CellValueConversionError;
            
            // Specify exporter options.
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.EmptyCellValue = 0;
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = barCheckItemSkipErrors.Checked;

            // Perform the export.
            exporter.Export();
            #endregion #DataExportWithOptions
            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

        #region #DataExportWithCustomConverter
        private void barButtonItemUseCustomConverter_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            Worksheet worksheet = spreadsheetControl1.Document.Worksheets[0];
            Range range = worksheet.Tables[0].Range;
            
            // Create a data table with column names obtained from the first row in a range.
            // Column data types are obtained from cell value types of cells in the first data row of the worksheet range.
            DataTable dataTable = worksheet.CreateDataTable(range, true);
            // Change the data type of the "As Of" column to text.
            dataTable.Columns["As Of"].DataType = System.Type.GetType("System.String");
            // Create the exporter that obtains data from the specified range and populates the specified data table. 
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
            // Handle value conversion errors.
            exporter.CellValueConversionError += exporter_CellValueConversionError;

            // Specify a custom converter for the "As Of" column.
            DateTimeToStringConverter toDateStringConverter = new DateTimeToStringConverter();
            exporter.Options.CustomConverters.Add("As Of", toDateStringConverter);
            // Set the export value for empty cell.
            toDateStringConverter.EmptyCellValue = "N/A";
            // Specify that empty cells and cells with errors should be processed.
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = false;
            
            // Perform the export.
            exporter.Export();

            // A custom method that displays the resulting data table.
            ShowResult(dataTable);
        }

        // A custom converter that converts DateTime values to "Month-Year" text strings.
        class DateTimeToStringConverter : ICellValueToColumnTypeConverter
        {
            public bool SkipErrorValues { get; set; }
            public CellValue EmptyCellValue { get; set; }

            public ConversionResult Convert(Cell readOnlyCell, CellValue cellValue, Type dataColumnType, out object result)
            {
                result = DBNull.Value; 
                ConversionResult converted = ConversionResult.Success;
                if (cellValue.IsEmpty) {
                    result = EmptyCellValue;
                    return converted;
                }
                if (cellValue.IsError) {
                    // You can return an error, subsequently the exporter throws an exception if the CellValueConversionError event is unhandled.
                    //return SkipErrorValues ? ConversionResult.Success : ConversionResult.Error;
                    result = "N/A";
                    return ConversionResult.Success;
                }
                result =  String.Format("{0:MMMM-yyyy}",cellValue.DateTimeValue);
                return converted;
            }
        }
        #endregion #DataExportWithCustomConverter

        #region #CellValueConversionErrorHandler
        void exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e)
        {
            MessageBox.Show("Error in cell " + e.Cell.GetReferenceA1());
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
        }
        #endregion #CellValueConversionErrorHandler

        #region #ShowResultForm
        Form ShowResult(DataTable result)
        {
            Form newForm = new Form();
            newForm.Width = 600;
            newForm.Height = 300;

            DevExpress.XtraGrid.GridControl grid = new DevExpress.XtraGrid.GridControl();
            grid.Dock = DockStyle.Fill;
            grid.DataSource = result;

            newForm.Controls.Add(grid);
            grid.ForceInitialize();
            ((DevExpress.XtraGrid.Views.Grid.GridView)grid.FocusedView).OptionsView.ShowGroupPanel = false;

            newForm.ShowDialog(this);
            return newForm;
        }
        #endregion #ShowResultForm
    }
}