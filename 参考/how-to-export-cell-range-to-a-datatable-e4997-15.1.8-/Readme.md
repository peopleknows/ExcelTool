<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/ExportToDataTableExample/Form1.cs) (VB: [Form1.vb](./VB/ExportToDataTableExample/Form1.vb))
<!-- default file list end -->
# How to export a cell range to a DataTable


This example illustrates how you can export a cell range to a System.Data.DataTable object.

The following steps are required:

1) Add a reference to the **DevExpress.Docs.dll** assembly to your Spreadsheet project. The distribution of this assembly requires <a href="https://www.devexpress.com/products/net/office-file-api/">a license to the DevExpress Office File API or DevExpress Universal Subscription</a>.

2) Use the **DevExpress.Spreadsheet.Worksheet.CreateDataTableExporter** method to create a **DevExpress.Spreadsheet.Export.DataTableExporter** instance.

3) Call the DataTableExporter's **Export** method.

You can use the **Worksheet.CreateDataTable** method to create an empty DataTable from a cell range. This method obtains column names from the range headings, and determines the column data types based on the first row of the specified range.
