Imports System.IO
Imports System.File.IO
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration

Module CreateFile
    Public Sub CreateExcel()
        Dim objXLApp As New Excel.Application
        ' Check if user did not enter a timesheet filename
        objXLApp.DisplayAlerts = False

        objXLApp.DisplayFormulaBar = False
        Dim objXLWb As Excel.Workbook
        Dim xltPath As String = Path.GetFullPath(My.Application.Info.DirectoryPath) + "\resources\MHtemplate"
        objXLWb = objXLApp.Workbooks.Open(xltPath)
        Dim objXLWs As Excel.Worksheet = objXLWb.Worksheets("Sheet")


        objXLApp.DisplayAlerts = True
        'Save a copy
        Dim outPath As String = My.Settings.OutputFolder
        Dim fileNames As String = outPath + "\" + Format(Now, "yyyyMMdd") + "_" + Format(Now, "hhmmss") + " .xlsx"
        objXLWb.SaveAs(fileNames)
    End Sub


End Module
