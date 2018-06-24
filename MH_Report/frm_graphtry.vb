﻿Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices

Public Class frm_graphtry
    Dim erajPath As String = ""
    Dim totalsPath As String = ""


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button1.Enabled = False
        btn_UploadEraj.Enabled = False
        btn_uploadSource.Enabled = False
        txt_totSource.Enabled = False
        txtbox_Eraj.Enabled = False

        Dim objExcelApplication As New Excel.Application
        objExcelApplication.DisplayAlerts = False

        Dim objXLApp As New Excel.Application
        Dim objXLWb2 As Excel.Workbook = objXLApp.Workbooks.Add
        objXLApp.DisplayAlerts = False

        Dim teamName, jsysJob, nonJsysJob, jpiJob As Object
        Dim teamMHArray As New List(Of Object())
        Dim teamInfo(3) As Object

        Dim outPath As String = My.Settings.OutputFolder
        Dim filename As String = outPath + "\" + "FETP ERAJ" + "_" + Format(Now, "yyyyMMdd") + "_" + Format(Now, "hhmmss") + ".xlsx"
        objXLWb2.SaveAs(filename)

        ' open source excel file
        If System.IO.File.Exists(totalsPath) Then
            Dim bookSource As Excel.Workbook = objExcelApplication.Workbooks.Open(totalsPath)
            Dim sheetSource As Excel.Worksheet = bookSource.Worksheets("Summary")

            Dim sourceRow As Integer = 4

            While CStr(DirectCast(sheetSource.Cells(sourceRow, 1), Excel.Range).Value) <> ""
                ReDim teamInfo(3)

                teamName = CObj(DirectCast(sheetSource.Cells(sourceRow, 1), Excel.Range).Value)
                jsysJob = CObj(DirectCast(sheetSource.Cells(sourceRow, 2), Excel.Range).Value)
                nonJsysJob = CObj(DirectCast(sheetSource.Cells(sourceRow, 3), Excel.Range).Value)
                jpiJob = CObj(DirectCast(sheetSource.Cells(sourceRow, 4), Excel.Range).Value)

                teamInfo = {teamName, jsysJob, nonJsysJob, jpiJob}
                teamMHArray.Add(teamInfo)
                sourceRow += 1
            End While

            Dim tsYear As String = CStr(DirectCast(sheetSource.Cells(1, 1), Excel.Range).Value).Substring(4)
            Dim tsMonth As String = CStr(DirectCast(sheetSource.Cells(1, 1), Excel.Range).Value).Split(" ")(0)

            objExcelApplication.DisplayAlerts = False
            bookSource.Close()


            If System.IO.File.Exists(erajPath) Then
                Dim erajObjXLWB As Excel.Workbook = objXLApp.Workbooks.Open(erajPath)
                Dim sheetname As String = "FETP ERAJ " + tsYear

                Dim found As Boolean
                For Each xs As Excel.Worksheet In erajObjXLWB.Sheets
                    If xs.Name = sheetname Then
                        found = True
                        Exit For
                    End If
                Next

                Dim newExcelsheet As Excel.Worksheet

                If found = True Then
                    erajObjXLWB.Worksheets(sheetname).Copy(, objXLWb2.Worksheets(objXLWb2.Worksheets.Count))
                Else
                    ' Use worksheet number instead of name
                    ' erajObjXLWB.Worksheets("FETP ERAJ TEMP").Copy(, objXLWb2.Worksheets(objXLWb2.Worksheets.Count))
                    erajObjXLWB.Worksheets(1).Copy(, objXLWb2.Worksheets(objXLWb2.Worksheets.Count))
                End If

                newExcelsheet = objXLWb2.Worksheets(objXLWb2.Worksheets.Count)
                newExcelsheet = objXLWb2.ActiveSheet
                newExcelsheet.Name = sheetname

                objXLApp.DisplayAlerts = True

                Dim destrow As Integer = 6
                sourceRow = 6
                teamName = ""
                Dim monthCol As Integer = 1

                For no As Integer = 0 To 13
                    If CStr(DirectCast(newExcelsheet.Cells(4, monthCol), Excel.Range).Value) = tsMonth.ToUpper Then
                        Exit For
                    Else
                        monthCol += 1
                    End If
                Next

                For i As Integer = 0 To teamMHArray.Count - 1
                    sourceRow = 6
                    Do While CStr(DirectCast(newExcelsheet.Cells(sourceRow, 1), Excel.Range).Value) <> "FETP TOTAL"
                        teamName = CStr(DirectCast(newExcelsheet.Cells(sourceRow, 1), Excel.Range).Value).ToUpper
                        If teamName = teamMHArray(i)(0).ToString.ToUpper Then
                            newExcelsheet.Cells(sourceRow, monthCol) = teamMHArray(i)(1)
                            newExcelsheet.Cells(sourceRow + 1, monthCol) = teamMHArray(i)(2)
                            newExcelsheet.Cells(sourceRow + 2, monthCol) = teamMHArray(i)(3)
                            Exit Do
                        End If
                        sourceRow += 5
                    Loop
                Next

                objXLApp.DisplayAlerts = False
                erajObjXLWB.Close()
                Marshal.ReleaseComObject(erajObjXLWB)
                erajObjXLWB = Nothing
            End If
        End If

        ' Delete first sheet after generating eraj
        objXLWb2.Worksheets(1).Delete

        objXLWb2.Save()
        teamMHArray.Clear()
        MessageBox.Show("ERAJ successfully generated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        objXLWb2.Close()
        Marshal.ReleaseComObject(objXLWb2)
        objXLWb2 = Nothing
        'releaseObject(objXLWb2)

        objXLApp.Quit()
        Marshal.ReleaseComObject(objXLApp)
        objXLApp = Nothing

        Button1.Enabled = True
        btn_UploadEraj.Enabled = True
        btn_uploadSource.Enabled = True
        txt_totSource.Enabled = True
        txtbox_Eraj.Enabled = True
    End Sub

    Private Sub btn_uploadSource_Click(sender As Object, e As EventArgs) Handles btn_uploadSource.Click
        Dim openFile As OpenFileDialog = New OpenFileDialog


        ' Set filter options and filter index.
        openFile.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx"
        openFile.FilterIndex = 1

        ' Call ShowDialog.
        Dim result As DialogResult = openFile.ShowDialog()

        ' Test result.
        If result = Windows.Forms.DialogResult.OK Then
            ' check the file type
            If openFile.FileName.Contains(".xls") OrElse openFile.FileName.Contains(".xlsx") Then
                ' save the directory path if an excel file
                totalsPath = openFile.FileName
                txt_totSource.Text = totalsPath
            End If

            If Not openFile.FileName.Contains(".xls") AndAlso Not openFile.FileName.Contains(".xlsx") Then
                ' display error message if not an excel file
                MessageBox.Show("Incorrect file type.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' save the directory path if an excel file
                totalsPath = openFile.FileName
                txt_totSource.Text = totalsPath
            End If
        End If
    End Sub

    Private Sub btn_UploadEraj_Click(sender As Object, e As EventArgs) Handles btn_UploadEraj.Click
        Dim openFile As OpenFileDialog = New OpenFileDialog


        ' Set filter options and filter index.
        openFile.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx"
        openFile.FilterIndex = 1

        ' Call ShowDialog.
        Dim result As DialogResult = openFile.ShowDialog()

        ' Test result.
        If result = Windows.Forms.DialogResult.OK Then
            ' check the file type
            If openFile.FileName.Contains(".xls") OrElse openFile.FileName.Contains(".xlsx") Then
                ' save the directory path if an excel file
                erajPath = openFile.FileName
                txtbox_Eraj.Text = erajPath
            End If

            If Not openFile.FileName.Contains(".xls") AndAlso Not openFile.FileName.Contains(".xlsx") Then
                ' display error message if not an excel file
                MessageBox.Show("Incorrect file type.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' save the directory path if an excel file
                erajPath = openFile.FileName
                txtbox_Eraj.Text = erajPath
            End If
        End If
    End Sub

    Private Sub frm_graphtry_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class