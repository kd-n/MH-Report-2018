﻿Imports System.IO
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Imports System.Runtime.InteropServices

Public Class frm_erajSettings
    Dim objXLApp As Excel.Application
    Dim erajPath As String = ""

    Private Sub btnUpload_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        Dim openFile As OpenFileDialog = New OpenFileDialog
        ' Set filter options and filter index.
        openFile.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx"
        openFile.FilterIndex = 1
        ' Call ShowDialog.
        Dim result As DialogResult = openFile.ShowDialog()
        ' Test result.
        If result = Windows.Forms.DialogResult.OK Then
            ' check the file type
            If openFile.FileName.Contains(".xls") OrElse erajPath.Contains(".xlsx") Then
                ' save the directory path if an excel file
                erajPath = openFile.FileName
                txtTemplate.Text = erajPath
            End If
            If Not openFile.FileName.Contains(".xls") AndAlso Not erajPath.Contains(".xlsx") Then
                ' display error message if not an excel file
                MessageBox.Show("Incorrect file type.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' save the directory path if an excel file
                erajPath = openFile.FileName
                txtTemplate.Text = erajPath
            End If
        End If
    End Sub

    Private Sub txtTemplate_Changed(sender As Object, e As EventArgs) Handles txtTemplate.TextChanged
        If Len(txtTemplate.Text) > 0 Then
            btnSaveTemplate.Enabled = True
        Else
            btnSaveTemplate.Enabled = False
        End If
    End Sub

    Private Sub btnSaveTemplate_Click(sender As Object, e As EventArgs) Handles btnSaveTemplate.Click
        btnUpload.Enabled = False
        btnSaveTemplate.Enabled = False
        btnBack.Enabled = False

        objXLApp = New Excel.Application
        objXLApp.DisplayAlerts = False
        ' Dim objXLWb, objXLWb2 As Excel.Workbook
        Dim objXLWb As Excel.Workbook
        Dim erajXLPath As String = Path.GetFullPath(My.Application.Info.DirectoryPath) + "\resources\erajTemplate"

        objXLWb = objXLApp.Workbooks.Open(erajPath)
        'objXLWb2 = objXLApp.Workbooks.Open(erajXLPath)

        Dim found As Boolean
        'Dim newExcelsheet, xlSheet As Excel.Worksheet

        'For Each xs As Excel.Worksheet In objXLWb.Sheets
        '    'If xs.Name = "Sheet" Then
        '    If xs.Name = "Sheet1" Then
        '        found = True
        '        Exit For
        '    End If
        'Next

        If objXLWb.Worksheets.Count >= 1 Then
            found = True
        Else
            found = False
        End If

        objXLWb.Close(SaveChanges:=True)

        If found = True Then

            ' Copy the file to resources folder and save path to be used for initial eraj generation

            'xLsheet = objXLWb2.Worksheets.Add()
            'xlSheet.Name = "Sheet2"
            'objXLWb2.Close(SaveChanges:=True)
            'objXLWb2 = objXLApp1.Workbooks.Open(erajXLPath)
            'objXLWb2.Worksheets("Sheet1").Delete()
            'objXLWb.Worksheets("Sheet1").Copy(, objXLWb2.Worksheets(1))
            'newExcelsheet.Name = "Sheet1"
            'objXLApp.DisplayAlerts = True
            'objXLApp.DisplayAlerts = False
            'objXLWb2.Close()
            'Marshal.ReleaseComObject(objXLWb2)
            'objXLWb2 = Nothing

            My.Computer.FileSystem.CopyFile(erajPath, erajXLPath, True)
            My.Settings.ErajPath = erajPath
            My.Settings.Save()

            MessageBox.Show("ERAJ template successfully updated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' MessageBox.Show("Template Missing. (Check if template sheet name is Sheet1)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MessageBox.Show("Template Missing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        'objXLWb.Close()
        Marshal.ReleaseComObject(objXLWb)
        objXLWb = Nothing
        objXLApp.Quit()
        Marshal.ReleaseComObject(objXLApp)
        objXLApp = Nothing
        'txtTemplate.Clear()

        btnUpload.Enabled = True
        'btnSaveTemplate.Enabled = True
        btnBack.Enabled = True
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.Close()
    End Sub

    Private Sub frm_erajSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtTemplate.ReadOnly = True
        txtTemplate.Text = My.Settings.ErajPath
    End Sub
End Class