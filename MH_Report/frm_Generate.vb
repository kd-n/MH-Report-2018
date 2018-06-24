﻿Imports System.IO
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Imports System.Runtime.InteropServices

Public Class frm_Generate

    Dim objXLApp As Excel.Application

    Dim xltPath As String = Path.GetFullPath(My.Application.Info.DirectoryPath) + "\resources\MHtemplate"
    Dim erajXLPath As String = Path.GetFullPath(My.Application.Info.DirectoryPath) + "\resources\erajTemplate"
    Dim outPath As String = My.Settings.OutputFolder

    Dim fileNames As String
    Dim errCheck As Boolean = False
    Dim sheetcount As Integer = 0
    Dim nErr1 As Integer = 0
    Dim nErr2 As Integer = 0
    Dim nErr3 As Integer = 0
    Dim noTS(200) As String
    Dim jCError(200) As String
    Dim noData(200) As String
    Dim amtPerTeam As New List(Of Object())
    Dim jsystotH As Double = 0
    Dim nonjtotH As Double = 0
    Dim jpitotH As Double = 0
    Dim grandtotH As Double = 0
    Dim updateMH As Boolean = False
    Dim erajError As Boolean = False

    Private Const JSYS_JOBCODE = "FETEC"
    Private Const NONJSYS_JOBCODE = "Non FETEC"
    Private Const JPI_JOBCODE = "FETP"

    Public Sub GenerateOutput(memberlist As List(Of String), filePrefix() As String, GenerateType As String, headerVar As String, teamName As Object, objXLWb2 As Excel.Workbook, objXLWb As Excel.Workbook)

        Dim srcPath As String = My.Settings.SourceFolder
        Dim EmpInfoArray As New List(Of Object())
        Dim jobCodeList As New List(Of Object())
        Dim managersInfoList As New List(Of Object())

        Dim depCode, jobCode, pjName, rt, rot, pot As Object


        Dim strEmpInfo(4) As Object
        Dim strJobcode(9) As Object
        Dim strManagersInfo(3) As Object

        Dim jsystot As Object = 0
        Dim nonjtot As Object = 0
        Dim jpitot As Object = 0
        Dim grandtot As Object = 0
        Dim intH As Integer = 0
        Dim objExcelApplication As New Excel.Application
        objExcelApplication.DisplayAlerts = False
        Dim empfile As String = ""
        If GenerateType <> "Multi" Then
            nErr1 = 0
        End If
        Try

        
        For item As Integer = 0 To memberlist.Count - 1

            Dim empJsysTot As Double = 0
            Dim empJpiTot As Double = 0
            Dim empNonTot As Double = 0
            Dim emplace As String = ""
            Dim empName As String = ""
            Dim empNo As String = ""

            Dim strArrayCompare As New StrArrays(3)
            Dim noTimeS As Boolean = False

            For Int As Integer = 0 To filePrefix.Count - 1
                Dim fileNo(1) As String
                fileNo = {"", "(2)"}
                For i As Integer = 0 To fileNo.Count - 1
                    noTimeS = False
                    intH = Int

                        Dim filename As String = srcPath + "\" + filePrefix(Int) + memberlist(item) + fileNo(i)
                        empfile = filename
                        If System.IO.File.Exists(filename + ".xls") Or System.IO.File.Exists(filename + ".xlsx") Then
                            ReDim strEmpInfo(2)
                            Dim empErrCheck = False
                            Dim source As Excel.Workbook = objExcelApplication.Workbooks.Open(filename, False, True)
                            Dim sheet As Excel.Worksheet = DirectCast(source.Sheets("Form"), Excel.Worksheet)

                            empNo = CObj(DirectCast(sheet.Cells(3, 13), Excel.Range).Value)
                            empName = CObj(DirectCast(sheet.Cells(3, 17), Excel.Range).Value).ToString

                            Dim sourceRow As Integer = 14
                            Do While CObj(DirectCast(sheet.Cells(sourceRow, 3), Excel.Range).Value) <> ""

                                ReDim strJobcode(9)

                                emplace = CObj(DirectCast(sheet.Cells(sourceRow, 1), Excel.Range).Value)
                                depCode = CObj(DirectCast(sheet.Cells(sourceRow, 2), Excel.Range).Value)
                                jobCode = CObj(DirectCast(sheet.Cells(sourceRow, 3), Excel.Range).Value)
                                pjName = CObj(DirectCast(sheet.Cells((sourceRow + 1), 3), Excel.Range).Value) & CObj(DirectCast(sheet.Cells((sourceRow + 2), 3), Excel.Range).Value)
                                rt = CObj(DirectCast(sheet.Cells(sourceRow, 21), Excel.Range).Value)
                                rot = CObj(DirectCast(sheet.Cells((sourceRow + 1), 21), Excel.Range).Value)
                                pot = CObj(DirectCast(sheet.Cells((sourceRow + 2), 21), Excel.Range).Value)

                                'Edited to include new Fetec Jobcodes
                                '  If jobCode.ToString.Substring(1, 1) = "-" And jobCode.ToString.Length = 15 Or jobCode.ToString.Length = 13 Then
                                If jobCode.ToString.Substring(0, 2) = "2-" Or jobCode.ToString.Substring(0, 2) = "0-" Or (jobCode.ToString.Substring(0, 2) = "Z-" And jobCode.ToString.Length = 15) Then

                                    If rt.ToString <> "" And rot.ToString <> "" And pot.ToString <> "" Then
                                        Dim jobCodeType As String = ""
                                        If jobCode.ToString.StartsWith("2") Or jobCode.ToString.StartsWith("0") Then
                                            jobCodeType = JSYS_JOBCODE
                                        ElseIf jobCode.ToString.StartsWith("Z-0") Then
                                            jobCodeType = NONJSYS_JOBCODE
                                        Else
                                            jobCodeType = JPI_JOBCODE
                                        End If

                                        Dim subtotal As Object
                                        If rt.ToString = "" And rot.ToString = "" And pot.ToString = "" Then
                                            subtotal = ""
                                        Else
                                            subtotal = rt + rot + pot
                                            If jobCodeType = JSYS_JOBCODE Then
                                                jsystot += subtotal
                                                empJsysTot += subtotal
                                            ElseIf jobCodeType = NONJSYS_JOBCODE Then
                                                nonjtot += subtotal
                                                empNonTot += subtotal
                                            Else
                                                jpitot += subtotal
                                                empJpiTot += subtotal
                                            End If
                                        End If

                                        If jobCodeList.Count <> 0 Then
                                            If Int = 1 Or i = 1 Then
                                                Dim jCodeFound As Boolean = False
                                                For num As Integer = 0 To jobCodeList.Count - 1
                                                    If jCodeFound = False Then
                                                        If Not jobCodeList(num)(3) Is Nothing And Not jobCodeList(num)(0) Is Nothing Then
                                                            If jobCodeList(num)(2) = jobCode And jobCodeList(num)(3) = pjName And jobCodeList(num)(0) = emplace Then
                                                                jobCodeList(num)(4) += rt
                                                                jobCodeList(num)(5) += rot
                                                                jobCodeList(num)(6) += pot
                                                                jobCodeList(num)(7) += subtotal
                                                                jCodeFound = True
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                                If jCodeFound = False Then
                                                    strJobcode = {emplace, depCode, jobCode, pjName, rt, rot, pot, subtotal, jobCodeType}
                                                    jobCodeList.Add(strJobcode)
                                                End If
                                            Else
                                                strJobcode = {emplace, depCode, jobCode, pjName, rt, rot, pot, subtotal, jobCodeType}
                                                jobCodeList.Add(strJobcode)
                                            End If
                                        Else
                                            strJobcode = {emplace, depCode, jobCode, pjName, rt, rot, pot, subtotal, jobCodeType}
                                            jobCodeList.Add(strJobcode)
                                        End If
                                        sourceRow += 4
                                    End If
                                Else
                                    errCheck = True
                                    empErrCheck = True
                                    jCError(nErr2) = filePrefix(Int) + memberlist(item) + fileNo(i) + " Jobcode: " + jobCode

                                    nErr2 += 1
                                    sourceRow += 4
                                    'Removed function that deletes all jobcodes in timesheet
                                    'For j As Integer = 0 To jobCodeList.Count - 1
                                    '    Dim delsubt As Double = 0
                                    '    Dim delrt As Double = 0
                                    '    Dim delrot As Double = 0
                                    '    Dim delpot As Double = 0

                                    '    delrt = jobCodeList(j)(4)
                                    '    delrot = jobCodeList(j)(5)
                                    '    delpot = jobCodeList(j)(6)

                                    '    delsubt = delrt + delrot + delpot

                                    '    If jobCodeList(j)(8).ToString = JSYS_JOBCODE Then
                                    '        jsystot -= delsubt
                                    '        empJsysTot -= delsubt
                                    '    ElseIf jobCodeList(j)(8).ToString = NONJSYS_JOBCODE Then
                                    '        nonjtot -= delsubt
                                    '        empNonTot -= delsubt
                                    '    Else
                                    '        jpitot -= delsubt
                                    '        empJpiTot -= delsubt
                                    '    End If
                                    'Next
                                    'jobCodeList = New List(Of Object())
                                    'Exit Do
                                End If
                            Loop
                            objExcelApplication.DisplayAlerts = False
                            source.Close()
                            Marshal.ReleaseComObject(source)
                        Else
                            If i = 0 Then
                                noTS(nErr1) = filePrefix(Int) + memberlist(item).ToString
                                nErr1 += 1
                                errCheck = True
                                noTimeS = True
                            End If
                        End If
                    If jobCodeList.Count = 0 Then
                        If i = 0 Then
                            If noTimeS = False Then
                                noData(nErr3) = filePrefix(intH) + memberlist(item).ToString + fileNo(i)
                                nErr3 += 1
                                errCheck = True
                            End If
                        End If
                    End If
                Next
            Next
            If jobCodeList.Count <> 0 Then
                jobCodeList.Sort(strArrayCompare)
                strEmpInfo = {empNo, empName, jobCodeList, teamName, empJsysTot, empNonTot, empJpiTot}
                EmpInfoArray.Add(strEmpInfo)
                jobCodeList = New List(Of Object())
            End If
        Next
        objExcelApplication.DisplayAlerts = True
        objExcelApplication.Quit()
        Marshal.ReleaseComObject(objExcelApplication)
        Catch ex As Exception
            MessageBox.Show("Timesheet Error:" + empfile + " Restart System.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'formulas
        grandtot = jpitot + jsystot + nonjtot
        jsystotH += jsystot
        jpitotH += jpitot
        nonjtotH += nonjtot
        grandtotH += grandtot

        Dim jCodeTot As Object() = {jsystot, nonjtot, jpitot, grandtot}
        If nonjtot = 0 Then
            nonjtot = ""
        End If
        If jpitot = 0 Then
            nonjtot = ""
        End If

        Dim jCodeTots As Object() = {teamName, jsystot, nonjtot, jpitot, grandtot}
        amtPerTeam.Add(jCodeTots)

        If Not EmpInfoArray.Count = 0 Then
            CreateExcel(EmpInfoArray, jCodeTot, GenerateType, headerVar, objXLWb2, objXLWb, teamName)
        End If
    End Sub

    Public Sub CreateExcel(EmpInfoArray As List(Of Object()), jCodeTot As Object(), GenerateType As String, headerVar As String, objXLWb2 As Excel.Workbook, objXLWb As Excel.Workbook, teamName As Object)

        Dim ObjArrayCompare As New ObjArrays(0)
        EmpInfoArray.Sort(ObjArrayCompare)

        Dim found As Boolean = False
        If updateMH = True Then
            For Each xs As Excel.Worksheet In objXLWb2.Sheets
                If xs.Name = EmpInfoArray(0)(3).ToString Then
                    found = True
                    xs.Unprotect("admin")
                    Exit For
                End If
            Next
        End If
        If found = True Then
            Dim sheetNo As Integer = objXLWb2.Sheets(EmpInfoArray(0)(3).ToString).Index()
            objXLWb2.Worksheets(teamName.ToString).Delete()
            objXLWb.Worksheets("Sheet").Copy(, objXLWb2.Worksheets(objXLWb2.Worksheets.Count))
            Dim newExcelsheet As Excel.Worksheet
            newExcelsheet = objXLWb2.Worksheets(sheetNo - 1)
            newExcelsheet.Name = teamName.ToString
            newExcelsheet.Unprotect("admin")
            objXLApp.DisplayAlerts = True
            Plot(newExcelsheet, EmpInfoArray, jCodeTot, headerVar)
            newExcelsheet.Protect("admin")
            objXLApp.DisplayAlerts = True
        Else
            objXLWb.Worksheets("Sheet").Copy(, objXLWb2.Worksheets(objXLWb2.Worksheets.Count))
            Dim newExcelsheet As Excel.Worksheet
            newExcelsheet = objXLWb2.Worksheets(objXLWb2.Worksheets.Count)
            newExcelsheet.Name = teamName.ToString
            newExcelsheet.Unprotect("admin")
            objXLApp.DisplayAlerts = True
            Plot(newExcelsheet, EmpInfoArray, jCodeTot, headerVar)
            newExcelsheet.Protect("admin")
            objXLApp.DisplayAlerts = True
        End If
    End Sub

    Public Class ObjArrays
        Implements IComparer(Of Object())
        Dim empNoIdx As Integer = 0

        Public Sub New(empNoIdxVar As Integer)
            empNoIdx = empNoIdxVar
        End Sub

        Public Function Compare(ByVal x As Object(), ByVal y As Object()) _
           As Integer _
           Implements System.Collections.Generic.IComparer(Of Object()).Compare

            If (Val(x(empNoIdx)) > Val(y(empNoIdx))) Then
                Return 1
            ElseIf (Val(x(empNoIdx)) = Val(y(empNoIdx))) Then
                Return 1
            Else
                Return -1
            End If
        End Function
    End Class

    Public Class StrArrays
        Implements IComparer(Of Object())
        Dim jCodeIdx As Integer = 0

        Public Sub New(jCodeIdxVar As Integer)
            jCodeIdx = jCodeIdxVar
        End Sub

        Public Function Compare(ByVal x As Object(), ByVal y As Object()) _
           As Integer _
           Implements System.Collections.Generic.IComparer(Of Object()).Compare
            Dim c As Integer = 0
            Dim a As String = ""
            If Not x(2).ToString = "" Then
                'Edited to include new Fetec Jobcodes
                ' a = x(2).ToString.Substring(0, 1) & x(2).ToString.Split("-")(1).Substring(0) & x(2).ToString.Split("-")(2).Substring(0)
                a = x(2).ToString
            Else
                a = 0
            End If
            Dim b As String = ""
            If Not y(2).ToString = "" Then
                'Edited to include new Fetec Jobcodes
                'b = y(2).ToString.Substring(0, 1) & y(2).ToString.Split("-")(1).Substring(0, 1) & y(2).ToString.Split("-")(2).Substring(0)
                b = y(2).ToString
            Else
                b = 0
            End If
            c = String.Compare(a, b)
            Return c
        End Function
    End Class

    Public Sub Plot(objXLWsVar As Excel.Worksheet, EmpInfoArray As List(Of Object()), jCodeTot As Object(), headerVar As String)
        'Write on excel
        Dim rowArray(), amountsArrayHolder() As Object
        Dim Rng As Excel.Range
        Dim destrow As Integer = 5
        Dim destrowtot As Integer = 5
        Dim startRow As Integer = 0
        Dim Rowctr As Integer = 0
        Dim amountsCols As Integer = 0
        Dim amountsCole As Integer = 0
        Dim formulastring As String = ""
        Dim formulaRange As String = ""
        Dim c As String = ""
        Dim d As String = ""

        objXLWsVar.Cells(1, 1) = headerVar
        For i As Integer = 0 To EmpInfoArray.Count - 1
            startRow = destrow
            rowArray = EmpInfoArray(i)
            Rng = objXLWsVar.Range(objXLWsVar.Cells(destrow, 1), objXLWsVar.Cells(destrow, 2))
            Rng.Value = rowArray

            For j As Integer = 0 To EmpInfoArray(i)(2).count - 1
                Rng = objXLWsVar.Range(objXLWsVar.Cells(destrow, 3), objXLWsVar.Cells(destrow, 6))
                amountsArrayHolder = (EmpInfoArray(i)(2)(j))
                Rng.Value = amountsArrayHolder
                Dim jCodeType As String = EmpInfoArray(i)(2)(j)(8).ToString
                If jCodeType = JSYS_JOBCODE Then
                    amountsCols = 7
                    amountsCole = 9
                ElseIf jCodeType = NONJSYS_JOBCODE Then
                    amountsCols = 11
                    amountsCole = 13
                Else
                    amountsCols = 15
                    amountsCole = 17
                End If
                Rng = objXLWsVar.Range(objXLWsVar.Cells(destrow, amountsCols), objXLWsVar.Cells(destrow, amountsCole))
                amountsArrayHolder = {EmpInfoArray(i)(2)(j)(4), EmpInfoArray(i)(2)(j)(5), EmpInfoArray(i)(2)(j)(6)}
                Rng.Value = amountsArrayHolder
                destrow += 1
            Next

            objXLWsVar.Cells(destrow, 7) = "Employee Total:"
            Dim hideCol As Integer = 19
            Dim destcol As Integer = 10

            For f As Integer = 4 To 6
                c = Convert.ToChar(destcol + 64)
                formulastring = "=SUM(" + c.ToString + startRow.ToString + ":" + c.ToString + (destrow - 1).ToString + ")"
                formulaRange = c.ToString + destrow.ToString
                Rng = objXLWsVar.Range(formulaRange)
                Rng.Formula = formulastring

                c = Convert.ToChar(hideCol + 64)
                formulaRange = c.ToString + (i + 1).ToString
                Rng = objXLWsVar.Range(formulaRange)
                Rng.Formula = formulastring
                destcol += 4
                hideCol += 1
            Next

            Rng = objXLWsVar.Range(objXLWsVar.Cells(destrow, 7), objXLWsVar.Cells(destrow, 18))
            Rng.Font.Color = Color.Blue
            Rng.Font.Bold = True
            destrow += 2

            If i <> EmpInfoArray.Count - 1 Then
                Rng = objXLWsVar.Range(objXLWsVar.Cells(destrow, 1), objXLWsVar.Cells(destrow, 18))
                Rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                Rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                Rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                Rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            End If
        Next

        Dim totRowCtr As Integer = destrow
        Dim inCtr As Integer = 3
        Dim colCtr As Integer = 10

        formulastring = "=SUM(S1" + ":S" + (EmpInfoArray.Count + 1).ToString + ")"
        formulaRange = "S" + (EmpInfoArray.Count + 2).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        formulaRange = "R" + (totRowCtr + inCtr + 1).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        formulastring = "=SUM(T1" + ":T" + (EmpInfoArray.Count + 1).ToString + ")"
        formulaRange = "T" + (EmpInfoArray.Count + 2).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        formulaRange = "R" + (totRowCtr + inCtr + 2).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        formulastring = "=SUM(U1" + ":U" + (EmpInfoArray.Count + 1).ToString + ")"
        formulaRange = "U" + (EmpInfoArray.Count + 3).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        formulaRange = "R" + (totRowCtr + inCtr + 3).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        formulastring = "=SUM(R" + (totRowCtr + inCtr + 1).ToString + ":R" + (totRowCtr + inCtr + 3).ToString + ")"
        formulaRange = "R" + (totRowCtr + inCtr + 4).ToString
        Rng = objXLWsVar.Range(formulaRange)
        Rng.Formula = formulastring

        Dim totLbl As String() = {"TEAM TOTAL:", "FETEC Jobcode Total:", "Non-FETEC Jobcode Total:", "FETP Jobcode Total:", "GRAND TOTAL:"}
        inCtr = 3
        For Each item2 As String In totLbl
            objXLWsVar.Cells(totRowCtr + inCtr, 14) = item2
            inCtr += 1
        Next
        Rng = objXLWsVar.Range(objXLWsVar.Cells(totRowCtr, 10), objXLWsVar.Cells(totRowCtr + inCtr, 18))
        Rng.Font.Color = Color.DarkBlue
        Rng.Font.Bold = True
    End Sub

    Public Sub GenerateOutputMulti(memberlist As List(Of Object), fileSuffix() As String, headerVar As String, TeamList As List(Of Object), generateType As String, drawEraj As Boolean, tsYear As String, tsMonth As String, erajPath As String, mhPath As String)
        Me.Show()
        objXLApp = New Excel.Application
        Dim objXLWb As Excel.Workbook
        Dim objXLWs As Excel.Worksheet
        Dim objXLWb2 As Excel.Workbook
        nErr1 = 0
        nErr2 = 0
        nErr3 = 0
        jsystotH = 0
        nonjtotH = 0
        jpitotH = 0
        grandtotH = 0
        objXLApp.DisplayAlerts = False

        objXLWb = objXLApp.Workbooks.Open(xltPath)
        If System.IO.File.Exists(mhPath) Then
            objXLWb2 = objXLApp.Workbooks.Open(mhPath)
            'objXLWs = objXLWb2.Worksheets(TeamList(0).ToString)
            updateMH = True
        Else
            objXLWs = objXLWb.Worksheets("Sheet")
            fileNames = outPath + "\" + "MH Report " + Format(Now, "yyyyMMdd") + "_" + Format(Now, "HHmmss") + ".xlsx"
            objXLWb2 = objXLApp.Workbooks.Add
            objXLWb2.SaveAs(fileNames)
        End If
        'Try
        If generateType = "Multi" Then
            For i As Integer = 0 To memberlist.Count - 1
                GenerateOutput(memberlist(i), fileSuffix, generateType, headerVar, TeamList(i), objXLWb2, objXLWb)
                progressBar.Value = 90 / (memberlist.Count / (i + 1))
                lbl_Progress.Text = progressBar.Value.ToString + "%"
            Next
        Else
            GenerateOutput(memberlist(0), fileSuffix, generateType, headerVar, TeamList(0), objXLWb2, objXLWb)
        End If

        ' Delete the first worksheet only

        'If objXLWb2.Worksheets.Count > 2 Then
        '    For i As Integer = 0 To 2
        '        objXLWb2.Worksheets(1).Delete()
        '    Next
        'Else
        '    If objXLWb2.Worksheets.Count >= 2 Then
        '        objXLWb2.Worksheets(1).Delete()
        '    End If
        'End If

        If objXLWb2.Worksheets.Count >= 2 Then
            objXLWb2.Worksheets(1).Delete()
        End If

        If drawEraj = True Then
            generateEraj(objXLWb2, objXLWb, tsYear, tsMonth, erajPath)
        End If

        If noTS(0) <> "" Or jCError(0) <> "" Or noData(0) <> "" Or erajError Then
            generateErrorLog(objXLWb2, objXLWb)
        End If

        generateTotAmt(objXLWb2, tsMonth, tsYear, TeamList(0).ToString, objXLWb)

        objXLApp.DisplayAlerts = True
        objXLApp.ScreenUpdating = True
        objXLWb2.Save()
        progressBar.Value = 100
        ' Button1.Visible = True

        If errCheck = False Then
            lblMessage.Text = "MH Data Report successfully generated!"
            'MessageBox.Show("MH Data Report successfully generated!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            lblMessage.Text = "MH Data Generated with error. Check Error Log sheet."
            'MessageBox.Show("MH Data Generated with the error. Check Error Log sheet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        errCheck = False
        updateMH = False

        amtPerTeam.Clear()
        'Catch ex As Exception
        'MessageBox.Show("Generate Error.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'Finally
        objXLWb2.Close()
        Marshal.ReleaseComObject(objXLWb2)
        objXLWb2 = Nothing
        If Not objXLWb Is Nothing Then
            ' Save changes when closing workbook
            'objXLWb.Close()
            objXLWb.Close(SaveChanges:=True)
            Marshal.ReleaseComObject(objXLWb)
            objXLWb = Nothing
        End If
        objXLApp.Quit()
        Marshal.ReleaseComObject(objXLApp)
        objXLApp = Nothing
        'End Try
        lbl_Progress.Text = "100%"
        progressBar.Value = 100

        ' OK button will show after 100%
        Button1.Visible = True
    End Sub

    Public Sub generateEraj(objXLWb2 As Excel.Workbook, objXLWb As Excel.Workbook, tsYear As String, tsMonth As String, erajPath As String)

        Dim sheetname As String = "FETP ERAJ " + tsYear
        Dim found As Boolean = False
        Dim newExcelsheet As Excel.Worksheet

        Dim erajObjXLWB As Excel.Workbook = objXLApp.Workbooks.Open(erajPath)

        Try
            If updateMH = True Then
                For Each xs As Excel.Worksheet In objXLWb2.Sheets
                    If xs.Name = sheetname Then
                        found = True
                        Exit For
                    End If
                Next
                If found = True Then
                    newExcelsheet = objXLWb2.Worksheets(sheetname)
                Else
                    If System.IO.File.Exists(erajPath) Then
                        found = False
                        For Each xs As Excel.Worksheet In erajObjXLWB.Sheets
                            If xs.Name = sheetname Then
                                found = True
                                Exit For
                            End If
                        Next
                        If found = True Then
                            erajObjXLWB.Worksheets(sheetname).Copy(, before:=objXLWb2.Worksheets(1))
                        Else
                            erajObjXLWB.Worksheets("FETP ERAJ TEMP").Copy(, before:=objXLWb2.Worksheets(1))
                        End If
                        newExcelsheet = objXLWb2.Worksheets(1)
                        'newExcelsheet = objXLWb2.ActiveSheet
                        newExcelsheet.Name = sheetname
                        objXLApp.DisplayAlerts = True
                    Else
                        erajObjXLWB.Worksheets("FETP ERAJ TEMP").Copy(, before:=objXLWb2.Worksheets(1))
                    End If
                End If

            Else
                If System.IO.File.Exists(erajPath) Then
                    found = False
                    For Each xs As Excel.Worksheet In erajObjXLWB.Sheets
                        If xs.Name = sheetname Then
                            found = True
                            Exit For
                        End If
                    Next
                    If found = True Then
                        erajObjXLWB.Worksheets(sheetname).Copy(, before:=objXLWb2.Worksheets(1))
                    Else
                        ' Use sheet number instead of name. (There is no FETP ERAJ TEMP worksheet)
                        ' erajObjXLWB.Worksheets("FETP ERAJ TEMP").Copy(, before:=objXLWb2.Worksheets(1))
                        erajObjXLWB.Worksheets(1).Copy(, before:=objXLWb2.Worksheets(1))
                    End If
                    newExcelsheet = objXLWb2.Worksheets(1)
                    'newExcelsheet = objXLWb2.ActiveSheet
                    newExcelsheet.Name = sheetname
                    objXLApp.DisplayAlerts = True
                End If
            End If

            Dim destrow As Integer = 6
            Dim sourceRow As Integer = 6
            Dim teamName As String = ""
            Dim monthCol As Integer = 1
            For no As Integer = 0 To 13
                If CStr(DirectCast(newExcelsheet.Cells(4, monthCol), Excel.Range).Value) = tsMonth.ToUpper Then
                    Exit For
                Else
                    monthCol += 1
                End If
            Next

            For i As Integer = 0 To amtPerTeam.Count - 1
                sourceRow = 6
                Do While CStr(DirectCast(newExcelsheet.Cells(sourceRow, 1), Excel.Range).Value) <> "FETP TOTAL"
                    teamName = CStr(DirectCast(newExcelsheet.Cells(sourceRow, 1), Excel.Range).Value).ToUpper
                    If teamName = amtPerTeam(i)(0).ToString.ToUpper Then
                        newExcelsheet.Cells(sourceRow, monthCol) = amtPerTeam(i)(1)
                        newExcelsheet.Cells(sourceRow + 1, monthCol) = amtPerTeam(i)(2)
                        newExcelsheet.Cells(sourceRow + 2, monthCol) = amtPerTeam(i)(3)
                        Exit Do
                    End If
                    sourceRow += 5
                Loop
            Next

        Catch ex As Exception
            erajError = True
            errCheck = True
            ' Show message when error occurs
            ' MessageBox.Show(ex.StackTrace, "Error Occured", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Public Sub generateTotAmt(objXLWb2 As Excel.Workbook, tsmonth As String, tsyear As String, teamName As String, objXLWb As Excel.Workbook)
        Dim amtWs As Excel.Worksheet
        Dim Rng As Excel.Range


        If updateMH = True Then
            amtWs = objXLWb2.Worksheets("Summary")
            'amtWs.Unprotect("admin")
            Dim sourceRow As Integer = 4
            Dim teamFound As Boolean = False

            Do While CStr(DirectCast(amtWs.Cells(sourceRow, 1), Excel.Range).Value).ToString <> ""
                If CStr(DirectCast(amtWs.Cells(sourceRow, 1), Excel.Range).Value) = teamName Then
                    teamFound = True
                    Exit Do
                Else
                    sourceRow += 1
                End If
            Loop
            Dim rowArray() As Object
            For i As Integer = 0 To amtPerTeam.Count - 1
                Rng = amtWs.Range(amtWs.Cells(sourceRow, 1), amtWs.Cells(sourceRow, 5))
                rowArray = amtPerTeam(i)
                Rng.Value = rowArray
                Rng.NumberFormat = "0.00"
            Next
        Else
            amtWs = CType(objXLWb2.Worksheets.Add(Before:=objXLWb2.Worksheets(1)), Excel.Worksheet)
            'amtWs = objXLWb2.ActiveSheet
            amtWs.Name = "Summary"
            Rng = amtWs.Range(amtWs.Cells(1, 1), amtWs.Cells(1, 5))
            With Rng
                .Merge()
                .Font.Size = 14
                .Font.Bold = True
                .NumberFormat = "@"
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = True
            End With
            amtWs.Cells(1, 1) = tsmonth.ToUpper + " " + tsyear
            Rng = amtWs.Range(amtWs.Cells(2, 1), amtWs.Cells(2, 5))
            amtWs.Cells(2, 1) = "Job Code Totals per Team"
            With Rng
                .Merge()
                .Font.Size = 14
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
                .WrapText = True
            End With
            Dim header() As String = {"Team Name", "FETEC Jobcode Total", "Non-FETEC Jobcode Total", "FETP Jobcode Total", "Sub Total"}
            Dim ctr As Integer = 1
            For i As Integer = 0 To header.Count - 1
                amtWs.Cells(3, ctr) = header(i)
                ctr += 1
            Next
            Rng = amtWs.Range(amtWs.Cells(3, 1), amtWs.Cells(3, 5))
            With Rng
                .Font.Bold = True
                .Font.Size = 12
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                .EntireColumn.AutoFit()
            End With
            Dim rowArray() As Object
            Dim destrow As Integer = 4
            Dim teamTotForm As String = ""
            For i As Integer = 0 To amtPerTeam.Count - 1
                Rng = amtWs.Range(amtWs.Cells(destrow, 1), amtWs.Cells(destrow, 4))
                rowArray = amtPerTeam(i)
                Rng.Value = rowArray

                Dim rngStr As String = "E" + destrow.ToString
                Rng = amtWs.Range(rngStr)
                teamTotForm = "=sum(B" + destrow.ToString + ":D" + destrow.ToString + ")"
                Rng.Formula = teamTotForm

                destrow += 1
                Rng.NumberFormat = "0.00"
            Next


            Rng = amtWs.Range(amtWs.Cells(2, 1), amtWs.Cells(2, 5))
            With Rng
                .Font.Bold = True
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With
            Dim totLbl As String() = {"FETEC Jobcode Total:", "Non-FETEC Jobcode Total:", "FETP Jobcode Total:", "GRAND TOTAL:"}
            Dim destrowAdd = 2
            For Each item2 As String In totLbl
                amtWs.Cells(destrow + destrowAdd, 3) = item2
                destrowAdd += 1
            Next

            Rng = amtWs.Range(amtWs.Cells(destrow, 2), amtWs.Cells(destrow + destrowAdd, 5))
            With Rng
                .Font.Color = Color.DarkBlue
                .Font.Bold = True
                .NumberFormat = "0.00"
            End With

            Dim sumtotalformula As String = ""
            Dim sumtotalRange As String = ""

            Dim c As Char
            For i As Integer = 2 To 5
                c = Convert.ToChar(i + 64)
                sumtotalformula = "=SUM(" + c.ToString + "3:" + c.ToString + (amtPerTeam.Count + 3).ToString + ")"
                sumtotalRange = c.ToString + (amtPerTeam.Count + 4).ToString
                Rng = amtWs.Range(sumtotalRange)
                Rng.Formula = sumtotalformula
                Rng.Locked = True

                sumtotalRange = "E" + (amtPerTeam.Count + i + 4).ToString
                Rng = amtWs.Range(sumtotalRange)
                Rng.Formula = sumtotalformula
                Rng.Locked = True
            Next
        End If
        'amtWs.Protect("admin")
    End Sub

    Public Sub generateErrorLog(objXLWb2 As Excel.Workbook, objXLWb As Excel.Workbook)
        Dim errWs As Excel.Worksheet
        Dim Rng As Excel.Range
        Dim found As Boolean = False

        For Each xs As Excel.Worksheet In objXLWb2.Sheets
            If xs.Name = "Error Log" Then
                found = True
                Exit For
            End If
        Next
        'If found = True Then
        '    Dim sheetNo As Integer = objXLWb2.Sheets("Error Log").Index()
        '    objXLWb2.Worksheets("Error Log").Delete()
        '    objXLWb2.Save()
        '    'errWs = CType(objXLWb2.Worksheets.Add(Before:=objXLWb2.Worksheets(sheetNo)), Excel.Worksheet)
        'Else
        '    errWs = CType(objXLWb2.Worksheets.Add(Before:=objXLWb2.Worksheets(1)), Excel.Worksheet)
        'End If

        If found = True Then
            errWs = objXLWb2.Worksheets("Error Log")
            errWs.Unprotect("admin")
            objXLApp.DisplayAlerts = True
        Else
            errWs = CType(objXLWb2.Worksheets.Add(Before:=objXLWb2.Worksheets(1)), Excel.Worksheet)
            errWs.Name = "Error Log"
        End If

        errWs = objXLWb2.ActiveSheet
        errWs.Cells(1, 1) = "Timesheet not found"
        errWs.Cells(1, 2) = "Timesheet with Job Code Discrepancies"
        errWs.Cells(1, 3) = "Empty Timesheet"
        Rng = errWs.Range(errWs.Cells(1, 1), errWs.Cells(1, 3))

        With Rng
            .Font.Bold = True
            .Font.Size = 12
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            .HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            .EntireColumn.AutoFit()
        End With

        Dim destrow = 2
        Dim lastrow = 0

        For Each item In noTS
            errWs.Cells(destrow, 1) = item
            If item = "" Then
                Exit For
            End If
            destrow += 1
        Next

        lastrow = destrow
        destrow = 2

        For Each item In jCError
            errWs.Cells(destrow, 2) = item
            If item = "" Then
                Exit For
            End If
            destrow += 1
        Next

        If lastrow < destrow Then
            lastrow = destrow
        End If

        destrow = 2

        For Each item In noData
            errWs.Cells(destrow, 3) = item
            If item = "" Then
                Exit For
            End If
            destrow += 1
        Next

        If erajError Then
            errWs.Cells(lastrow + 2, 1) = "Error encountered when attempting to draw ERAJ."
        End If

        Array.Clear(noTS, 0, 200)
        Array.Clear(noData, 0, 200)
        Array.Clear(jCError, 0, 200)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub frm_Generate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Visible = False
    End Sub
End Class
