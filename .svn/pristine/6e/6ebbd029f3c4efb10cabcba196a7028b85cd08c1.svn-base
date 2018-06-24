Imports System.IO
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Module ExcelUtil
    'set default value
    Dim endFile As String = "end"

    Public Function GetTeamList(directoryPath As String) As List(Of String)
        Try
            ' variable declaration
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim teamList As New List(Of String)
            Dim rowStart As Integer = 1
            Dim colStart As Integer = 1

            ' open excel file
            xlWorkBook = xlApp.Workbooks.Open(directoryPath)
            xlWorkSheet = xlWorkBook.Worksheets(1)

            ' get first item (team name) from excel file
            Dim teamName As String = xlWorkSheet.Cells(rowStart, colStart).value

            ' loop all items (team names) and store in a list
            While Not teamName.Contains(endFile)
                ' store the item in the list
                teamList.Add(teamName)
                ' increment the column for the next item
                colStart = colStart + 1
                'get the next item
                teamName = xlWorkSheet.Cells(rowStart, colStart).value()
            End While

            ' close excel file
            xlWorkBook.Close()
            xlWorkBook = Nothing
            xlApp.Quit()
            xlApp = Nothing

            ' return the list
            Return teamList
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetMemberList(team As String, directoryPath As String) As List(Of String)
        Try
            ' variable declaration
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim teamList As New List(Of String)
            Dim memberList As New List(Of String)
            Dim colStart As Integer
            Dim rowStart As Integer = 2

            ' open excel file
            xlWorkBook = xlApp.Workbooks.Open(directoryPath)
            xlWorkSheet = xlWorkBook.Worksheets(1)

            ' call function to get team name
            teamList = GetTeamList(directoryPath)

            ' get the column index of the desired team
            colStart = teamList.IndexOf(team) + 1

            ' get first item (member's name) of the desired team from excel file
            Dim memberName As String = xlWorkSheet.Cells(rowStart, colStart).value

            ' loop all items (member's name) and store in a list
            While Not memberName.Contains(endFile)
                ' store the item in the list
                memberList.Add(memberName)
                ' increment the row for the next item
                rowStart = rowStart + 1
                'get the next item 
                memberName = xlWorkSheet.Cells(rowStart, colStart).value()
            End While

            ' close excel file
            xlWorkBook.Close()
            xlWorkBook = Nothing
            xlApp.Quit()
            xlApp = Nothing

            ' return the list
            Return memberList
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function GetAllMemberList(directoryPath As String) As List(Of String)
        Try
            ' variable declaration
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim allMemberList As New List(Of String)
            Dim rowTeam As Integer = 1
            Dim colTeam As Integer = 1
            Dim rowName As Integer
            xlApp.DisplayAlerts = False

            ' open excel file
            xlWorkBook = xlApp.Workbooks.Open(directoryPath)
            xlWorkSheet = xlWorkBook.Worksheets(1)

            ' get first item (team name) from excel file
            Dim teamName As String = xlWorkSheet.Cells(rowTeam, colTeam).value

            ' loop all items (team names)
            While Not teamName.Contains(endFile)
                ' get first item (member's name) from excel file
                rowName = 2
                Dim memberName As String = xlWorkSheet.Cells(rowName, colTeam).value

                ' loop all items (member's names) per team and store in a list
                While Not memberName.Contains(endFile)
                    ' store the item in the list
                    allMemberList.Add(memberName)
                    ' increment the row for the next item
                    rowName = rowName + 1
                    'get the next item 
                    memberName = xlWorkSheet.Cells(rowName, colTeam).value()
                End While
                ' increment the column for the next team
                colTeam = colTeam + 1
                'get the next team 
                teamName = xlWorkSheet.Cells(rowTeam, colTeam).value
            End While

            ' close excel file
            xlApp.DisplayAlerts = True
            xlWorkBook.Close()
            xlWorkBook = Nothing
            xlApp.Quit()
            xlApp = Nothing

            ' return the list
            Return allMemberList
        Catch ex As Exception
            Throw
        End Try
    End Function
End Module
