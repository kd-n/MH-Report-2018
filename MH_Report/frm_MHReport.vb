Public Class frm_MHReport

    ' constant declaration
    Private Const MM_FORMAT As String = "00"
    Private Const PERIOD1 As String = "B_"
    Private Const PERIOD2 As String = "E_"
    Private Const PROCESS_MSG As String = "MH Data Report generation completed successfully."
    Dim erajPath As String = ""
    Dim mhPath As String = ""



    Private Sub frm_MHReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btn_Generate.Enabled = True
        btn_Close.Enabled = True
        txtbox_Eraj.Enabled = False
        txtbox_Eraj.Clear()
        btn_UploadEraj.Enabled = False
        btn_UploadMH.Enabled = False
        txt_MHUpdate.Enabled = False
        txt_MHUpdate.Clear()
        chk_UpdateMHReport.Hide()
        txt_MHUpdate.Hide()
        btn_UploadMH.Hide()

        Dim i As Integer
        ' populate year combo box
        With cmb_Year
            For i = -2 To 1
                .Items.Add(Year(Now) + i)
            Next

            ' set default value
            .Text = Year(Now)
        End With

        ' populate month combo box
        With cmb_Month
            .Items.Add("Jan")
            .Items.Add("Feb")
            .Items.Add("Mar")
            .Items.Add("Apr")
            .Items.Add("May")
            .Items.Add("Jun")
            .Items.Add("Jul")
            .Items.Add("Aug")
            .Items.Add("Sept")
            .Items.Add("Oct")
            .Items.Add("Nov")
            .Items.Add("Dec")
        End With

        ' set default values
        If DatePart(DateInterval.Day, Now) < 15 Then
            'period
            rb_FirstPeriod.Checked = False
            rb_SecondPeriod.Checked = False
            rb_Both.Checked = True

            'month
            cmb_Month.SelectedIndex = Month(Now) - 2
        Else
            'period
            rb_FirstPeriod.Checked = True
            rb_SecondPeriod.Checked = False
            rb_Both.Checked = False

            'month
            cmb_Month.SelectedIndex = Month(Now) - 1
        End If

        ' populate team/dept combo box
        With cmb_Team
            Dim teamList As New List(Of String)
            teamList = TeamResource.GetTeamList()
            ' add "All Teams" 
            .Items.Add("All Teams")
            For Each teamName As String In teamList
                .Items.Add(teamName)
            Next
            'set default value
            .SelectedIndex = cmb_Team.Items.IndexOf("All Teams")
        End With
    End Sub

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.Close()
    End Sub

    Private Sub btn_Generate_Click(sender As Object, e As EventArgs) Handles btn_Generate.Click
        btn_Generate.Enabled = False
        btn_Close.Enabled = False

        Me.Cursor = Cursors.WaitCursor
        Dim filePrefix As String = ""
        Dim fileName2 As String = ""
        Dim fileName As String = ""

        Dim memberList As New List(Of String)
        Dim memberListAll As New List(Of Object)
        Dim teamList As New List(Of Object)
        Dim strTeamList(1) As Object
        Dim teamName As String = cmb_Team.SelectedItem
        Dim tsYear As String = cmb_Year.SelectedItem
        Dim tsMonth As String = cmb_Month.SelectedItem
        Dim i As Integer = 0
        Dim periodHeader As String = ""
        Dim drawEraj As Boolean = False
        If chk_DrawGraph.Checked Then
            If txtbox_Eraj.Text = "" Then
                ' Use blank template when no file is uploaded
                erajPath = My.Settings.ErajPath
                If erajPath = "" Then
                    ' Error when there is no default template uploaded in ERAJ settings
                    MessageBox.Show("Please specify ERAJ file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    btn_Generate.Enabled = True
                    btn_Close.Enabled = True
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If

                drawEraj = True
            Else
                drawEraj = True
            End If
        Else
            drawEraj = False
        End If

        Dim updateMH As Boolean = False
        If chk_UpdateMHReport.Checked Then
            If txt_MHUpdate.Text = "" Then
                MessageBox.Show("Please specify MH Report file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                btn_Generate.Enabled = True
                btn_Close.Enabled = True
                Me.Cursor = Cursors.Default
                Exit Sub
            Else
                updateMH = True
            End If
        Else
            updateMH = False
        End If

        If teamName.Equals("All Teams") Then
            For Each teamNames As String In TeamResource.GetTeamList()
                memberListAll.Add(TeamResource.GetMemberList(teamNames))
                teamList.Add(teamNames)
            Next
        Else
            memberListAll.Add(TeamResource.GetMemberList(teamName))
            teamList.Add(teamName)
        End If

        ' set year
        filePrefix = cmb_Year.Text.Substring(2)

        ' set month
        filePrefix = filePrefix & Format(cmb_Month.SelectedIndex + 1, MM_FORMAT)

        ' set period
        Dim fileSuffix(0) As String
        If rb_FirstPeriod.Checked Then
            fileName = filePrefix & PERIOD1
            fileSuffix = {fileName}
            periodHeader = "1st Period"
        ElseIf rb_SecondPeriod.Checked Then
            fileName = filePrefix & PERIOD2
            periodHeader = "2nd Period"
            fileSuffix = {fileName}
        Else
            ReDim fileSuffix(1)
            fileName = filePrefix & PERIOD1
            fileName2 = filePrefix & PERIOD2
            fileSuffix = {fileName, fileName2}
        End If
        Dim header As String = "MH Data for " + cmb_Month.SelectedItem.ToString + " " + cmb_Year.Text.ToString + " " + periodHeader
        If teamName.Equals("All Teams") Then
            frm_Generate.GenerateOutputMulti(memberListAll, fileSuffix, header, teamList, "Multi", drawEraj, tsYear, tsMonth, erajPath, mhPath)
        Else
            frm_Generate.GenerateOutputMulti(memberListAll, fileSuffix, header, teamList, "Uni", drawEraj, tsYear, tsMonth, erajPath, mhPath)
        End If
        Me.Cursor = Cursors.Default
        btn_Generate.Enabled = True
        btn_Close.Enabled = True
    End Sub


    Private Sub btn_Output_Click(sender As Object, e As EventArgs) Handles btn_UploadEraj.Click
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
                txtbox_Eraj.Text = erajPath
            End If

            If Not openFile.FileName.Contains(".xls") AndAlso Not erajPath.Contains(".xlsx") Then
                ' display error message if not an excel file
                MessageBox.Show("Incorrect file type.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' save the directory path if an excel file
                erajPath = openFile.FileName
                txtbox_Eraj.Text = erajPath
            End If
        End If
    End Sub

    Private Sub chk_DrawGraph_CheckedChanged(sender As Object, e As EventArgs) Handles chk_DrawGraph.CheckedChanged
        If chk_DrawGraph.Checked = True Then
            txtbox_Eraj.Enabled = True
            btn_UploadEraj.Enabled = True
        Else
            txtbox_Eraj.Enabled = False
            btn_UploadEraj.Enabled = False
        End If
    End Sub

    Private Sub btn_UploadMH_Click(sender As Object, e As EventArgs) Handles btn_UploadMH.Click
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
                mhPath = openFile.FileName
                txt_MHUpdate.Text = mhPath
            End If

            If Not openFile.FileName.Contains(".xls") AndAlso Not erajPath.Contains(".xlsx") Then
                ' display error message if not an excel file
                MessageBox.Show("Incorrect file type.  Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' save the directory path if an excel file
                mhPath = openFile.FileName
                txt_MHUpdate.Text = mhPath
            End If
        End If
    End Sub


    Private Sub chk_UpdateMHReport_CheckedChanged(sender As Object, e As EventArgs) Handles chk_UpdateMHReport.CheckedChanged
        If chk_UpdateMHReport.Checked = True Then
            btn_UploadMH.Enabled = True
            txt_MHUpdate.Enabled = True
        Else
            btn_UploadMH.Enabled = False
            txt_MHUpdate.Enabled = False
        End If
    End Sub

    Private Sub cmb_Team_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmb_Team.SelectedValueChanged
        If cmb_Month.SelectedItem = "All Teams" Then
            chk_UpdateMHReport.Enabled = False
        Else
            chk_UpdateMHReport.Enabled = True
        End If
    End Sub
End Class