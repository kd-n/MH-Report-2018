Imports System.Resources

Module TeamResource
    Public Const TEAM_RES As String = ".\teamGroup.resx"
    Private Const TITLE As String = "Title"
    Private Const TITLE_TXT As String = "Team Resource"
    Private Const DESCRIPTION As String = "Description"
    Private Const DESC_TXT As String = "List of FETP Teams, Departments, Divisions, and their members"
    Private Const TEAMS As String = "Teams"
    Private Const MEMBERS As String = "Members"

    Public Function UpdateTeamList(directoryPath As String) As Integer
        Dim output As Integer
        Try
            Dim teamList As List(Of String) ' list of team names
            Dim membersList As List(Of String) ' list for member's names
            Dim allMemberList As List(Of String) ' list for all member's names
            Dim teamGroup As List(Of Team) = New List(Of Team) ' list for team names and members

            ' get list of all team names
            teamList = ExcelUtil.GetTeamList(directoryPath)
            ' get list of all member's names
            allMemberList = ExcelUtil.GetAllMemberList(directoryPath)

            Using teamRes As New ResXResourceWriter(TEAM_RES)
                ' title
                teamRes.AddResource(TITLE, TITLE_TXT)
                ' description
                teamRes.AddResource(DESCRIPTION, DESC_TXT)
                ' stores all team's names
                teamRes.AddResource(TEAMS, teamList)
                ' store all member's names
                teamRes.AddResource(MEMBERS, allMemberList)
                ' loop all the team names and get all member's names per team
                For Each teamName As String In teamList
                    ' get list of all member's names per team
                    membersList = ExcelUtil.GetMemberList(teamName, directoryPath)
                    ' store the team name and member's names  
                    teamRes.AddResource(teamName, membersList)
                Next
                ' close resource file
                teamRes.Close()
            End Using

            Return output
        Catch ex As Exception
            ' return -1 if there is an error
            output = -1
            Return output
        End Try
    End Function

    Public Function GetTeamList() As List(Of String)
        Dim teamNames As List(Of String)
        ' get list of all teams
        Using teamResSet As New ResXResourceSet(TEAM_RES)
            teamNames = CType(teamResSet.GetObject(TEAMS), List(Of String))
            teamResSet.Close()
        End Using
        GetTeamList = teamNames
    End Function

    Public Function GetAllMemberList() As List(Of String)
        Dim allMemberNames As List(Of String)
        ' get list of all members
        Using teamResSet As New ResXResourceSet(TEAM_RES)
            allMemberNames = CType(teamResSet.GetObject(MEMBERS), List(Of String))
            teamResSet.Close()
        End Using
        GetAllMemberList = allMemberNames
    End Function

    Public Function GetMemberList(teamName As String) As List(Of String)
        Dim memberNames As List(Of String)
        ' get list of members per team
        Using teamResSet As New ResXResourceSet(TEAM_RES)
            memberNames = CType(teamResSet.GetObject(teamName), List(Of String))
            teamResSet.Close()
        End Using
        GetMemberList = memberNames
    End Function
End Module
