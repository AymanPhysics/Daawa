Imports System.Data

Public Class RPT3
    Dim bm As New BasicMethods
    Public Flag As Integer = 0
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Dim rpt As New ReportViewer
        rpt.paraname = {"Header", "@SeasonId", "@CityId", "@AreaId", "@TownId", "@SubTownId", "@CaseTypeId", "@FromDate", "@ToDate", "CaseTypeName", "@CaseLevelId", "CaseLevelId", "@IncomeTypeId", "@CaseId", "@ProblemId", "@IllId"}
        rpt.paravalue = {CType(CType(Parent, TabItem).Header, TabsHeader).MyTabHeader, Val(SeasonId.Text), Val(CityId.Text), Val(AreaId.Text), Val(TownId.Text), Val(SubTownId.Text), Val(CaseTypeId.Text), FromDate.SelectedDate, ToDate.SelectedDate, CaseTypeName.Text, CaseLevelId.SelectedValue.ToString, CaseLevelId.Text, 0, 0, ProblemTypeId.SelectedValue.ToString(), IllTypeId.SelectedValue.ToString()}
        rpt.Header = Md.MyProject.ToString
        Select Case Flag
            Case 1
                rpt.RptPath = "CaseNeeds8.rpt"
        End Select
        rpt.ShowDialog()
    End Sub

    Private Sub UserControl_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If bm.TestIsLoaded Then Return

        FromDate.SelectedDate = Now
        ToDate.SelectedDate = Now
        bm.FillCombo("IllTypes", IllTypeId, "")
        bm.FillCombo("CaseLevels", CaseLevelId, "")
        bm.FillCombo("ProblemTypes", ProblemTypeId, "")

        SeasonId.Text = bm.ExecuteScalar("select CurrentSeason from Statics")
        SeasonId_LostFocus(Nothing, Nothing)

    End Sub

    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles CityId.KeyDown, AreaId.KeyDown, TownId.KeyDown, SeasonId.KeyDown
        bm.MyKeyPress(sender, e)
    End Sub


    Private Sub SeasonId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles SeasonId.LostFocus
        If SeasonId.Visibility = Visibility.Visible Then
            bm.LostFocus(SeasonId, SeasonName, "select Name from Seasons where Id=" & SeasonId.Text.Trim())
        End If
    End Sub

    Private Sub SeasonId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles SeasonId.KeyUp
        If bm.ShowHelp("المواسم", SeasonId, SeasonName, e, "select cast(Id as varchar(100)) Id,Name from Seasons") Then
        End If
    End Sub

    Private Sub CityId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles CityId.KeyUp
        bm.ShowHelp("المحافظات", CityId, CityName, e, "select cast(Id as varchar(100)) Id,Name from Cities ")
    End Sub

    Private Sub AreaId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles AreaId.KeyUp
        bm.ShowHelp("المراكز", AreaId, AreaName, e, "select cast(Id as varchar(100)) Id,Name from Areas where CityId=" & CityId.Text.Trim)
    End Sub
    Private Sub TownId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles TownId.KeyUp
        bm.ShowHelp("القرى", TownId, TownName, e, "select cast(Id as varchar(100)) Id,Name from Towns where CityId=" & CityId.Text.Trim & " and AreaId=" & AreaId.Text)
    End Sub

    Private Sub SubTownId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles SubTownId.KeyUp
        bm.ShowHelp("النجوع", TownId, TownName, e, "select cast(Id as varchar(100)) Id,Name from SubTowns where CityId=" & CityId.Text.Trim & " and AreaId=" & AreaId.Text & " and TownId=" & TownId.Text)
    End Sub

    Private Sub CityId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CityId.LostFocus
        bm.LostFocus(CityId, CityName, "select Name from Cities where Id=" & CityId.Text.Trim())
    End Sub
    Private Sub AreaId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles AreaId.LostFocus
        bm.LostFocus(AreaId, AreaName, "select Name from Areas where CityId=" & CityId.Text.Trim() & " and Id=" & AreaId.Text.Trim)
    End Sub
    Private Sub TownId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles TownId.LostFocus
        bm.LostFocus(TownId, TownName, "select Name from Towns where CityId=" & CityId.Text.Trim() & " and AreaId=" & AreaId.Text.Trim() & " and Id=" & TownId.Text.Trim)
    End Sub

    Private Sub SubTownId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles SubTownId.LostFocus
        bm.LostFocus(SubTownId, SubTownName, "select Name from SubTowns where CityId=" & CityId.Text.Trim() & " and AreaId=" & AreaId.Text.Trim() & " and TownId=" & TownId.Text.Trim & " and Id=" & SubTownId.Text.Trim)
    End Sub

    Private Sub CaseTypeId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles CaseTypeId.KeyUp
        bm.ShowHelp("أنواع الحالات", CaseTypeId, CaseTypeName, e, "select cast(Id as varchar(100)) Id,Name from CaseTypes")
    End Sub

    Private Sub CaseTypeId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CaseTypeId.LostFocus
        bm.LostFocus(CaseTypeId, CaseTypeName, "select Name from CaseTypes where Id=" & CaseTypeId.Text.Trim())
    End Sub


End Class
