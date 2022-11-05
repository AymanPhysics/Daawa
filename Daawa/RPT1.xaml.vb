Imports System.Data

Public Class RPT1
    Dim bm As New BasicMethods
    Public Flag As Integer = 0
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click 
        Dim rpt As New ReportViewer
        rpt.paraname = {"Header", "@SeasonId", "@CityId", "@AreaId", "@TownId", "@SubTownId", "@CaseTypeId", "@NeedId", "@NeedPeriod", "@FromDate", "@ToDate", "CaseTypeName", "@CaseLevelId", "CaseTypeName", "NeedPeriod", "CaseLevelId", "@IncomeTypeId", "ViewSubReport", "@CaseId"}
        rpt.paravalue = {CType(CType(Parent, TabItem).Header, TabsHeader).MyTabHeader, Val(SeasonId.Text), Val(CityId.Text), Val(AreaId.Text), Val(TownId.Text), Val(SubTownId.Text), Val(CaseTypeId.Text), Val(NeedId.Text), NeedPeriod.SelectedValue.ToString, FromDate.SelectedDate, ToDate.SelectedDate, CaseTypeName.Text, CaseLevelId.SelectedValue.ToString, CaseTypeName.Text, NeedPeriod.Text, CaseLevelId.Text, 0, IIf(ViewSubReport.IsChecked, 1, 0), 0}
        rpt.Header = Md.MyProject.ToString
        Select Case Flag
            Case 1
                rpt.RptPath = "Cases.rpt"
            Case 2
                rpt.RptPath = "CaseNeeds.rpt"
            Case 3
                rpt.RptPath = "CaseNeeds2.rpt"
            Case 4
                rpt.RptPath = "CaseNeeds3.rpt"
            Case 5
                rpt.RptPath = "CaseNeeds5.rpt"
            Case 6
                rpt.RptPath = "CaseNeeds6.rpt"
            Case 7
                rpt.RptPath = "CaseNeeds7.rpt"
        End Select
        rpt.ShowDialog()
    End Sub

    Private Sub UserControl_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If bm.TestIsLoaded Then Return
        Select Flag
            Case 1
                lblNeedId.Visibility = Windows.Visibility.Hidden
                NeedId.Visibility = Windows.Visibility.Hidden
                NeedName.Visibility = Windows.Visibility.Hidden
                lblNeedPeriod.Visibility = Windows.Visibility.Hidden
                NeedPeriod.Visibility = Windows.Visibility.Hidden
                lblCaseLevelId.Visibility = Windows.Visibility.Hidden
                CaseLevelId.Visibility = Windows.Visibility.Hidden
            Case 4
                ViewSubReport.Visibility = Windows.Visibility.Hidden
        End Select

        FromDate.SelectedDate = Now
        ToDate.SelectedDate = Now
        bm.FillCombo("NeedPeriod", NeedPeriod, "")
        bm.FillCombo("CaseLevels", CaseLevelId, "")

        SeasonId.Text = bm.ExecuteScalar("select CurrentSeason from Statics")
        SeasonId_LostFocus(Nothing, Nothing)

    End Sub

    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles CityId.KeyDown, AreaId.KeyDown, TownId.KeyDown, SeasonId.KeyDown, NeedId.KeyDown, NeedPeriod.KeyDown
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


    Private Sub NeedId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles NeedId.KeyUp
        bm.ShowHelp("الاحتياجات", NeedId, NeedName, e, "select cast(Id as varchar(100)) Id,Name from Needs ")
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

    Private Sub NeedId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles NeedId.LostFocus
        bm.LostFocus(NeedId, NeedName, "select Name from Needs where Id=" & NeedId.Text.Trim())
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
