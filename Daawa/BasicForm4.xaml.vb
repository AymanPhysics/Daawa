Imports System.Data

Public Class BasicForm4
    Public MainTableName As String = "Cities"
    Public MainSubId As String = "Id"
    Public MainSubName As String = "Name"

    Public Main2TableName As String = "Areas"
    Public Main2MainId As String = "CityId"
    Public Main2SubId As String = "Id"
    Public Main2SubName As String = "Name"

    Public Main3TableName As String = "Towns"
    Public Main3MainId As String = "CityId"
    Public Main3MainId2 As String = "AreaId"
    Public Main3SubId As String = "Id"
    Public Main3SubName As String = "Name"

    Public TableName As String = "SubTowns"
    Public MainId As String = "CityId"
    Public MainId2 As String = "AreaId"
    Public MainId3 As String = "TownId"
    Public SubId As String = "Id"
    Public SubName As String = "Name"


    Dim dt As New DataTable
    Dim bm As New BasicMethods

    Public Flag As Integer = 0

    Private Sub BasicForm4_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If bm.TestIsLoaded Then Return
        bm.FillCombo(MainTableName, CboMain, "")
        bm.Fields = New String() {MainId, MainId2, MainId3, SubId, SubName}
        bm.control = New Control() {CboMain, CboMain2, CboMain3, txtID, txtName}
        bm.KeyFields = New String() {MainId, MainId2, MainId3, SubId}

        bm.Table_Name = TableName
        btnNew_Click(sender, e)
        
    End Sub

    Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
        bm.FirstLast(New String() {MainId, MainId2, MainId3, SubId}, "Max", dt)
        If dt.Rows.Count = 0 Then Return
        bm.FillControls()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        bm.NextPrevious(New String() {MainId, MainId2, MainId3, SubId}, New String() {CboMain.SelectedValue.ToString, CboMain2.SelectedValue.ToString, CboMain3.SelectedValue.ToString, txtID.Text}, "Next", dt)
        If dt.Rows.Count = 0 Then Return
        bm.FillControls()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtName.Text.Trim = "" Or CboMain.SelectedValue.ToString = 0 Or CboMain2.SelectedValue.ToString = 0 Or CboMain3.SelectedValue.ToString = 0 Then
            Return
        End If
        DelivaryCost.Text = Val(DelivaryCost.Text)
        bm.DefineValues()
        If Not bm.Save(New String() {MainId, MainId2, MainId3, SubId}, New String() {CboMain.SelectedValue.ToString, CboMain2.SelectedValue.ToString, CboMain3.SelectedValue.ToString, txtID.Text.Trim}) Then Return
        btnNew_Click(sender, e)
        AllowClose = True
    End Sub

    Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click

        bm.FirstLast(New String() {MainId, MainId2, MainId3, SubId}, "Min", dt)
        If dt.Rows.Count = 0 Then Return
        bm.FillControls()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        'bm.ClearControls()
        ClearControls()
    End Sub

    Sub ClearControls()
        Try
            txtName.Clear()
            DelivaryCost.Clear()
            txtID.Text = bm.ExecuteScalar("select max(" & SubId & ")+1 from " & TableName & " where " & MainId & "='" & CboMain.SelectedValue.ToString & "' and " & MainId2 & "='" & CboMain2.SelectedValue.ToString & "' and " & MainId3 & "='" & CboMain3.SelectedValue.ToString & "'")
            If txtID.Text = "" Then txtID.Text = "1"

            txtName.Focus()
        Catch
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من المسح؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & SubId & "='" & txtID.Text.Trim & "' and " & MainId & " ='" & CboMain.SelectedValue.ToString & "' and " & MainId2 & " ='" & CboMain2.SelectedValue.ToString & "' and " & MainId3 & " ='" & CboMain3.SelectedValue.ToString & "'")
            btnNew_Click(sender, e)
        End If
    End Sub

    Private Sub btnPrevios_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevios.Click
        bm.NextPrevious(New String() {MainId, MainId2, SubId}, New String() {CboMain.SelectedValue.ToString, CboMain2.SelectedValue.ToString, CboMain3.SelectedValue.ToString, txtID.Text}, "Back", dt)
        If dt.Rows.Count = 0 Then Return
        bm.FillControls()
    End Sub
    Dim lv As Boolean = False
    Private Sub txtID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.LostFocus
        If lv Then
            Return
        End If
        lv = True

        bm.DefineValues()
        Dim dt As New DataTable
        bm.RetrieveAll(New String() {MainId, MainId2, MainId3, SubId}, New String() {CboMain.SelectedValue.ToString, CboMain2.SelectedValue.ToString, CboMain3.SelectedValue.ToString, txtID.Text.Trim}, dt)
        If dt.Rows.Count = 0 Then
            ClearControls()
            lv = False
            Return
        End If
        bm.FillControls()
        lv = False
    End Sub

    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles txtID.KeyDown
        bm.MyKeyPress(sender, e)
    End Sub

    Private Sub DelivaryCost_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles DelivaryCost.KeyDown
        bm.MyKeyPress(sender, e, True)
    End Sub

    Private Sub CboMain_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboMain.SelectionChanged
        Dim s As String = ""
        Try
            s = CboMain.SelectedValue.ToString
        Catch ex As Exception
        End Try
        bm.FillCombo(Main2TableName, CboMain2, " where " & Main2MainId & "='" & s & "'")
        ClearControls()
    End Sub

    Private Sub CboMain2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboMain2.SelectionChanged
        Dim s As String = ""
        Try
            s = CboMain.SelectedValue.ToString
        Catch ex As Exception
        End Try
        Dim s2 As String = ""
        Try
            s2 = CboMain2.SelectedValue.ToString
        Catch ex As Exception
        End Try
        bm.FillCombo(Main3TableName, CboMain3, " where " & Main2MainId & "='" & s & "' and " & Main3MainId2 & "='" & s2 & "'")
        ClearControls()
    End Sub

    Private Sub CboMain3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CboMain3.SelectionChanged
        ClearControls()
    End Sub

    Dim AllowClose As Boolean = False
    'Private Sub MyBase_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
    '    If Not btnSave.Enabled Then Exit Sub
    '    Select Case bm.RequestDelete
    '        Case BasicMethods.CloseState.Yes
    '            AllowClose = False
    '            btnSave_Click(Nothing, Nothing)
    '            If Not AllowClose Then e.Cancel = True
    '        Case BasicMethods.CloseState.No

    '        Case BasicMethods.CloseState.Cancel
    '            e.Cancel = True
    '    End Select
    'End Sub

End Class
