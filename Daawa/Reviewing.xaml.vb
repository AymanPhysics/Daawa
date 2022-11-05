Imports System.Data
Imports System.IO
Imports System.ComponentModel

Public Class Reviewing
    Public TableName As String = "Reviewing"
    Public SubId As String = "Id"

    Dim dt As New DataTable
    Dim bm As New BasicMethods
    WithEvents G1 As New MyGrid

    Dim WithEvents BackgroundWorker1 As New BackgroundWorker
    Public LagnaId, OperationId As Integer

    Private Sub Reviewing_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If bm.TestIsLoaded Then Return
        bm.Fields = New String() {SubId, "BookId", "ReviewerId", "Daydate", "Notes"}
        bm.control = New Control() {txtID, BookId, ReviewerId, DayDate, Notes}
        bm.KeyFields = New String() {SubId}

        bm.Table_Name = TableName
         
        LoadWFH1()

        btnNew_Click(sender, e)

    End Sub

     

    Structure GC1
        Shared Id As String = "Id"
        Shared Name As String = "Name"
    End Structure

    Private Sub LoadWFH1()
        WFH1.Child = G1
        G1.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        G1.Grid.ForeColor = System.Drawing.Color.DarkBlue
        G1.Grid.Columns.Add(GC1.Id, "مسلسل")
        G1.Grid.Columns.Add(GC1.Name, "الأخطــــــــــــــــــــاء")

        G1.Grid.Columns(GC1.Id).FillWeight = 100
        G1.Grid.Columns(GC1.Name).FillWeight = 900

        G1.Grid.Columns(GC1.Id).ReadOnly = True
        G1.Grid.AllowUserToDeleteRows = True
        G1.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        AddHandler G1.Grid.CellEndEdit, AddressOf G1_CellEndEdit
        AddHandler G1.Grid.UserDeletedRow, AddressOf G1_UserDeletedRow
    End Sub

    Dim lop As Boolean = False
    Sub FillControls()
        bm.FillControls()
        BookId_LostFocus(Nothing, Nothing)
        ReviewerId_LostFocus(Nothing, Nothing)

        Dim dt As DataTable = bm.ExcuteAdapter("select * from ReviewingDetails1 where ReviewingId=" & txtID.Text)

        G1.Grid.Rows.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            G1.Grid.Rows.Add()
            G1.Grid.Rows(i).Cells(GC1.Id).Value = dt.Rows(i)("Id").ToString
            G1.Grid.Rows(i).Cells(GC1.Name).Value = dt.Rows(i)("Name").ToString
        Next

        Notes.Focus()
        G1.Grid.RefreshEdit()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CType(Application.Current.MainWindow, MainWindow).TabControl1.Items.Remove(Me.Parent)
    End Sub

    Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
        bm.FirstLast(New String() {SubId}, "Max", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        bm.NextPrevious(New String() {SubId}, New String() {txtID.Text}, "Next", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If BookId.Text.Trim = "" Then
            bm.ShowMSG("برجاء تحديد الكتاب")
            BookId.Focus()
            Return
        End If
        If ReviewerId.Text.Trim = "" Then
            bm.ShowMSG("برجاء تحديد المراجع")
            ReviewerId.Focus()
            Return
        End If

        bm.DefineValues()
        If Not bm.Save(New String() {SubId}, New String() {txtID.Text.Trim}) Then Return

        bm.SaveGrid(G1.Grid, "ReviewingDetails1", New String() {"ReviewingId"}, New String() {txtID.Text}, New String() {"Id", "Name"}, New String() {GC1.Id, GC1.Name}, New VariantType() {VariantType.Integer, VariantType.String}, New String() {GC1.Id})

        btnNew_Click(sender, e)
        AllowClose = True
    End Sub

    Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click

        bm.FirstLast(New String() {SubId}, "Min", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        bm.ClearControls()
        ClearControls()
    End Sub

    Sub ClearControls()
        bm.ClearControls()
        BookName.Clear()
        ReviewerName.Clear()

        G1.Grid.Rows.Clear()
        txtID.Text = bm.ExecuteScalar("select max(" & SubId & ")+1 from " & TableName & " where 1=1 " & bm.AppendWhere)
        If txtID.Text = "" Then txtID.Text = "1"
        BookId.Focus()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من المسح؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & SubId & "='" & txtID.Text.Trim & "' " & bm.AppendWhere)
            bm.ExcuteNonQuery("delete from ReviewingDetails1 where ReviewingId='" & txtID.Text & "' ")
            btnNew_Click(sender, e)
        End If
    End Sub

    Private Sub btnPrevios_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevios.Click
        bm.NextPrevious(New String() {SubId}, New String() {txtID.Text}, "Back", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub
    Dim lv As Boolean = False

    Private Sub txtID_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.LostFocus
        If lv Then
            Return
        End If
        lv = True

        bm.DefineValues()
        Dim dt As New DataTable
        bm.RetrieveAll(New String() {SubId}, New String() {txtID.Text.Trim}, dt)
        If dt.Rows.Count = 0 Then
            ClearControls()
            DayDate.Focus()
            lv = False
            Return
        End If
        FillControls()
        lv = False
        'txtName.Text = dt(0)("Name")
    End Sub

    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles txtID.KeyDown, BookId.KeyDown, ReviewerId.KeyDown
        bm.MyKeyPress(sender, e)
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

    Private Sub ReviewerId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles ReviewerId.KeyUp
        bm.ShowHelp("المراجعين", ReviewerId, ReviewerName, e, "select cast(Id as varchar(100)) Id,Name from Reviewers ")
    End Sub

    Private Sub BookId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles BookId.KeyUp
        bm.ShowHelp("الكتب", BookId, BookName, e, "select cast(Id as varchar(100)) Id,Name from Items where IsBook=1")
    End Sub

    Private Sub ReviewerId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ReviewerId.LostFocus
        bm.LostFocus(ReviewerId, ReviewerName, "select Name from Reviewers where Id=" & ReviewerId.Text.Trim())
    End Sub

    Private Sub BookId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BookId.LostFocus
        bm.LostFocus(BookId, BookName, "select Name from Items where IsBook=1 and Id=" & BookId.Text.Trim())
    End Sub

    Private Sub G1_CellEndEdit(ByVal sender As Object, ByVal e As Forms.DataGridViewCellEventArgs)
        G1.Grid.Rows(G1.Grid.CurrentRow.Index).Cells(0).Value = G1.Grid.CurrentRow.Index + 1
    End Sub

    Private Sub G1_UserDeletedRow(ByVal sender As Object, ByVal e As Forms.DataGridViewRowEventArgs)
        For i As Integer = 0 To G1.Grid.Rows.Count - 1
            G1.Grid.Rows(i).Cells(0).Value = i + 1
        Next
    End Sub

End Class
