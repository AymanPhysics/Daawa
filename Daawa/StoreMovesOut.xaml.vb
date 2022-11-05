Imports System.Data
Imports System.IO
Imports System.ComponentModel

Public Class StoreMovesOut
    Public TableName As String = "StoreMovesOut"
    Public SubId As String = "Id"

    Dim dt As New DataTable
    Dim bm As New BasicMethods
    WithEvents G As New MyGrid

    Dim WithEvents BackgroundWorker1 As New BackgroundWorker
    Public LagnaId, OperationId As Integer

    Private Sub StoreMovesOut_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If bm.TestIsLoaded Then Return
        bm.Fields = New String() {SubId, "Daydate", "Notes"}
        bm.control = New Control() {txtID, DayDate, Notes}
        bm.KeyFields = New String() {SubId}

        bm.Table_Name = TableName

        LoadWFH()

        btnNew_Click(sender, e)

    End Sub


    Structure GC
        Shared Id As String = "Id"
        Shared Name As String = "Name"
        Shared Qty As String = "Qty"
        Shared Price As String = "Price"
        Shared Value As String = "Value"
    End Structure

    Private Sub LoadWFH()
        WFH.Child = G

        G.Grid.ForeColor = System.Drawing.Color.DarkBlue
        G.Grid.Columns.Add(GC.Id, "كود الصنف")
        G.Grid.Columns.Add(GC.Name, "اسم الصنف")
        G.Grid.Columns.Add(GC.Qty, "الكمية")
        G.Grid.Columns.Add(GC.Price, "السعر")
        G.Grid.Columns.Add(GC.Value, "القيمة")

        G.Grid.Columns(GC.Id).FillWeight = 100
        G.Grid.Columns(GC.Name).FillWeight = 300
        G.Grid.Columns(GC.Qty).FillWeight = 100
        G.Grid.Columns(GC.Price).FillWeight = 100
        G.Grid.Columns(GC.Value).FillWeight = 100

        G.Grid.Columns(GC.Name).ReadOnly = True
        G.Grid.Columns(GC.Value).ReadOnly = True
        G.Grid.AllowUserToDeleteRows = True

        AddHandler G.Grid.CellEndEdit, AddressOf GridCalcRow
        AddHandler G.Grid.KeyDown, AddressOf GridKeyDown
        AddHandler G.Grid.EditingControlShowing, AddressOf GridEditingControlShowing
    End Sub

    Private Sub GridCalcRow(ByVal sender As Object, ByVal e As Forms.DataGridViewCellEventArgs)
        If G.Grid.Columns(e.ColumnIndex).Name = GC.Id Then
            AddItem(G.Grid.Rows(e.RowIndex).Cells(GC.Id).Value, e.RowIndex, 0)
        End If
        G.Grid.Rows(e.RowIndex).Cells(GC.Qty).Value = Val(G.Grid.Rows(e.RowIndex).Cells(GC.Qty).Value)
        G.Grid.Rows(e.RowIndex).Cells(GC.Price).Value = Val(G.Grid.Rows(e.RowIndex).Cells(GC.Price).Value)
        G.Grid.Rows(e.RowIndex).Cells(GC.Value).Value = Val(G.Grid.Rows(e.RowIndex).Cells(GC.Qty).Value) * Val(G.Grid.Rows(e.RowIndex).Cells(GC.Price).Value)
        If Val(G.Grid.Rows(e.RowIndex).Cells(GC.Id).Value) = 0 Then
            G.Grid.Rows(e.RowIndex).Cells(GC.Name).Value = ""
            G.Grid.Rows(e.RowIndex).Cells(GC.Qty).Value = ""
            G.Grid.Rows(e.RowIndex).Cells(GC.Price).Value = ""
            G.Grid.Rows(e.RowIndex).Cells(GC.Value).Value = ""
        End If
        G.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter
    End Sub


    Sub AddItem(ByVal Id As String, Optional ByVal i As Integer = -1, Optional ByVal Add As Decimal = 1)
        Try
            Dim Exists As Boolean = False
            Dim Move As Boolean = False
            If i = -1 Then Move = True

            G.Grid.AutoSizeColumnsMode = Forms.DataGridViewAutoSizeColumnsMode.Fill
            If i = -1 Then
                For x As Integer = 0 To G.Grid.Rows.Count - 1
                    If Not G.Grid.Rows(x).Cells(GC.Id).Value Is Nothing AndAlso G.Grid.Rows(x).Cells(GC.Id).Value.ToString = Id.ToString Then
                        i = x
                        Exists = True
                        GoTo Br
                    End If
                Next
                i = G.Grid.Rows.Add()
Br:
            End If

            Dim dt As DataTable = bm.ExcuteAdapter("Select * From Items where Id='" & Id & "'")
            Dim dr() As DataRow = dt.Select("Id='" & Id & "'")
            If dr.Length = 0 Then
                If Not G.Grid.Rows(i).Cells(GC.Id).Value Is Nothing AndAlso G.Grid.Rows(i).Cells(GC.Id).Value <> "" Then bm.ShowMSG("هذا الصنف غير موجود")
                ClearRow(i)
                Return
            End If
            G.Grid.Rows(i).Cells(GC.Id).Value = dr(0)(GC.Id)
            G.Grid.Rows(i).Cells(GC.Name).Value = dr(0)(GC.Name)
            G.Grid.CurrentCell = G.Grid.Rows(i).Cells(GC.Qty)
            G.Grid.CurrentCell = G.Grid.Rows(i).Cells(GC.Name)

            If Val(G.Grid.Rows(i).Cells(GC.Qty).Value) = 0 Then G.Grid.Rows(i).Cells(GC.Qty).Value = 1

            If Move Then
                G.Grid.Focus()
                G.Grid.Rows(i).Selected = True
                G.Grid.FirstDisplayedScrollingRowIndex = i
                G.Grid.CurrentCell = G.Grid.Rows(i).Cells(GC.Qty)
                G.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter
                G.Grid.BeginEdit(True)
            End If
            If Exists Then
                G.Grid.Rows(i).Selected = True
                G.Grid.FirstDisplayedScrollingRowIndex = i
                G.Grid.CurrentCell = G.Grid.Rows(i).Cells(GC.Name)
                G.Grid.CurrentCell = G.Grid.Rows(i).Cells(GC.Qty)
                G.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter
                G.Grid.BeginEdit(True)
            End If
        Catch
            If i <> -1 Then
                ClearRow(i)
            End If
        End Try
    End Sub

    Dim lop As Boolean = False

    Sub ClearRow(ByVal i As Integer)
        G.Grid.Rows(i).Cells(GC.Id).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Name).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Qty).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Price).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Value).Value = Nothing
    End Sub

    Sub FillControls()
        bm.FillControls()


        Dim dt As DataTable = bm.ExcuteAdapter("select * from StoreMovesOutDetails1 where StoreMovesOutId=" & txtID.Text)
        G.Grid.Rows.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            G.Grid.Rows.Add()
            G.Grid.Rows(i).Cells(GC.Id).Value = dt.Rows(i)("Id").ToString
            G.Grid.Rows(i).Cells(GC.Name).Value = dt.Rows(i)("Name").ToString
            G.Grid.Rows(i).Cells(GC.Qty).Value = dt.Rows(i)("Qty").ToString
            G.Grid.Rows(i).Cells(GC.Price).Value = dt.Rows(i)("Price").ToString
            G.Grid.Rows(i).Cells(GC.Value).Value = dt.Rows(i)("Value").ToString
        Next

        Notes.Focus()
        G.Grid.RefreshEdit()
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
        bm.DefineValues()
        If Not bm.Save(New String() {SubId}, New String() {txtID.Text.Trim}) Then Return

        bm.SaveGrid(G.Grid, "StoreMovesOutDetails1", New String() {"StoreMovesOutId"}, New String() {txtID.Text}, New String() {"Id", "Name", "Qty", "Price", "Value"}, New String() {GC.Id, GC.Name, GC.Qty, GC.Price, GC.Value}, New VariantType() {VariantType.Integer, VariantType.String, VariantType.Decimal, VariantType.Decimal, VariantType.Decimal}, New String() {GC.Id})

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

        G.Grid.Rows.Clear()
        txtID.Text = bm.ExecuteScalar("select max(" & SubId & ")+1 from " & TableName & " where 1=1 " & bm.AppendWhere)
        If txtID.Text = "" Then txtID.Text = "1"
        DayDate.Focus()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من المسح؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & SubId & "='" & txtID.Text.Trim & "' ")
            bm.ExcuteNonQuery("delete from StoreMovesOutDetails1 where StoreMovesOutId='" & txtID.Text & "' ")
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

    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles txtID.KeyDown
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

    Private Sub GridKeyDown(ByVal sender As Object, ByVal e As Forms.KeyEventArgs)
        Try
            If G.Grid.CurrentCell Is G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Id) AndAlso bm.ShowHelpGrid("الأصنــــاف", G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Id), G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Name), e, "select cast(Id as varchar(100)) Id,Name from Items") Then
                G.Grid.CurrentCell = G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Qty)
                G.Grid.CurrentCell = G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Id)
            End If
        Catch
        End Try
    End Sub


    Private Sub GridEditingControlShowing(ByVal sender As Object, ByVal e As Forms.DataGridViewEditingControlShowingEventArgs)
        Dim c = e.Control
        RemoveHandler c.KeyDown, AddressOf GridKeyDown
        AddHandler c.KeyDown, AddressOf GridKeyDown
    End Sub

End Class
