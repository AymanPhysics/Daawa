Imports System.Data

Public Class Help
    Dim bm As New BasicMethods
    Public FirstColumn As String = "الكود", SecondColumn As String = "الاسم", Statement As String = "", TableName As String = ""

    Dim dt As New DataTable
    Dim dv As New DataView
    Public Header As String = ""

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Banner1.StopTimer = True
        Banner1.Header = Header
        If TableName = "" Then btnSave.Visibility = Visibility.Hidden
        Try
            dt = bm.ExcuteAdapter(Statement)
            dt.TableName = "tbl"
            dt.Columns(0).ColumnName = FirstColumn
            dt.Columns(1).ColumnName = SecondColumn

            dv.Table = dt
            DataGridView1.ItemsSource = dv
            DataGridView1.Columns(0).Width = 120
            DataGridView1.Columns(1).Width = 300

            DataGridView1.SelectedIndex = 0
        Catch
        End Try
        txtID.Focus()
    End Sub

    Private Sub txtId_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.GotFocus
        Try
            dv.Sort = FirstColumn
        Catch
        End Try
    End Sub

    Private Sub txtName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.GotFocus
        Try
            dv.Sort = SecondColumn
        Catch
        End Try
    End Sub

    Private Sub txtId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.TextChanged, txtName.TextChanged
        dv.RowFilter = " [" & FirstColumn & "] like '%" & txtID.Text & "%' and [" & SecondColumn & "] like '%" & txtName.Text & "%'"
    End Sub

    Public SelectedId As Integer = 0
    Public SelectedName As String = ""
    
    Private Sub DataGridView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown
        Try
            If e.Key = System.Windows.Input.Key.Enter Then
                SelectedId = DataGridView1.Items(DataGridView1.SelectedIndex)(0)
                SelectedName = DataGridView1.Items(DataGridView1.SelectedIndex)(1)
                Close()
            ElseIf e.Key = Input.Key.Escape Then
                Close()
            ElseIf e.Key = Input.Key.Up Then
                DataGridView1.SelectedIndex = DataGridView1.SelectedIndex - 1
            ElseIf e.Key = Input.Key.Down Then
                DataGridView1.SelectedIndex = DataGridView1.SelectedIndex + 1
            End If
        Catch ex As Exception
        End Try
    End Sub


    Private Sub DataGridView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles DataGridView1.MouseDoubleClick
        Try
            SelectedId = DataGridView1.Items(DataGridView1.SelectedIndex)(0)
            SelectedName = DataGridView1.Items(DataGridView1.SelectedIndex)(1)
            Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnSave.Click
        If txtName.Text.Trim = "" Then Return
        txtID.Clear()
        For i As Integer = 0 To dv.Table.Rows.Count - 1
            If txtName.Text.Trim = dv.Table.Rows(i).Item(1).ToString Then Return
        Next
        If Not bm.AddItemToTable(TableName, txtName.Text.Trim) Then Return
        Window_Loaded(Nothing, Nothing)
        DataGridView1.SelectedIndex = dv.Table.Rows.Count - 1
        txtId_TextChanged(Nothing, Nothing)
    End Sub

End Class