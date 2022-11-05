Imports System.Data

Public Class HelpCustomers
    Dim bm As New BasicMethods
    Public FirstColumn As String = "الكــــــــــــود", SecondColumn As String = "الاســــــــــــم", ThirdColumn As String = "الرقم القومى", FourthColumn As String = "التليفــــون", FifthColumn As String = "الموبايـــــل", SixthColumn As String = "العنـــــــــــوان", SeventhColumn As String = "المحافظة", EightthColumn As String = "المركز", NinethColumn As String = "القرية"

    Dim dt As New DataTable
    Dim dv As New DataView
    Public Header As String = ""
    Private Sub Help_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Loaded
        Banner1.StopTimer = True
        Banner1.Header = "الحالات" 'Header
        Try
            dt = bm.ExcuteAdapter("CasesSearch")
            dt.TableName = "tbl"
            dt.Columns(0).ColumnName = FirstColumn
            dt.Columns(1).ColumnName = SecondColumn
            dt.Columns(2).ColumnName = ThirdColumn
            dt.Columns(3).ColumnName = FourthColumn
            dt.Columns(4).ColumnName = FifthColumn
            dt.Columns(5).ColumnName = SixthColumn
            dt.Columns(6).ColumnName = SeventhColumn
            dt.Columns(7).ColumnName = EightthColumn
            dt.Columns(8).ColumnName = NinethColumn

            dv.Table = dt
            DataGridView1.ItemsSource = dv
            DataGridView1.Columns(0).Width = 85
            DataGridView1.Columns(1).Width = 180
            DataGridView1.Columns(2).Width = 85
            DataGridView1.Columns(3).Width = 85
            DataGridView1.Columns(4).Width = 85
            DataGridView1.Columns(5).Width = 120
            DataGridView1.Columns(6).Width = 85
            DataGridView1.Columns(7).Width = 85
            DataGridView1.Columns(8).Width = 85

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

    Private Sub NationalId_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NationalId.GotFocus
        Try
            dv.Sort = ThirdColumn
        Catch
        End Try
    End Sub

    Private Sub txtTel_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTel.GotFocus
        Try
            dv.Sort = FourthColumn
        Catch
        End Try
    End Sub

    Private Sub txtMob_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMob.GotFocus
        Try
            dv.Sort = FifthColumn
        Catch
        End Try
    End Sub

    Private Sub txtAddress_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddress.GotFocus
        Try
            dv.Sort = SixthColumn
        Catch
        End Try
    End Sub

    Private Sub CityName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CityName.GotFocus
        Try
            dv.Sort = SeventhColumn
        Catch
        End Try
    End Sub

    Private Sub AreaName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AreaName.GotFocus
        Try
            dv.Sort = EightthColumn
        Catch
        End Try
    End Sub

    Private Sub TownName_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TownName.GotFocus
        Try
            dv.Sort = NinethColumn
        Catch
        End Try
    End Sub

    Private Sub txtId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.TextChanged, txtName.TextChanged, NationalId.TextChanged, txtTel.TextChanged, txtMob.TextChanged, txtAddress.TextChanged, CityName.TextChanged, AreaName.TextChanged, TownName.TextChanged
        dv.RowFilter = " [" & FirstColumn & "] like '%" & txtID.Text & "%' and [" & SecondColumn & "] like '%" & txtName.Text & "%' and [" & ThirdColumn & "] like '%" & NationalId.Text & "%' and ([" & FourthColumn & "] like '%" & txtTel.Text & "%' or [" & FifthColumn & "] like '%" & txtTel.Text & "%') and ([" & FourthColumn & "] like '%" & txtMob.Text & "%' or [" & FifthColumn & "] like '%" & txtMob.Text & "%') and [" & SixthColumn & "] like '%" & txtAddress.Text & "%' and [" & SeventhColumn & "] like '%" & CityName.Text & "%' and [" & EightthColumn & "] like '%" & AreaName.Text & "%' and [" & NinethColumn & "] like '%" & TownName.Text & "%'"
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




End Class