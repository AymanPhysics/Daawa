Imports System.Data

Public Class Chart

    Public TableName As String = "Chart"
    Public SubId As String = "Id"

    Dim dt As New DataTable
    Dim bm As New BasicMethods

    Public Flag As Integer = 0

    Private Sub Chart_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If bm.TestIsLoaded Then Return
        bm.FillCombo("LinkFile", LinkFile, "")
        Dim v() As String = {SubId, "Name", "MainAccId", "MotionType", "SubType", "LinkFile"}
        bm.Fields = v
        Dim c() As Control = {txtID, txtName, AccNo, MotionType, SubType, LinkFile}
        bm.control = c

        Dim k() As String = {SubId}
        bm.KeyFields = k

        bm.Table_Name = TableName
        LoadTree()
        btnNew_Click(sender, e)
        
    End Sub


    Sub LoadTree()
        TreeView1.Items.Clear()
        Dim dt As DataTable = bm.ExcuteAdapter("Select * from Chart")
        TreeView1.Items.Add(New TreeViewItem With {.Header = "دليل الحسابات"})
        TreeView1.Items(0).Tag = ""
        TreeView1.Items(0).FontSize = 20
        Dim dr() As DataRow = dt.Select("MainAccId=0")
        For i As Integer = 0 To dr.Length - 1
            Try
                Dim nn As New TreeViewItem
                nn.Foreground = Brushes.DarkRed
                nn.FontSize = 18
                nn.Name = "Node_" & dr(i)("Id")
                nn.Tag = dr(i)("Id")
                nn.Header = dr(i)("Id") & "          " & dr(i)("Name")
                TreeView1.Items(0).Items.Add(nn)
                loadNode(dt, nn)
            Catch
            End Try
        Next
        CType(TreeView1.Items(0), TreeViewItem).IsExpanded = True
    End Sub

    Sub loadNode(ByVal dt As DataTable, ByVal nn As TreeViewItem)
        Dim dr() As DataRow = dt.Select("MainAccId=" & nn.Tag)
        For i As Integer = 0 To dr.Length - 1
            Try
                Dim nn2 As New TreeViewItem
                nn2.Foreground = Brushes.DarkBlue
                nn2.FontSize = nn.FontSize - 1
                nn2.Name = "Node_" & dr(i)("Id")
                nn2.Tag = dr(i)("Id")
                nn2.Header = dr(i)("Id") & "          " & dr(i)("Name")
                nn.Items.Add(nn2)
                loadNode(dt, nn2)
            Catch
            End Try
        Next
        nn.IsExpanded = True
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

    Sub FillControls()
        bm.FillControls()
        AccNo_LostFocus(Nothing, Nothing)
        txtName.Focus()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtName.Text.Trim = "" Then
            txtName.Focus()
            Return
        End If
        bm.DefineValues()
        If Not bm.Save(New String() {SubId}, New String() {txtID.Text.Trim}) Then Return
        UpdateMotionType(txtID.Text, MotionType.SelectedIndex)
        LoadTree()
        btnNew_Click(sender, e)
        AllowClose = True
    End Sub

    Sub UpdateMotionType(ByVal id As Integer, ByVal Type As Integer)
        Dim dt As DataTable = bm.ExcuteAdapter("select Id from Chart where MainAccId=" & id)
        For i As Integer = 0 To dt.Rows.Count - 1
            bm.ExcuteNonQuery("update Chart set MotionType=" & MotionType.SelectedIndex & " where Id=" & dt.Rows(i)(0))
            UpdateMotionType(dt.Rows(i)(0), Type)
        Next
    End Sub

    Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click

        bm.FirstLast(New String() {SubId}, "Min", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        ClearControls()
    End Sub

    Sub ClearControls()
        txtID.Clear()
        bm.ClearControls()
        AccNo_LostFocus(Nothing, Nothing)
        txtID.Focus()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من المسح؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & SubId & "='" & txtID.Text.Trim & "'")
            LoadTree()
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
            Dim s As String = txtID.Text
            ClearControls()
            txtID.Text = s
            txtName.Focus()
            lv = False
            Return
        End If
        FillControls()
        txtName.Focus()
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

    Private Sub AccNo_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles AccNo.LostFocus
        MotionType.IsEnabled = True
        bm.AccNoLostFocus(AccNo, AccName, 0, )
        If AccNo.Text = "" Then Return

        If AccNo.Text = txtID.Text Then
            bm.ShowMSG("لا يمكن للحساب العام أن يساوى الحساب الفرعى")
            AccNo.Clear()
            AccName.Clear()
        End If

        MotionType.IsEnabled = False
        MotionType.SelectedIndex = Val(bm.ExecuteScalar("select MotionType from Chart where Id=" & AccNo.Text))
    End Sub

    Private Sub AccNo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles AccNo.KeyUp
        bm.AccNoShowHelp(AccNo, AccName, e, 0, )
    End Sub


    Private Sub SubType_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles SubType.SelectionChanged
        If SubType.SelectedIndex = 1 Then
            LinkFile.IsEnabled = True
        Else
            LinkFile.SelectedIndex = 0
            LinkFile.IsEnabled = False
        End If
    End Sub

    Private Sub TreeView1_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of System.Object)) Handles TreeView1.SelectedItemChanged
        If TreeView1.SelectedItem Is Nothing Then Return
        txtID.Text = TreeView1.SelectedItem.Tag
        txtID_LostFocus(Nothing, Nothing)
        TreeView1.Focus()
    End Sub
End Class

