Imports System.Data
Imports System.Windows.Forms

Public Class NLevels

    Public TableName As String = "NLevels"
    Public SubId As String = "Id"
    Public SubName As String = "Name"

    Dim dt As New DataTable
    Dim bm As New BasicMethods4Forms

    Dim Length As Integer = 0

    Private Sub NLevels_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtID.MaxLength = Length
        LoadTree()

        Dim v() As String = {SubId, SubName}
        bm.Fields = v
        Dim c() As Control = {txtID, txtName}
        bm.control = c

        Dim k() As String = {SubId}
        bm.KeyFields = k

        bm.Table_Name = TableName
        btnNew_Click(sender, e)

    End Sub

    Sub LoadTree()
        Dim Main As New MainMenuItem
        Main.NLevel = True
        Main.Opacity = 0
        Try
            'Main.Show()
        Catch ex As Exception
        End Try

        TreeView1.Nodes.Add("Contents")
        For i As Integer = 0 To Main.Menu1.Items.Count - 1
            Try
                Dim item As System.Windows.Controls.MenuItem
                item = Main.Menu1.Items(i)
                If Not item.Tag Is Nothing And Not item.Tag = "" Then Continue For
                Dim nn As New TreeNode
                nn.Name = item.Name
                nn.Text = item.Header
                TreeView1.Nodes(0).Nodes.Add(nn)
                loadNode(item, nn)
            Catch
            End Try
            Try
                Dim item As Separator
                item = Main.Menu1.Items(i)
                If Not item.Tag Is Nothing And Not item.Tag = "" Then Continue For
                Dim nn As New TreeNode
                nn.Name = item.Name
                nn.Text = "ــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                TreeView1.Nodes(0).Nodes.Add(nn)
            Catch
            End Try
        Next

        Try
            'Main.Nlvl = True
            'Main.Close()
        Catch ex As Exception
        End Try
        TreeView1.ExpandAll()
    End Sub

    Sub loadNode(ByVal item As System.Windows.Controls.MenuItem, ByVal nn As TreeNode)
        For i As Integer = 0 To item.Items.Count - 1
            Try
                Dim item2 As System.Windows.Controls.MenuItem
                item2 = item.Items(i)
                If Not item2.Tag Is Nothing And Not item2.Tag = "" Then Continue For
                Dim nn2 As New TreeNode
                nn2.Name = item2.Name
                nn2.Text = item2.Header
                nn.Nodes.Add(nn2)
                loadNode(item2, nn2)
            Catch
            End Try
            Try
                Dim item2 As Separator
                item2 = item.Items(i)
                If Not item2.Tag Is Nothing And Not item2.Tag = "" Then Continue For
                Dim nn2 As New TreeNode
                nn2.Name = item2.Name
                nn2.Text = "ــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــــ"
                nn.Nodes.Add(nn2)
            Catch
            End Try
        Next
    End Sub


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
        bm.FirstLast(New String() {SubId}, "Max", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub
    Sub FillControls()
        bm.FillControls()
        Dim dt As DataTable = bm.ExcuteAdapter("select * from NLevels where Id='" & txtID.Text & "'")
        For Each nn As TreeNode In TreeView1.Nodes(0).Nodes
            Try
                nn.Checked = dt.Rows(0)(nn.Name)
            Catch ex As Exception
                nn.Checked = False
            End Try
            FillSubNode(dt, nn)
        Next

    End Sub
    Sub FillSubNode(ByVal dt As DataTable, ByVal nn As TreeNode)
        For Each nn2 As TreeNode In nn.Nodes
            Try
                nn2.Checked = dt.Rows(0)(nn2.Name)
            Catch ex As Exception
                nn2.Checked = False
            End Try
            FillSubNode(dt, nn2)
        Next
    End Sub


    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        bm.NextPrevious(New String() {SubId}, New String() {txtID.Text}, "Next", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtName.Text.Trim = "" Then
            Return
        End If
        bm.DefineValues()
        If Not bm.Save(New String() {SubId}, New String() {txtID.Text.Trim}) Then Return
        SaveTree()
        btnNew_Click(sender, e)
        AllowClose = True
    End Sub

    Sub SaveTree()
        Dim ss As String = "Update NLevels set "
        For Each nn As TreeNode In TreeView1.Nodes(0).Nodes
            If nn.Name = "" Then Continue For
            If nn.Checked Then
                ss &= nn.Name & "=1,"
            Else
                ss &= nn.Name & "=0,"
            End If

            GetSubNode(nn, ss)
        Next
        ss = ss.Substring(0, ss.Length - 1) & " where Id='" & txtID.Text & "'"

        bm.ExcuteNonQuery(ss)
    End Sub

    Sub GetSubNode(ByVal nn As TreeNode, ByRef ss As String)
        For Each nn2 As TreeNode In nn.Nodes
            If nn2.Name = "" Then Continue For
            If nn2.Checked Then
                ss &= nn2.Name & "=1,"
            Else
                ss &= nn2.Name & "=0,"
            End If

            GetSubNode(nn2, ss)
        Next
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
        txtName.Clear()
        txtID.Text = bm.ExecuteScalar("select max(" & SubId & ")+1 from " & TableName)
        If txtID.Text = "" Then txtID.Text = "1"
        txtID.Text = txtID.Text.Trim.PadLeft(Length, "0")
        TreeView1.Nodes(0).Checked = False
        TreeView1_AfterCheck(Nothing, Nothing)
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من الDelete ؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & SubId & "='" & txtID.Text.Trim & "'")
            btnNew_Click(sender, e)
        End If
    End Sub

    Private Sub btnPrevios_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevios.Click
        bm.NextPrevious(New String() {SubId}, New String() {txtID.Text}, "Back", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub
    Dim lv As Boolean = False
    Private Sub txtID_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.Leave
        If lv Then
            Return
        End If
        lv = True
        txtID.Text = txtID.Text.Trim.PadLeft(Length, "0")
        bm.DefineValues()
        Dim dt As New DataTable
        bm.RetrieveAll(New String() {SubId}, New String() {txtID.Text.Trim}, dt)
        Dim s As String = txtID.Text.Trim
        If dt.Rows.Count = 0 Then
            bm.ClearControls()
            txtID.Text = s
            txtName.Focus()
            lv = False
            Return
        End If
        FillControls()
        lv = False
        'txtName.Text = dt(0)("Name")
    End Sub

    Private Sub btnNew_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.KeyDown, txtID.KeyDown, btnSave.KeyDown, btnPrevios.KeyDown, btnNext.KeyDown, btnNew.KeyDown, btnLast.KeyDown, btnFirst.KeyDown, btnDelete.KeyDown, btnClose.KeyDown
        If e.KeyData = Keys.Enter Then
            SelectNextControl(ctl:=sender, forward:=True, nested:=True, tabStopOnly:=True, wrap:=True)
        End If
    End Sub

    Private Sub txtID_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtID.KeyPress
        If e.KeyChar = "" Then
            Return
        End If
        Dim s As New TextBox
        s = sender
        'If e.KeyChar = "." And Not s.Text.Contains(".") Then
        '    Return
        'End If
        If Not IsNumeric(e.KeyChar) Then
            e.Handled = True
        End If

    End Sub

    Private Sub TreeView1_AfterCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterCheck
        Try
            For Each nn As TreeNode In e.Node.Nodes
                nn.Checked = e.Node.Checked
            Next
        Catch
        End Try
    End Sub

    Dim AllowClose As Boolean = False

End Class