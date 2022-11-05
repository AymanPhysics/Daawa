Imports System.Data
Imports System.IO
Imports System.ComponentModel

Public Class Tasks
    Public TableName As String = "Tasks"
    Public SubId As String = "Id"



    Dim dt As New DataTable
    Dim bm As New BasicMethods
    WithEvents G As New MyGrid

    Dim WithEvents BackgroundWorker1 As New BackgroundWorker
    Public LagnaId, OperationId As Integer

    Private Sub Tasks_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If bm.TestIsLoaded Then Return
        bm.Fields = New String() {SubId, "lagnaId", "OperationId", "CityId", "AreaId", "TownId", "Details", "Daydate", "Daydate2", "Time1", "Time2", "Cost", "Notes"}
        bm.control = New Control() {txtID, TextLagnaId, TextOperationId, CityId, AreaId, TownId, Details, DayDate, DayDate2, Time1, Time2, Cost, Notes}
        bm.KeyFields = New String() {SubId}

        bm.Table_Name = TableName
        bm.AppendWhere = " and LagnaId=" & LagnaId & " and OperationId=" & OperationId
        
        LoadWFH()
        btnNew_Click(sender, e)
    End Sub


    Sub NewId()
        txtID.Clear()
        txtID.IsEnabled = False
    End Sub

    Sub UndoNewId()
        txtID.IsEnabled = True
    End Sub


    Structure GC
        Shared CaseId As String = "CaseId"
        Shared CaseName As String = "CaseName"
        Shared Id As String = "Id"
        Shared Name As String = "Name"
        Shared Qty As String = "Qty"
        Shared Price As String = "Price"
        Shared Value As String = "Value"
    End Structure

    Private Sub LoadWFH()
        WFH.Child = G

        G.Grid.ForeColor = System.Drawing.Color.DarkBlue
        G.Grid.Columns.Add(GC.CaseId, "كود الحالة")
        G.Grid.Columns.Add(GC.CaseName, "اسم الحالة")
        G.Grid.Columns.Add(GC.Id, "كود البند")
        G.Grid.Columns.Add(GC.Name, "اسم البند")
        G.Grid.Columns.Add(GC.Qty, "الكمية")
        G.Grid.Columns.Add(GC.Price, "السعر")
        G.Grid.Columns.Add(GC.Value, "القيمة")

        G.Grid.Columns(GC.CaseId).FillWeight = 100
        G.Grid.Columns(GC.CaseName).FillWeight = 300
        G.Grid.Columns(GC.Id).FillWeight = 100
        G.Grid.Columns(GC.Name).FillWeight = 300
        G.Grid.Columns(GC.Qty).FillWeight = 100
        G.Grid.Columns(GC.Price).FillWeight = 100
        G.Grid.Columns(GC.Value).FillWeight = 100

        G.Grid.Columns(GC.CaseName).ReadOnly = True
        G.Grid.Columns(GC.Name).ReadOnly = True
        G.Grid.Columns(GC.Value).ReadOnly = True
        G.Grid.AllowUserToDeleteRows = True

        AddHandler G.Grid.CellEndEdit, AddressOf GridCalcRow
        AddHandler G.Grid.KeyDown, AddressOf GridKeyDown
        AddHandler G.Grid.EditingControlShowing, AddressOf GridEditingControlShowing
    End Sub

    Private Sub GridCalcRow(ByVal sender As Object, ByVal e As Forms.DataGridViewCellEventArgs)
        If G.Grid.Columns(e.ColumnIndex).Name = GC.CaseId Then
            AddCase(G.Grid.Rows(e.RowIndex).Cells(GC.CaseId).Value, e.RowIndex, 0)
        else If G.Grid.Columns(e.ColumnIndex).Name = GC.Id Then
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

            Dim dt As DataTable = bm.ExcuteAdapter("Select * From OutcomeReasons where Id='" & Id & "'")
            Dim dr() As DataRow = dt.Select("Id='" & Id & "'")
            If dr.Length = 0 Then
                If Not G.Grid.Rows(i).Cells(GC.Id).Value Is Nothing AndAlso G.Grid.Rows(i).Cells(GC.Id).Value <> "" Then bm.ShowMSG("هذا البند غير موجود")
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

    Sub AddCase(ByVal Id As String, Optional ByVal i As Integer = -1, Optional ByVal Add As Decimal = 1)
        Try
            Dim Exists As Boolean = False
            Dim Move As Boolean = False
            If i = -1 Then Move = True

            G.Grid.AutoSizeColumnsMode = Forms.DataGridViewAutoSizeColumnsMode.Fill
            If i = -1 Then
                For x As Integer = 0 To G.Grid.Rows.Count - 1
                    If Not G.Grid.Rows(x).Cells(GC.CaseId).Value Is Nothing AndAlso G.Grid.Rows(x).Cells(GC.CaseId).Value.ToString = Id.ToString Then
                        i = x
                        Exists = True
                        GoTo Br
                    End If
                Next
                i = G.Grid.Rows.Add()
Br:
            End If

            Dim dt As DataTable = bm.ExcuteAdapter("Select * From Cases where SeasonId=dbo.GetCurrentSeason() and Id='" & Id & "'")
            Dim dr() As DataRow = dt.Select("Id='" & Id & "'")
            If dr.Length = 0 Then
                If Not G.Grid.Rows(i).Cells(GC.CaseId).Value Is Nothing AndAlso G.Grid.Rows(i).Cells(GC.CaseId).Value <> "" Then bm.ShowMSG("هذه الحالة غير موجودة")
                G.Grid.Rows(i).Cells(GC.CaseId).Value = ""
                G.Grid.Rows(i).Cells(GC.CaseName).Value = ""
                Return
            End If
            G.Grid.Rows(i).Cells(GC.CaseId).Value = dr(0)("Id")
            G.Grid.Rows(i).Cells(GC.CaseName).Value = dr(0)("ArName")
        Catch
        End Try
    End Sub

    Dim lop As Boolean = False

    Sub ClearRow(ByVal i As Integer)
        'G.Grid.Rows(i).Cells(GC.CaseId).Value = Nothing
        'G.Grid.Rows(i).Cells(GC.CaseName).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Id).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Name).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Qty).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Price).Value = Nothing
        G.Grid.Rows(i).Cells(GC.Value).Value = Nothing
    End Sub

    Sub FillControls()
        UndoNewId()
        bm.FillControls()
        CityId_LostFocus(Nothing, Nothing)
        AreaId_LostFocus(Nothing, Nothing)
        TownId_LostFocus(Nothing, Nothing)

        LoadTree()
        LoadTree2()

        Dim dt As DataTable = bm.ExcuteAdapter("select * from TaskOutcomeReasons where TaskId=" & txtID.Text & bm.AppendWhere)

        G.Grid.Rows.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            G.Grid.Rows.Add()
            G.Grid.Rows(i).Cells(GC.CaseId).Value = dt.Rows(i)("CaseId").ToString
            G.Grid.Rows(i).Cells(GC.CaseName).Value = dt.Rows(i)("CaseName").ToString
            G.Grid.Rows(i).Cells(GC.Id).Value = dt.Rows(i)("ReasonId").ToString
            G.Grid.Rows(i).Cells(GC.Name).Value = dt.Rows(i)("ReasonName").ToString
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


    Dim AllowPrint As Boolean = False
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        AllowPrint = False
        If Details.Text.Trim = "" Then
            bm.ShowMSG("برجاء تحديد الموضوع")
            Details.Focus()
            Return
        End If
        Cost.Text = Val(Cost.Text)


        Dim State As BasicMethods.SaveState = BasicMethods.SaveState.Update
        If txtID.Text.Trim = "" Then
            txtID.Text = bm.ExecuteScalar("select max(Id)+1 from " & TableName & " where lagnaId=" & TextLagnaId.Text & " and OperationId=" & TextOperationId.Text)
            If txtID.Text = "" Then txtID.Text = "1"
            lblLastEntry.Content = txtID.Text
            'Begin Animation
            State = BasicMethods.SaveState.Insert
        End If


        bm.DefineValues()
        If Not bm.Save(New String() {SubId}, New String() {txtID.Text.Trim}, State) Then
            If State = BasicMethods.SaveState.Insert Then
                txtID.Text = ""
                lblLastEntry.Content = ""
            End If
            Return
        End If

        If Not bm.SaveGrid(G.Grid, "TaskOutcomeReasons", New String() {"lagnaId", "OperationId", "TaskId"}, New String() {TextLagnaId.Text, TextOperationId.Text, txtID.Text}, New String() {"CaseId", "CaseName", "ReasonId", "ReasonName", "Qty", "Price", "Value"}, New String() {GC.CaseId, GC.CaseName, GC.Id, GC.Name, GC.Qty, GC.Price, GC.Value}, New VariantType() {VariantType.Integer, VariantType.String, VariantType.Integer, VariantType.String, VariantType.Decimal, VariantType.Decimal, VariantType.Decimal}, New String() {GC.CaseId, GC.Id}) Then Return

        AllowPrint = True

        If Not DontClear Then btnNew_Click(sender, e)
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
        CityName.Clear()
        AreaName.Clear()
        TownName.Clear()

        TreeView1.Items.Clear()
        TreeView2.Items.Clear()
        TextLagnaId.Text = LagnaId
        TextOperationId.Text = OperationId
        G.Grid.Rows.Clear()
        txtID.Text = bm.ExecuteScalar("select max(" & SubId & ")+1 from " & TableName & " where 1=1 " & bm.AppendWhere)
        If txtID.Text = "" Then txtID.Text = "1"
        NewId()
        Details.Focus()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من المسح؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & SubId & "='" & txtID.Text.Trim & "' " & bm.AppendWhere)
            bm.ExcuteNonQuery("delete from TaskAttachments where TaskId='" & txtID.Text & "' " & bm.AppendWhere)
            bm.ExcuteNonQuery("delete from TaskPersons where TaskId='" & txtID.Text & "' " & bm.AppendWhere)
            bm.ExcuteNonQuery("delete from TaskOutcomeReasons where TaskId='" & txtID.Text & "' " & bm.AppendWhere)
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
            Details.Focus()
            lv = False
            Return
        End If
        FillControls()
        lv = False
        Details.SelectAll()
        Details.Focus()
        Details.SelectAll()
        Details.Focus()
        'txtName.Text = dt(0)("Name")
    End Sub

    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles txtID.KeyDown
        bm.MyKeyPress(sender, e)
    End Sub

    Private Sub txtID_KeyPress2(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Cost.KeyDown
        bm.MyKeyPress(sender, e, True)
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


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        Dim o As New Forms.OpenFileDialog
        o.Multiselect = True
        If o.ShowDialog = Forms.DialogResult.OK Then
            For i As Integer = 0 To o.FileNames.Length - 1
                bm.SaveFile("TaskAttachments", "LagnaId", LagnaId, "OperationId", OperationId, "TaskId", txtID.Text, "AttachedName", (o.FileNames(i).Split("\"))(o.FileNames(i).Split("\").Length - 1), "Image", o.FileNames(i))
            Next
        End If
        LoadTree()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button4.Click
        Try
            MyImagedata = Nothing
            If CType(TreeView1.SelectedItem, TreeViewItem).FontSize <> 18 Then Return
            Dim s As New Forms.SaveFileDialog
            s.FileName = CType(TreeView1.SelectedItem, TreeViewItem).Header

            If IsNothing(sender) Then
                MyBath = bm.GetNewTempName(s.FileName)
            Else
                If Not s.ShowDialog = Forms.DialogResult.OK Then Return
                MyBath = s.FileName
            End If

            Button4.IsEnabled = False
            F1 = txtID.Text
            F2 = CType(TreeView1.SelectedItem, TreeViewItem).Tag
            BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
        End Try
    End Sub
    Dim F2 As String = "", F1 As String = ""
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Dim myCommand As SqlClient.SqlCommand
            myCommand = New SqlClient.SqlCommand("select Image from TaskAttachments where TaskId='" & F1 & "' and AttachedName='" & F2 & "'" & bm.AppendWhere, con)
            MyImagedata = CType(myCommand.ExecuteScalar(), Byte())
        Catch ex As Exception
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            File.WriteAllBytes(MyBath, MyImagedata)
            Process.Start(MyBath)
        Catch ex As Exception
        End Try
        Button4.IsEnabled = True
    End Sub

    Dim MyImagedata() As Byte
    Dim MyBath As String = ""
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button3.Click
        Try
            If CType(TreeView1.SelectedItem, TreeViewItem).FontSize = 18 Then
                If bm.ShowDeleteMSG("هل أنت متأكد من إجراء عملية Delete الملف """ & TreeView1.SelectedItem.Header & """?") Then
                    bm.ExcuteNonQuery("delete from TaskAttachments where TaskId='" & txtID.Text & "' and AttachedName='" & TreeView1.SelectedItem.Header & "'" & bm.AppendWhere)
                    LoadTree()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadTree()
        Dim dt As DataTable = bm.ExcuteAdapter("select AttachedName from TaskAttachments where TaskId=" & txtID.Text & bm.AppendWhere)
        TreeView1.Items.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim nn As New TreeViewItem
            nn.Foreground = Brushes.DarkRed
            nn.FontSize = 18
            nn.Tag = dt.Rows(i)(0).ToString
            nn.Header = dt.Rows(i)(0).ToString
            TreeView1.Items.Add(nn)
        Next
    End Sub

    Private Sub LoadTree2()
        Dim dt As DataTable = bm.ExcuteAdapter("select PersonId,dbo.GetEmpArName(PersonId) from TaskPersons where TaskId=" & txtID.Text & bm.AppendWhere)
        TreeView2.Items.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim nn As New TreeViewItem
            nn.Foreground = Brushes.DarkRed
            nn.FontSize = 18
            nn.Tag = dt.Rows(i)(0).ToString
            nn.Header = dt.Rows(i)(0).ToString & " - " & dt.Rows(i)(1).ToString
            TreeView2.Items.Add(nn)
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        Dim hh As New Help
        hh.Header = "الموظفين"
        hh.Statement = "select cast(Id as varchar(100)) Id,ArName Name from Employees where Stopped=0"
        hh.ShowDialog()
        If hh.SelectedId = 0 Then Return
        bm.SaveText("TaskPersons", "LagnaId", LagnaId, "OperationId", OperationId, "TaskId", txtID.Text, "PersonId", hh.SelectedId)
        LoadTree2()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button6.Click
        Try
            If CType(TreeView2.SelectedItem, TreeViewItem).FontSize = 18 Then
                If bm.ShowDeleteMSG("هل أنت متأكد من  Delete الموظف """ & TreeView2.SelectedItem.Header & """?") Then
                    bm.ExcuteNonQuery("delete from TaskPersons where TaskId='" & txtID.Text & "' and PersonId='" & TreeView2.SelectedItem.Tag & "'" & bm.AppendWhere)
                    LoadTree2()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Dim c1, c2 As New TextBox
    Private Sub GridKeyDown(ByVal sender As Object, ByVal e As Forms.KeyEventArgs)

        Dim statement As String = "select cast(Id as varchar(100)) Id,ArName from Cases where SeasonId=(select MAX(s.SeasonId) from Cases s)"
        If Val(CityId.Text) <> 0 Then statement &= " and CityId=" & CityId.Text
        If Val(AreaId.Text) <> 0 Then statement &= " and AreaId=" & AreaId.Text
        If Val(TownId.Text) <> 0 Then statement &= " and TownId=" & TownId.Text

        If G.Grid.CurrentCell Is G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.CaseId) AndAlso bm.ShowHelpGrid("الحالات", G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.CaseId), G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.CaseName), e, statement) Then
            G.Grid.CurrentCell = G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Qty)
            G.Grid.CurrentCell = G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.CaseId)
        ElseIf G.Grid.CurrentCell Is G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Id) AndAlso bm.ShowHelpGrid("بنود المصروفات", G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Id), G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Name), e, "select cast(Id as varchar(100)) Id,Name from OutcomeReasons") Then
            G.Grid.CurrentCell = G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Qty)
            G.Grid.CurrentCell = G.Grid.Rows(G.Grid.CurrentRow.Index).Cells(GC.Id)
        End If
    End Sub

    Private Sub GridEditingControlShowing(ByVal sender As Object, ByVal e As Forms.DataGridViewEditingControlShowingEventArgs)
        Dim c = e.Control
        RemoveHandler c.KeyDown, AddressOf GridKeyDown
        AddHandler c.KeyDown, AddressOf GridKeyDown
    End Sub


    Private Sub TreeView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TreeView1.MouseDoubleClick
        Button4_Click(Nothing, Nothing)
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
    Private Sub CityId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CityId.LostFocus
        bm.LostFocus(CityId, CityName, "select Name from Cities where Id=" & CityId.Text.Trim())
    End Sub
    Private Sub AreaId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles AreaId.LostFocus
        bm.LostFocus(AreaId, AreaName, "select Name from Areas where CityId=" & CityId.Text.Trim() & " and Id=" & AreaId.Text.Trim)
    End Sub
    Private Sub TownId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles TownId.LostFocus
        bm.LostFocus(TownId, TownName, "select Name from Towns where CityId=" & CityId.Text.Trim() & " and AreaId=" & AreaId.Text.Trim() & " and Id=" & TownId.Text.Trim)
    End Sub


    Dim DontClear As Boolean = False
    Private Sub btnPrint_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint.Click
        DontClear = True
        btnSave_Click(Nothing, Nothing)
        DontClear = False
        If Not AllowPrint Then Return

        Dim rpt As New ReportViewer
        rpt.paraname = {"Header", "@lagnaId", "@OperationId", "@TaskId"}
        rpt.paravalue = {CType(CType(Parent, TabItem).Header, TabsHeader).MyTabHeader, LagnaId, OperationId, Val(txtID.Text)}
        rpt.Header = Md.MyProject.ToString
        rpt.RptPath = "Tasks.rpt"
        rpt.ShowDialog()
    End Sub
End Class
