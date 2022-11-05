Imports System.Data
Imports System.ComponentModel
Imports System.IO

Public Class Cases
    Public TableName As String = "Cases"
    Public MainId As String = "SeasonId"
    Public SubId As String = "Id"


    Dim dt As New DataTable
    Dim bm As New BasicMethods

    WithEvents G1 As New MyGrid
    WithEvents G2 As New MyGrid
    WithEvents G3 As New MyGrid

    Public Flag As Integer = 0
    Dim WithEvents BackgroundWorker1 As New BackgroundWorker

    Private Sub BasicForm_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        If bm.TestIsLoaded Then Return
        
        LoadWFH1()
        LoadWFH2()
        LoadWFH3()

        bm.Fields = New String() {MainId, SubId, "ArName", "Gender", "SSN", "CityId", "AreaId", "TownId", "SubTownId", "ReligionId", "Address", "DateOfBirth", "JobId", "Notes", "Manager", "SystemUser", "BankAccount", "NationalId", "HomePhone", "Mobile", "Email", "Password", "EnName", "LevelId", "SalaryType", "SearchIndex", "mm", "GeneralManager", "HasAttendance", "BasicSalary", "Stopped", "Accountant", "Board", "SearchDate", "Cashier", "Waiter", "Deliveryman", "EmpId", "CaseTypeId", "SearcherNotes"}
        bm.control = New Control() {SeasonId, txtID, ArName, Gender, SSN, CityId, AreaId, TownId, SubTownId, ReligionId, Address, DateOfBirth, JobId, Notes, Manager, SystemUser, BankAccount, NationalId, HomePhone, Mobile, Email, Password, EnName, LevelId, SalaryType, SearchIndex, mm, GeneralManager, HasAttendance, BasicSalary, Stopped, Accountant, Board, SearchDate, Cashier, Waiter, Deliveryman, EmpId, CaseTypeId, SearcherNotes}
        bm.KeyFields = New String() {MainId, SubId}
        bm.Table_Name = TableName
        SeasonId.Text = bm.ExecuteScalar("select CurrentSeason from Statics")
        SeasonId_LostFocus(Nothing, Nothing)
        

    End Sub



    Sub NewId()
        txtID.Clear()
        txtID.IsEnabled = False
    End Sub

    Sub UndoNewId()
        txtID.IsEnabled = True
    End Sub



    Structure GC1
        Shared Id As String = "Id"
        Shared Name As String = "Name"
        Shared Descrip As String = "Descrip"
        Shared JobId As String = "JobId"
        Shared BirthDate As String = "BirthDate"
        Shared IllTypeId As String = "IllTypeId"
        Shared ProblemTypeId As String = "ProblemTypeId"
        Shared CaseLevelId As String = "CaseLevelId"
        Shared NationalId As String = "NationalId"
        Shared Notes As String = "Notes"
    End Structure

    Private Sub LoadWFH1()
        WFH1.Child = G1
        G1.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        Dim Descrip As New Forms.DataGridViewComboBoxColumn
        Descrip.HeaderText = "صلة القرابة"
        Descrip.Name = GC1.Descrip
        bm.FillCombo("select Id,Name from SelaKaraba union select 0 Id,'-' Name", Descrip)

        Dim cJobId As New Forms.DataGridViewComboBoxColumn
        cJobId.HeaderText = "الوظيفة"
        cJobId.Name = GC1.JobId
        bm.FillCombo("select Id,Name from CaseJobs union select 0 Id,'-' Name", cJobId)

        Dim IllTypeId As New Forms.DataGridViewComboBoxColumn
        IllTypeId.HeaderText = "المرض"
        IllTypeId.Name = GC1.IllTypeId
        bm.FillCombo("select Id,Name from IllTypes union select 0 Id,'-' Name", IllTypeId)

        Dim ProblemTypeId As New Forms.DataGridViewComboBoxColumn
        ProblemTypeId.HeaderText = "الشكوى"
        ProblemTypeId.Name = GC1.ProblemTypeId
        bm.FillCombo("select Id,Name from ProblemTypes union select 0 Id,'-' Name", ProblemTypeId)

        Dim CaseLevelId As New Forms.DataGridViewComboBoxColumn
        CaseLevelId.HeaderText = "درجة الحالة"
        CaseLevelId.Name = GC1.CaseLevelId
        bm.FillCombo("select Id,Name from CaseLevels union select 0 Id,'-' Name", CaseLevelId)

        G1.Grid.ForeColor = System.Drawing.Color.DarkBlue
        G1.Grid.Columns.Add(GC1.Id, "مسلسل")
        G1.Grid.Columns.Add(GC1.Name, "الاسم")
        G1.Grid.Columns.Add(Descrip)
        G1.Grid.Columns.Add(cJobId)
        G1.Grid.Columns.Add(GC1.BirthDate, "تاريخ الميلاد")
        G1.Grid.Columns.Add(IllTypeId)
        G1.Grid.Columns.Add(ProblemTypeId)
        G1.Grid.Columns.Add(CaseLevelId)
        G1.Grid.Columns.Add(GC1.NationalId, "الرقم القومى")
        G1.Grid.Columns.Add(GC1.Notes, "ملاحظات")

        G1.Grid.Columns(GC1.Id).FillWeight = 100
        G1.Grid.Columns(GC1.Name).FillWeight = 300
        G1.Grid.Columns(GC1.Descrip).FillWeight = 180
        G1.Grid.Columns(GC1.JobId).FillWeight = 180
        G1.Grid.Columns(GC1.BirthDate).FillWeight = 180
        G1.Grid.Columns(GC1.IllTypeId).FillWeight = 200
        G1.Grid.Columns(GC1.ProblemTypeId).FillWeight = 200
        G1.Grid.Columns(GC1.CaseLevelId).FillWeight = 200
        G1.Grid.Columns(GC1.NationalId).FillWeight = 180
        G1.Grid.Columns(GC1.Notes).FillWeight = 300

        G1.Grid.Columns(GC1.Id).ReadOnly = True
        G1.Grid.AllowUserToDeleteRows = True
        G1.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        AddHandler G1.Grid.EditingControlShowing, AddressOf G1_EditingControlShowing
        AddHandler G1.Grid.CellEndEdit, AddressOf G1_CellEndEdit
        AddHandler G1.Grid.UserDeletedRow, AddressOf G1_UserDeletedRow
    End Sub


    Structure GC2
        Shared Id As String = "Id"
        Shared Name As String = "Name"
        Shared Count1 As String = "Count1"
        Shared Count2 As String = "Count2"
        Shared Diff As String = "Diff"
        Shared NeedPeriodId As String = "NeedPeriodId"
        Shared DayDate As String = "DayDate"
        Shared Notes As String = "Notes"
    End Structure

    Private Sub LoadWFH2()
        WFH2.Child = G2
        G2.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        G2.Grid.ForeColor = System.Drawing.Color.DarkBlue
        G2.Grid.Columns.Add(GC2.Id, "كود البند")
        G2.Grid.Columns.Add(GC2.Name, "اسم البند")
        G2.Grid.Columns.Add(GC2.Count1, "العدد")
        G2.Grid.Columns.Add(GC2.Count2, "مراجع")
        G2.Grid.Columns.Add(GC2.Diff, "تأكيد بيانات")

        Dim NeedPeriodId As New Forms.DataGridViewComboBoxColumn
        NeedPeriodId.HeaderText = "التكرار"
        NeedPeriodId.Name = GC2.NeedPeriodId
        bm.FillCombo("select Id,Name from NeedPeriod union select 0 Id,'-' Name", NeedPeriodId)

        G2.Grid.Columns.Add(NeedPeriodId)

        G2.Grid.Columns.Add(GC2.DayDate, "التاريخ")
        G2.Grid.Columns.Add(GC2.Notes, "ملاحظات")

        G2.Grid.Columns(GC2.Id).FillWeight = 80
        G2.Grid.Columns(GC2.Name).FillWeight = 1200
        G2.Grid.Columns(GC2.Count1).FillWeight = 80
        G2.Grid.Columns(GC2.Count2).FillWeight = 80
        G2.Grid.Columns(GC2.Diff).FillWeight = 80
        G2.Grid.Columns(GC2.NeedPeriodId).FillWeight = 80
        G2.Grid.Columns(GC2.DayDate).FillWeight = 80
        G2.Grid.Columns(GC2.Notes).FillWeight = 2000

        G2.Grid.Columns(GC2.Name).ReadOnly = True
        G2.Grid.Columns(GC2.Diff).ReadOnly = True
        G2.Grid.AllowUserToDeleteRows = True
        G2.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        AddHandler G2.Grid.CellEndEdit, AddressOf G2_CellEndEdit
        AddHandler G2.Grid.KeyDown, AddressOf G2_KeyDown
        AddHandler G2.Grid.EditingControlShowing, AddressOf G2_EditingControlShowing
    End Sub

    Structure GC3
        Shared Id As String = "Id"
        Shared Name As String = "Name"
        Shared Value As String = "Value"
        Shared MassValue As String = "MassValue"
        Shared Notes As String = "Notes"
    End Structure

    Private Sub LoadWFH3()
        WFH3.Child = G3
        G3.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        G3.Grid.ForeColor = System.Drawing.Color.DarkBlue
        G3.Grid.Columns.Add(GC3.Id, "كود الدخل")
        G3.Grid.Columns.Add(GC3.Name, "اسم الدخل")
        G3.Grid.Columns.Add(GC3.Value, "قيمة")
        G3.Grid.Columns.Add(GC3.MassValue, "عينى")
        G3.Grid.Columns.Add(GC3.Notes, "ملاحظات")

        G3.Grid.Columns(GC3.Id).FillWeight = 80
        G3.Grid.Columns(GC3.Name).FillWeight = 250
        G3.Grid.Columns(GC3.Value).FillWeight = 80
        G3.Grid.Columns(GC3.MassValue).FillWeight = 80
        G3.Grid.Columns(GC3.Notes).FillWeight = 400

        G3.Grid.Columns(GC3.Name).ReadOnly = True
        G3.Grid.AllowUserToDeleteRows = True
        G3.Grid.EditMode = Forms.DataGridViewEditMode.EditOnEnter

        AddHandler G3.Grid.CellEndEdit, AddressOf G3_CellEndEdit
        AddHandler G3.Grid.KeyDown, AddressOf G3_KeyDown
        AddHandler G3.Grid.EditingControlShowing, AddressOf G3_EditingControlShowing
    End Sub


    Private Sub btnLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLast.Click
        bm.FirstLast(New String() {MainId, SubId}, "Max", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Sub FillControls()
        UndoNewId()
        bm.FillControls()
        bm.GetImage(TableName, New String() {MainId, SubId}, New String() {SeasonId.Text, txtID.Text.Trim}, "Image", Image1)
        CityId_LostFocus(Nothing, Nothing)
        AreaId_LostFocus(Nothing, Nothing)
        TownId_LostFocus(Nothing, Nothing)
        SubTownId_LostFocus(Nothing, Nothing)
        ReligionId_LostFocus(Nothing, Nothing)
        LevelId_LostFocus(Nothing, Nothing)
        JobId_LostFocus(Nothing, Nothing)
        CaseTypeId_LostFocus(Nothing, Nothing)
        EmpId_LostFocus(Nothing, Nothing)
        LoadTree()


        Dim dt As DataTable = bm.ExcuteAdapter("select * from CaseDetails where SeasonId=" & SeasonId.Text & " and CaseId=" & txtID.Text)

        G1.Grid.Rows.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            G1.Grid.Rows.Add()
            G1.Grid.Rows(i).Cells(GC1.Id).Value = dt.Rows(i)("Id").ToString
            G1.Grid.Rows(i).Cells(GC1.Name).Value = dt.Rows(i)("Name").ToString
            G1.Grid.Rows(i).Cells(GC1.Descrip).Value = dt.Rows(i)("Descrip").ToString
            G1.Grid.Rows(i).Cells(GC1.JobId).Value = dt.Rows(i)("JobId").ToString
            G1.Grid.Rows(i).Cells(GC1.IllTypeId).Value = dt.Rows(i)("IllTypeId").ToString
            G1.Grid.Rows(i).Cells(GC1.ProblemTypeId).Value = dt.Rows(i)("ProblemTypeId").ToString
            G1.Grid.Rows(i).Cells(GC1.CaseLevelId).Value = dt.Rows(i)("CaseLevelId").ToString
            G1.Grid.Rows(i).Cells(GC1.NationalId).Value = dt.Rows(i)("NationalId").ToString
            G1.Grid.Rows(i).Cells(GC1.Notes).Value = dt.Rows(i)("Notes").ToString
            Try
                G1.Grid.Rows(i).Cells(GC1.BirthDate).Value = dt.Rows(i)("BirthDate").ToString.Substring(0, 10)
            Catch ex As Exception
                G1.Grid.Rows(i).Cells(GC1.BirthDate).Value = ""
            End Try
        Next
        G1.Grid.RefreshEdit()

        dt = bm.ExcuteAdapter("select * from CaseNeeds where SeasonId=" & SeasonId.Text & " and CaseId=" & txtID.Text)
        G2.Grid.Rows.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            G2.Grid.Rows.Add()
            G2.Grid.Rows(i).Cells(GC2.Id).Value = dt.Rows(i)("Id").ToString
            G2.Grid.Rows(i).Cells(GC2.Name).Value = dt.Rows(i)("Name").ToString
            G2.Grid.Rows(i).Cells(GC2.Count1).Value = dt.Rows(i)("Count1").ToString
            G2.Grid.Rows(i).Cells(GC2.Count2).Value = dt.Rows(i)("Count2").ToString
            G2.Grid.Rows(i).Cells(GC2.Diff).Value = dt.Rows(i)("Diff").ToString
            G2.Grid.Rows(i).Cells(GC2.NeedPeriodId).Value = dt.Rows(i)("NeedPeriodId").ToString
            G2.Grid.Rows(i).Cells(GC2.Notes).Value = dt.Rows(i)("Notes").ToString
            Try
                G2.Grid.Rows(i).Cells(GC2.DayDate).Value = dt.Rows(i)("DayDate").ToString.Substring(0, 10)
            Catch ex As Exception
                G2.Grid.Rows(i).Cells(GC2.DayDate).Value = ""
            End Try
        Next
        G2.Grid.RefreshEdit()

        dt = bm.ExcuteAdapter("select * from CaseIncomeTypes where SeasonId=" & SeasonId.Text & " and CaseId=" & txtID.Text)
        G3.Grid.Rows.Clear()
        For i As Integer = 0 To dt.Rows.Count - 1
            G3.Grid.Rows.Add()
            G3.Grid.Rows(i).Cells(GC3.Id).Value = dt.Rows(i)("Id").ToString
            G3.Grid.Rows(i).Cells(GC3.Name).Value = dt.Rows(i)("Name").ToString
            G3.Grid.Rows(i).Cells(GC3.Value).Value = dt.Rows(i)("Value").ToString
            G3.Grid.Rows(i).Cells(GC3.MassValue).Value = dt.Rows(i)("MassValue").ToString
            G3.Grid.Rows(i).Cells(GC3.Notes).Value = dt.Rows(i)("Notes").ToString
        Next
        G3.Grid.RefreshEdit()
    End Sub
    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        bm.NextPrevious(New String() {MainId, SubId}, New String() {SeasonId.Text, txtID.Text}, "Next", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Dim AllowPrint As Boolean = False
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        AllowPrint = False

        If ArName.Text.Trim = "" OrElse Not TestNames() OrElse Not TestDuplicate() Then
            ArName.Focus()
            Return
        End If

        If SearchDate.SelectedDate Is Nothing Or SearchDate.SelectedDate > Now Then
            bm.ShowMSG("برجاء تحديد تاريخ صحيح")
            SearchDate.Focus()
            Return
        End If


        Dim State As BasicMethods.SaveState = BasicMethods.SaveState.Update
        If txtID.Text.Trim = "" Then
            txtID.Text = bm.ExecuteScalar("select max(Id)+1 from " & TableName & " where " & MainId & "=" & SeasonId.Text)
            If txtID.Text = "" Then txtID.Text = "1"
            lblLastEntry.Content = txtID.Text
            'Begin Animation
            State = BasicMethods.SaveState.Insert
        End If



        BasicSalary.Text = Val(BasicSalary.Text)
        CityId.Text = Val(CityId.Text)
        AreaId.Text = Val(AreaId.Text)
        TownId.Text = Val(TownId.Text)
        SubTownId.Text = Val(SubTownId.Text)
        ReligionId.Text = Val(ReligionId.Text)
        JobId.Text = Val(JobId.Text)
        CaseTypeId.Text = Val(CaseTypeId.Text)
        EmpId.Text = Val(EmpId.Text)


        bm.DefineValues()
        If Not bm.Save(New String() {MainId, SubId}, New String() {SeasonId.Text, txtID.Text.Trim}, State) Then
            If State = BasicMethods.SaveState.Insert Then
                txtID.Text = ""
                lblLastEntry.Content = ""
            End If
            Return
        End If

        G1.Grid.EndEdit()
        G2.Grid.EndEdit()
        G3.Grid.EndEdit()
        
        bm.SaveGrid(G1.Grid, "CaseDetails", New String() {"SeasonId", "CaseId"}, New String() {SeasonId.Text, txtID.Text}, New String() {"Id", "Name", "Descrip", "JobId", "BirthDate", "IllTypeId", "ProblemTypeId", "CaseLevelId", "NationalId", "Notes"}, New String() {GC1.Id, GC1.Name, GC1.Descrip, GC1.JobId, GC1.BirthDate, GC1.IllTypeId, GC1.ProblemTypeId, GC1.CaseLevelId, GC1.NationalId, GC1.Notes}, New VariantType() {VariantType.Integer, VariantType.String, VariantType.String, VariantType.String, VariantType.Date, VariantType.Integer, VariantType.Integer, VariantType.Integer, VariantType.String, VariantType.String}, New String() {GC1.Name})

        bm.SaveGrid(G2.Grid, "CaseNeeds", New String() {"SeasonId", "CaseId"}, New String() {SeasonId.Text, txtID.Text}, New String() {"Id", "Name", "Count1", "Count2", "Diff", "NeedPeriodId", "DayDate", "Notes"}, New String() {GC2.Id, GC2.Name, GC2.Count1, GC2.Count2, GC2.Diff, GC2.NeedPeriodId, GC2.DayDate, GC2.Notes}, New VariantType() {VariantType.Integer, VariantType.String, VariantType.Integer, VariantType.Integer, VariantType.Integer, VariantType.String, VariantType.Date, VariantType.String}, New String() {GC2.Id})

        bm.SaveGrid(G3.Grid, "CaseIncomeTypes", New String() {"SeasonId", "CaseId"}, New String() {SeasonId.Text, txtID.Text}, New String() {"Id", "Name", "Value", "MassValue", "Notes"}, New String() {GC3.Id, GC3.Name, GC3.Value, GC3.MassValue, GC3.Notes}, New VariantType() {VariantType.Integer, VariantType.String, VariantType.Decimal, VariantType.String, VariantType.String}, New String() {GC3.Id})

        bm.SaveImage(TableName, New String() {MainId, SubId}, New String() {SeasonId.Text, txtID.Text.Trim}, "Image", Image1)

        AllowPrint = True

        If Not DontClear Then btnNew_Click(sender, e)
        AllowClose = True
    End Sub
    Function TestNames() As Boolean

        ArName.Text = ArName.Text.Trim
        EnName.Text = EnName.Text.Trim
        While ArName.Text.Contains("  ")
            ArName.Text = ArName.Text.Replace("  ", " ")
        End While
        While EnName.Text.Contains("  ")
            EnName.Text = EnName.Text.Replace("  ", " ")
        End While

        Dim Ar() As String
        Ar = ArName.Text.Split(" ")
        Dim En() As String
        En = EnName.Text.Split(" ")
        If Ar.Length <> En.Length Then
            bm.ShowMSG("Arabic Name Length must be EQUALE English Name Length")
            ArName.Focus()
            Return False
        End If

        Dim x As Integer = 0
        For i As Integer = 0 To Ar.Length - 1
            If Ar(i) = En(i) And Not IsNumeric(Ar(i)) Then
                bm.ShowMSG("Arabic Name could not be EQUALE English Name")
                EnName.Select(x, En(i).Length)
                EnName.Focus()
                Return False
            End If
            x += En(i).Length + 1
        Next


        For i As Integer = 0 To Ar.Length - 1
            Dim a As String = bm.ExecuteScalar("delete from Names  where Arabic_Name='" & Ar(i) & "' insert into Names (Arabic_Name,English_Name) values ('" & Ar(i) & "','" & En(i) & "')")
        Next

        Return True
    End Function


    Function TestDuplicate() As Boolean
        Dim s As String = bm.ExecuteScalar("select top 1 ID from " & TableName & " where ID<>'" & txtID.Text & "' and ArName='" & ArName.Text & "' and NationalId='" & NationalId.Text & "' order by ID")

        If s <> "" Then
            bm.ShowMSG("تم تكرار الاسم والرقم القومى مع الحالة رقم " & s)
            Return False
        Else
            Return True
        End If

    End Function


    Private Sub btnFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFirst.Click

        bm.FirstLast(New String() {MainId, SubId}, "Min", dt)
        If dt.Rows.Count = 0 Then Return
        FillControls()
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        bm.ClearControls()
        ClearControls()
    End Sub

    Sub ClearControls()
        TreeView1.Items.Clear()
        bm.ClearControls()
        bm.SetNoImage(Image1, True)

        CityName.Clear()
        AreaName.Clear()
        TownName.Clear()
        SubTownName.Clear()
        ReligionName.Clear()
        LevelName.Clear()
        JobName.Clear()
        CaseTypeName.Clear()
        EmpName.Clear()

        G1.Grid.Rows.Clear()
        G2.Grid.Rows.Clear()
        G3.Grid.Rows.Clear()

        ArName.Clear()
        txtID.Text = bm.ExecuteScalar("select max(" & SubId & ")+1 from " & TableName & " where " & MainId & "=" & SeasonId.Text)
        If txtID.Text = "" Then txtID.Text = "1"

        NewId()
        ArName.Focus()
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If bm.ShowDeleteMSG("هل أنت متأكد من المسح؟") Then
            bm.ExcuteNonQuery("delete from " & TableName & " where " & MainId & "='" & SeasonId.Text.Trim & "' and " & SubId & "='" & txtID.Text.Trim & "'")
            bm.ExcuteNonQuery("delete from CaseDetails where " & MainId & "='" & SeasonId.Text.Trim & "' and CaseId='" & txtID.Text.Trim & "'")
            bm.ExcuteNonQuery("delete from CaseNeeds where " & MainId & "='" & SeasonId.Text.Trim & "' and CaseId='" & txtID.Text.Trim & "'")
            bm.ExcuteNonQuery("delete from CaseAttachments where " & MainId & "='" & SeasonId.Text.Trim & "' and CaseId='" & txtID.Text.Trim & "'")
            bm.ExcuteNonQuery("delete from CaseIncomeTypes where " & MainId & "='" & SeasonId.Text.Trim & "' and CaseId='" & txtID.Text.Trim & "'")

            btnNew_Click(sender, e)
        End If
    End Sub

    Private Sub btnPrevios_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevios.Click
        bm.NextPrevious(New String() {MainId, SubId}, New String() {SeasonId.Text, txtID.Text}, "Back", dt)
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
        bm.RetrieveAll(New String() {MainId, SubId}, New String() {SeasonId.Text, txtID.Text.Trim}, dt)
        If dt.Rows.Count = 0 Then
            ClearControls()
            ArName.Focus()
            lv = False
            Return
        End If
        FillControls()
        lv = False
        ArName.SelectAll()
        ArName.Focus()
        ArName.SelectAll()
        ArName.Focus()
        'arName.Text = dt(0)("Name")
    End Sub



    Private Sub SeasonId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles SeasonId.KeyUp
        If bm.ShowHelp("المواسم", SeasonId, SeasonName, e, "select cast(Id as varchar(100)) Id,Name from Seasons") Then
            btnNew_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub ReligionId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles ReligionId.KeyUp
        bm.ShowHelp("الديانات", ReligionId, ReligionName, e, "select cast(Id as varchar(100)) Id,Name from Religions")
    End Sub

    Private Sub LevelId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles LevelId.KeyUp
        bm.ShowHelp("المستويات", LevelId, LevelName, e, "select cast(Id as varchar(100)) Id,Name from NLevels")
    End Sub

    Private Sub JobId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles JobId.KeyUp
        bm.ShowHelp("الوظائف", JobId, JobName, e, "select cast(Id as varchar(100)) Id,Name from CaseJobs")
    End Sub

    Private Sub CaseTypeId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles CaseTypeId.KeyUp
        bm.ShowHelp("أنواع الحالات", CaseTypeId, CaseTypeName, e, "select cast(Id as varchar(100)) Id,Name from CaseTypes")
    End Sub

    Private Sub EmpId_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles EmpId.KeyUp
        bm.ShowHelp("الموظفين", EmpId, EmpName, e, "select cast(Id as varchar(100)) Id,ArName Name from Employees where Stopped=0")
    End Sub


    Private Sub txtID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles txtID.KeyDown, CityId.KeyDown, AreaId.KeyDown, TownId.KeyDown, SubTownId.KeyDown, SeasonId.KeyDown, NationalId.KeyDown, SearchIndex.KeyDown
        bm.MyKeyPress(sender, e)
    End Sub

    Private Sub txtID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles txtID.KeyUp
        bm.ShowHelpCases(txtID, ArName, e)
    End Sub

    Private Sub txtID_KeyPress2(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs)
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
    Dim lop As Boolean = False
    Private Sub SeasonId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles SeasonId.LostFocus
        If lop Then Return
        lop = True
        If SeasonId.Visibility = Visibility.Visible Then
            bm.LostFocus(SeasonId, SeasonName, "select Name from Seasons where Id=" & SeasonId.Text.Trim())
        End If
        btnNew_Click(Nothing, Nothing)
        lop = False
    End Sub
    Private Sub ReligionId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ReligionId.LostFocus
        bm.LostFocus(ReligionId, ReligionName, "select Name from Religions where Id=" & ReligionId.Text.Trim())
    End Sub
    Private Sub LevelId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles LevelId.LostFocus
        bm.LostFocus(LevelId, LevelName, "select Name from NLevels where Id=" & LevelId.Text.Trim())
    End Sub

    Dim lop4 As Boolean = False
    Private Sub JobId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles JobId.LostFocus
        If lop4 Then Return
        lop4 = True
        bm.LostFocus(JobId, JobName, "select Name from CaseJobs where Id=" & JobId.Text.Trim())
        AddCurrent()
        lop4 = False
    End Sub

    Private Sub CaseTypeId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles CaseTypeId.LostFocus
        bm.LostFocus(CaseTypeId, CaseTypeName, "select Name from CaseTypes where Id=" & CaseTypeId.Text.Trim())
    End Sub

    Private Sub EmpId_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles EmpId.LostFocus
        bm.LostFocus(EmpId, EmpName, "select ArName Name from Employees where Stopped=0 and Id=" & EmpId.Text.Trim())
    End Sub

    Private Sub btnSetImage_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnSetImage.Click
        bm.SetImage(Image1)
    End Sub

    Private Sub btnSetNoImage_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnSetNoImage.Click
        bm.SetNoImage(Image1, True, True)
    End Sub

    Dim lop2 As Boolean = False
    Private Sub ArName_LostFocus(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles ArName.LostFocus
        If lop2 Then Return
        lop2 = True
        ArName.Text = ArName.Text.Trim
        While ArName.Text.Contains("  ")
            ArName.Text = ArName.Text.Replace("  ", " ")
        End While
        Dim s() As String
        s = ArName.Text.Split(" ")
        EnName.Clear()
        For i As Integer = 0 To s.Length - 1
            Dim a As String = bm.ExecuteScalar("select top 1 English_Name from Names where Arabic_Name='" & s(i) & "'")
            If a = "" Then
                EnName.Text &= s(i)
            Else
                EnName.Text &= a
            End If
            EnName.Text &= " "
        Next
        EnName.Text = EnName.Text.Trim
        AddCurrent()
        lop2 = False
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
            F0 = SeasonId.Text
            F1 = txtID.Text
            F2 = CType(TreeView1.SelectedItem, TreeViewItem).Tag
            BackgroundWorker1.RunWorkerAsync()
        Catch ex As Exception
        End Try
    End Sub
    Dim F2 As String = "", F1 As String = "", F0 As String = ""
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Dim myCommand As SqlClient.SqlCommand
            myCommand = New SqlClient.SqlCommand("select Image from CaseAttachments where SeasonId=" & F0 & " and CaseId='" & F1 & "' and AttachedName='" & F2 & "'" & bm.AppendWhere, con)
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
                    bm.ExcuteNonQuery("delete from CaseAttachments where SeasonId=" & SeasonId.Text & " and CaseId='" & txtID.Text & "' and AttachedName='" & TreeView1.SelectedItem.Header & "'" & bm.AppendWhere)
                    LoadTree()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadTree()
        Dim dt As DataTable = bm.ExcuteAdapter("select AttachedName from CaseAttachments where SeasonId=" & SeasonId.Text & " and CaseId=" & txtID.Text & bm.AppendWhere)
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

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        Dim o As New Forms.OpenFileDialog
        o.Multiselect = True
        If o.ShowDialog = Forms.DialogResult.OK Then
            For i As Integer = 0 To o.FileNames.Length - 1
                bm.SaveFile("CaseAttachments", "SeasonId", SeasonId.Text, "CaseId", txtID.Text, "AttachedName", (o.FileNames(i).Split("\"))(o.FileNames(i).Split("\").Length - 1), "Image", o.FileNames(i))
            Next
        End If
        LoadTree()
    End Sub


    Private Sub TreeView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TreeView1.MouseDoubleClick
        Button4_Click(Nothing, Nothing)
    End Sub


    Dim WithEvents d As New Forms.DateTimePicker
    Private Sub G1_EditingControlShowing(ByVal sender As Object, ByVal e As Forms.DataGridViewEditingControlShowingEventArgs)
        e.Control.Controls.Clear()
        If G1.Grid.CurrentCell.ColumnIndex = 4 Then
            d = New Forms.DateTimePicker
            e.Control.Controls.Add(d)
            d.Width = G1.Grid.CurrentCell.OwningColumn.Width
            AddHandler d.Validated, AddressOf d_ValueChanged
            Try
                e.Control.Controls(0).Text = G1.Grid.CurrentCell.Value
            Catch
            End Try
        End If
    End Sub

    Private Sub d_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            G1.Grid.CurrentCell.Value = d.Text.Substring(0, 10)
        Catch ex As Exception
        End Try
    End Sub

    Dim WithEvents d2 As New Forms.DateTimePicker
    Private Sub d2_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            G2.Grid.CurrentCell.Value = d2.Text.Substring(0, 10)
        Catch ex As Exception
        End Try
    End Sub

    Dim lop3 As Boolean = False
    Private Sub G1_CellEndEdit(ByVal sender As Object, ByVal e As Forms.DataGridViewCellEventArgs)
        If lop3 Then Return
        lop3 = True
        Try
            If G1.Grid.CurrentCell.ColumnIndex = 4 Then
                Try
                    Dim dd As DateTime = DateTime.Parse(G1.Grid.CurrentCell.Value)
                Catch ex As Exception
                    G1.Grid.CurrentCell.Value = ""
                End Try
            End If
            G1.Grid.Rows(G1.Grid.CurrentRow.Index).Cells(0).Value = G1.Grid.CurrentRow.Index + 1
            If G1.Grid.CurrentRow.Index = 0 Then
                AddCurrent(True)
            End If
        Catch ex As Exception
        End Try
        lop3 = False
    End Sub

    Private Sub G1_UserDeletedRow(ByVal sender As Object, ByVal e As Forms.DataGridViewRowEventArgs)
        For i As Integer = 0 To G1.Grid.Rows.Count - 1
            G1.Grid.Rows(i).Cells(0).Value = i + 1
        Next
    End Sub

    Private Sub G2_CellEndEdit(ByVal sender As Object, ByVal e As Forms.DataGridViewCellEventArgs)
        If G2.Grid.Columns(e.ColumnIndex).Name = GC2.Id Then
            AddItem(G2.Grid.Rows(e.RowIndex).Cells(GC2.Id).Value, e.RowIndex, 0)
        End If
        G2.Grid.Rows(e.RowIndex).Cells(GC2.Count1).Value = Val(G2.Grid.Rows(e.RowIndex).Cells(GC2.Count1).Value)
        G2.Grid.Rows(e.RowIndex).Cells(GC2.Count2).Value = Val(G2.Grid.Rows(e.RowIndex).Cells(GC2.Count2).Value)
        G2.Grid.Rows(e.RowIndex).Cells(GC2.Diff).Value = Val(G2.Grid.Rows(e.RowIndex).Cells(GC2.Count1).Value) - Val(G2.Grid.Rows(e.RowIndex).Cells(GC2.Count2).Value)
    End Sub

    Private Sub G2_KeyDown(ByVal sender As Object, ByVal e As Forms.KeyEventArgs)
        If bm.ShowHelpGrid("الاحتياجات", G2.Grid.Rows(G2.Grid.CurrentRow.Index).Cells(GC2.Id), G2.Grid.Rows(G2.Grid.CurrentRow.Index).Cells(GC2.Name), e, "select cast(Id as varchar(100)) Id,Name from Needs") Then
            G2.Grid.CurrentCell = G2.Grid.Rows(G2.Grid.CurrentRow.Index).Cells(1)
        End If
    End Sub

    Sub AddItem(ByVal Id As String, Optional ByVal i As Integer = -1, Optional ByVal Add As Decimal = 1)
        Try
            Dim Exists As Boolean = False
            Dim Move As Boolean = False
            If i = -1 Then Move = True

            G2.Grid.AutoSizeColumnsMode = Forms.DataGridViewAutoSizeColumnsMode.Fill
            If i = -1 Then
                For x As Integer = 0 To G2.Grid.Rows.Count - 1
                    If Not G2.Grid.Rows(x).Cells(GC2.Id).Value Is Nothing AndAlso G2.Grid.Rows(x).Cells(GC2.Id).Value.ToString = Id.ToString Then
                        i = x
                        Exists = True
                        GoTo Br
                    End If
                Next
                i = G2.Grid.Rows.Add()
Br:
            End If

            Dim dt As DataTable = bm.ExcuteAdapter("Select * From Needs where Id='" & Id & "'")
            Dim dr() As DataRow = dt.Select("Id='" & Id & "'")
            If dr.Length = 0 Then
                If Not G2.Grid.Rows(i).Cells(GC2.Id).Value Is Nothing Or G2.Grid.Rows(i).Cells(GC2.Id).Value <> "" Then bm.ShowMSG("هذا البند غير موجود")
                G2.Grid.Rows(i).Cells(GC2.Id).Value = ""
                G2.Grid.Rows(i).Cells(GC2.Name).Value = ""
                Return
            End If
            G2.Grid.Rows(i).Cells(GC2.Id).Value = dr(0)(GC2.Id)
            G2.Grid.Rows(i).Cells(GC2.Name).Value = dr(0)(GC2.Name)
            
        Catch
        End Try
    End Sub

    Sub AddCurrent(Optional ByVal FromGrid As Boolean = False)
        If G1.Grid.Rows.Count <= 1 Then
            G1.Grid.Rows.Add()
        End If
        If FromGrid Then
            ArName.Text = G1.Grid.Rows(0).Cells(GC1.Name).Value
            NationalId.Text = G1.Grid.Rows(0).Cells(GC1.NationalId).Value
            JobId.Text = G1.Grid.Rows(0).Cells(GC1.JobId).Value
            DateOfBirth.SelectedDate = DateTime.Parse(G1.Grid.Rows(0).Cells(GC1.BirthDate).Value)
        End If
        G1.Grid.Rows(0).Cells(GC1.Name).Value = ArName.Text
        G1.Grid.Rows(0).Cells(GC1.NationalId).Value = NationalId.Text
        G1.Grid.Rows(0).Cells(GC1.JobId).Value = JobId.Text
        If Not DateOfBirth.SelectedDate Is Nothing Then G1.Grid.Rows(0).Cells(GC1.BirthDate).Value = DateOfBirth.SelectedDate.Value.ToShortDateString

        If Val(G1.Grid.Rows(0).Cells(GC1.Descrip).Value) < 1 Then
            G1.Grid.Rows(0).Cells(GC1.Descrip).Value = "1"
        End If
        G1.Grid.CurrentCell = G1.Grid.Rows(0).Cells(1)
        G1_CellEndEdit(G1.Grid, New Forms.DataGridViewCellEventArgs(1, 0))
        ArName_LostFocus(ArName, Nothing)
        JobId_LostFocus(ArName, Nothing)
    End Sub

    Private Sub NationalId_LostFocus(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles NationalId.LostFocus
        AddCurrent()
    End Sub

    Private Sub G3_CellEndEdit(ByVal sender As Object, ByVal e As Forms.DataGridViewCellEventArgs)
        Try
            If G3.Grid.CurrentCell.ColumnIndex <> 0 Then Return
            Dim i As Integer = G3.Grid.CurrentRow.Index
            Dim id As String = G3.Grid.Rows(i).Cells(GC3.Id).Value
            Dim dt As DataTable = bm.ExcuteAdapter("Select * From IncomeTypes where Id='" & id & "'")
            Dim dr() As DataRow = dt.Select("Id='" & Id & "'")
            If dr.Length = 0 Then
                If Not G3.Grid.Rows(i).Cells(GC3.Id).Value Is Nothing AndAlso G3.Grid.Rows(i).Cells(GC3.Id).Value <> "" Then bm.ShowMSG("هذا الكود غير موجود")
                G3.Grid.Rows(i).Cells(GC3.Id).Value = ""
                G3.Grid.Rows(i).Cells(GC3.Name).Value = ""
                Return
            End If
            G3.Grid.Rows(i).Cells(GC3.Id).Value = dr(0)(GC3.Id)
            G3.Grid.Rows(i).Cells(GC3.Name).Value = dr(0)(GC3.Name)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub G3_KeyDown(ByVal sender As Object, ByVal e As Forms.KeyEventArgs)
        If bm.ShowHelpGrid("أنواع الدخل", G3.Grid.Rows(G3.Grid.CurrentRow.Index).Cells(GC3.Id), G3.Grid.Rows(G3.Grid.CurrentRow.Index).Cells(GC3.Name), e, "select cast(Id as varchar(100)) Id,Name from IncomeTypes") Then
            G3.Grid.CurrentCell = G3.Grid.Rows(G3.Grid.CurrentRow.Index).Cells(1)
        End If
    End Sub

    Private Sub G2_EditingControlShowing(ByVal sender As Object, ByVal e As Forms.DataGridViewEditingControlShowingEventArgs)
        
        e.Control.Controls.Clear()
        If G2.Grid.CurrentCell.ColumnIndex = 0 Or G2.Grid.CurrentCell.ColumnIndex = 1 Then
            Dim c = e.Control
            RemoveHandler c.KeyDown, AddressOf G2_KeyDown
            AddHandler c.KeyDown, AddressOf G2_KeyDown
        ElseIf G2.Grid.CurrentCell.ColumnIndex = 6 Then
            d2 = New Forms.DateTimePicker
            e.Control.Controls.Add(d2)
            d2.Width = G2.Grid.CurrentCell.OwningColumn.Width
            AddHandler d2.Validated, AddressOf d2_ValueChanged
            Try
                e.Control.Controls(0).Text = G2.Grid.CurrentCell.Value
            Catch
            End Try
        End If

    End Sub

    Private Sub G3_EditingControlShowing(ByVal sender As Object, ByVal e As Forms.DataGridViewEditingControlShowingEventArgs)
        Dim c = e.Control
        RemoveHandler c.KeyDown, AddressOf G3_KeyDown
        AddHandler c.KeyDown, AddressOf G3_KeyDown
    End Sub

    Private Sub DateOfBirth_LostFocus(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles DateOfBirth.LostFocus
        AddCurrent()
    End Sub

    Dim DontClear As Boolean = False
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnPrint.Click
        DontClear = True
        btnSave_Click(Nothing, Nothing)
        DontClear = False
        If Not AllowPrint Then Return

        Dim rpt As New ReportViewer
        rpt.paraname = {"Header", "@SeasonId", "@Id"}
        rpt.paravalue = {"بيانات حالة", SeasonId.Text, txtID.Text}
        rpt.Header = Md.MyProject.ToString
        rpt.RptPath = "Case.rpt"
        rpt.ShowDialog()
    End Sub
End Class
