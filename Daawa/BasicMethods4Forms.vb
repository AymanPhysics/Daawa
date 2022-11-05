Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.Management
Imports System.Text
Imports System.Security.Cryptography
Imports System.Net.Mail
Imports System.Net

Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Windows.Controls.Primitives

'Imports System.Runtime.integereropServices

Public Class BasicMethods4Forms
    '''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Length As Integer = 0
    Public stat As String = ""
    ' Dim MyCmd = New SqlCommand
    Dim da As New SqlDataAdapter()
    Dim TableObject As String = "T"
    Public HelpProc As String = "Help"
    Public FirstColName As String = "الكــود"
    Public SecondColName As String = "الاســــم"
    Public ThirdColName As String = ""

    Public ColName1 As String = "الرقـــم"
    Public SelectRetrieve As String = ""
    Public AppendWhere As String = "", TableWhere As String = ""
    Public DiscountKeyFiels As Integer
    Public Result() As String = {}
    Public DeleteMsg As String = "هل أنت متأكد من الDelete"

    Public Table_Name As String = ""
    Public Fields() As String = {}
    Public GeneralString As String = ""
    Public KeyFields() As String = {}
    Public control() As Control = {}
    Public ReturnedValues(,) As String = {}
    Public Values() As String = {}
    Public returne As Boolean = False






    Public Structure SYSTEMTIME
        Public wYear As UInt16
        Public wMonth As UInt16
        Public wDayOfWeek As UInt16
        Public wDay As UInt16
        Public wHour As UInt16
        Public wMinute As UInt16
        Public wSecond As UInt16
        Public wMilliseconds As UInt16
    End Structure

    Public Sub SetTime()
        Try
            Dim dd As DateTime = CType(ExecuteScalar("select getdate()"), DateTime)
            Today = dd
            TimeOfDay = dd
        Catch ex As Exception
            ShowMSG("Please, Run " & Md.MyProject.ToString & " As Administrator" & vbCrLf & "to enable Time Updating")
        End Try
    End Sub



    Public Function Physics_processorId() As String
        Dim s As String = ""
        Dim search As New ManagementObjectSearcher(New SelectQuery("Win32_processor"))
        For Each info As ManagementObject In search.Get()
            Try
                s &= info("processorId").ToString()
            Catch ex As Exception
            End Try
        Next
        Return s.ToUpper
    End Function

    Public Function EnName(ByVal ArName As String, ByVal dtName As DataTable) As String
        ArName = ArName.Trim
        While ArName.Contains("  ")
            ArName = ArName.Replace("  ", " ")
        End While
        Dim s() As String
        s = ArName.Split(" ")
        Dim En As String = ""
        For i As Integer = 0 To s.Length - 1

            'Dim a As String = ExecuteScalar("select top 1 English_Name from Names where Arabic_Name='" & s(i) & "'")
            Dim a As String = ""
            Try
                a = (dtName.Select("Arabic_Name='" & s(i) & "'")(0))(1)
            Catch ex As Exception
                a = s(i)
            End Try
            If a = "" Then
                En &= s(i)
            Else
                En &= a
            End If
            En &= " "
        Next
        Return En.Trim
    End Function


    Public Function Physics_SerialNumber() As String
        Dim s As String = ""
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMedia")
        For Each wmi_HD As ManagementObject In searcher.Get()
            If wmi_HD("SerialNumber") <> Nothing Then
                s &= wmi_HD("SerialNumber").ToString()
            End If
        Next
        Return s.ToUpper
    End Function


    Public Function Physics_BaseBoard() As String
        Dim s As String = ""
        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")
        For Each wmi_HD As ManagementObject In searcher.Get()
            If wmi_HD("SerialNumber") <> Nothing Then
                s &= wmi_HD.Properties("SerialNumber").ToString()
            End If
            If wmi_HD("Product") <> Nothing Then
                s &= wmi_HD.Properties("Product").Value.ToString().Trim()
            End If
        Next

        searcher.Dispose()

        Return s.ToUpper
    End Function

    Public Function Physics_VolumeSerialNumber(ByVal Volume As String) As String
        Dim s As String = ""
        Dim disk As ManagementObject = New ManagementObject(String.Format("win32_logicaldisk.deviceid=""{0}:""", Volume))
        Try
            disk.Get()
        Catch ex As Exception
        End Try
        Try
            s &= disk("VolumeSerialNumber").ToString()
        Catch ex As Exception
        End Try
        Return s
    End Function

    Public Function Physics_MACAddress() As String
        Dim s As String = ""
        Dim mo2 As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_NetworkAdapterConfiguration")
        For Each mac As ManagementObject In mo2.Get()
            Try
                s &= mac("MACAddress").ToString.Replace(":", "")
            Catch ex As Exception
            End Try
            If s.Length > 0 Then Exit For
        Next
        Return s
    End Function

    Public Function Physics_MacAddress2() As String
        Dim s As String = ""
        Dim mc As ManagementClass = New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        Dim MACAddress As String = ""
        For Each mo As ManagementObject In moc
            If (MACAddress.Equals(String.Empty)) Then
                If CBool(mo("IPEnabled")) Then MACAddress = mo("MacAddress").ToString()
                mo.Dispose()
            End If
            Try
                s &= MACAddress.Replace(":", String.Empty)
            Catch ex As Exception
            End Try
            If s.Length > 0 Then Exit For
        Next
        Return s
    End Function

    Public Function Protect(ByVal s As String) As String
        Return s.Replace(" ", "").Replace("0", "P").Replace("1", "K").Replace("2", "U").Replace("3", "G").Replace("4", "R").Replace("5", "D").Replace("6", "W").Replace("7", "Q").Replace("8", "A").Replace("9", "Z").ToUpper()
    End Function



    Enum CloseState
        Yes
        No
        Cancel
    End Enum






    Public Sub TestCountry(ByVal Cbo As ComboBox)
        If Md.SystemAdmin = Md.UserName Or Md.GeneralManager = "1" Or Md.Board = "1" Then Return
        Cbo.SelectedValue = Md.CountryId 'MyChildCountry
        Cbo.Enabled = False
    End Sub

    Public Sub TestCity(ByVal Cbo As ComboBox)
        If Md.SystemAdmin = Md.UserName Or Md.GeneralManager = "1" Or Md.Board = "1" Then Return
        Cbo.SelectedValue = Md.CityId 'MyChildCity
        Cbo.Enabled = False
    End Sub

    Public Sub TestEmployee(ByVal Cbo As ComboBox)
        If Md.SystemAdmin = Md.UserName Or Md.GeneralManager = "1" Or Md.Board = "1" Then Return
        Cbo.SelectedValue = Md.UserName
        Cbo.Enabled = False
    End Sub



    Sub HandleKeyDown(ByVal frm As Form)
        For Each C As Control In frm.Controls
            If C.GetType Is GetType(TextBox) Then
                Dim cc As TextBox = C
                If cc.Multiline Then
                    Continue For
                End If
            End If
            AddHandler C.KeyPress, AddressOf _KeyPress
        Next
    End Sub


    Public Sub _KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Asc(e.KeyChar) = 13 Then e.Handled = True
    End Sub
    Public Sub DefineValues()
        ReDim Values(control.Length)
        Dim val As String = "", type As String = "", nn As String = ""
        For i As Integer = 0 To control.Length - 1
            val = control(i).Text.Trim()
            type = control(i).GetType().ToString().Trim
            nn = control(i).Name
            If ((type = "AMS.TextBox.integeregerTextBox" Or type = "AMS.TextBox.NumericTextBox" Or type = "ACCOUNTING.DecimalTextbox" Or type = "System.Windows.Forms.NumericUpDown") And val = "") Then
                Values(i) = "0"
            ElseIf ((type = "AMS.TextBox.integeregerTextBox" Or type = "AMS.TextBox.NumericTextBox" Or type = "ACCOUNTING.DecimalTextbox" Or type = "System.Windows.Forms.NumericUpDown") And val <> "") Then
                Values(i) = val

            ElseIf (type = "System.Windows.Forms.ComboBox") Then

                Dim cbo As New ComboBox
                cbo = control(i)
                If cbo.SelectedValue Is Nothing Then
                    '   Values(i) = DBNull.Value
                    Values(i) = cbo.SelectedIndex.ToString().Trim

                Else
                    Values(i) = cbo.SelectedValue.ToString().Trim

                End If

            ElseIf (type = "System.Windows.Forms.CheckBox") Then
                Dim chk As New CheckBox()
                chk = control(i)
                If (chk.Checked) Then
                    Values(i) = "1"

                Else
                    Values(i) = "0"
                End If

            ElseIf (type = "System.Windows.Forms.RadioButton") Then
                Dim rbd As New RadioButton()
                rbd = control(i)
                If (rbd.Checked) Then
                    Values(i) = "1"
                Else
                    Values(i) = "0"
                End If

            ElseIf type = "System.Windows.Forms.DateTimePicker" Then

                Dim dd As DateTimePicker = control(i)
                Values(i) = ToStrDate(dd.Value)

            ElseIf type = "System.Windows.Forms.TextBox" Then
                Dim txt As TextBox = control(i)
                If txt.PasswordChar <> Nothing Or (Table_Name = "PCs" And txt.Name = "txtName") Then
                    Values(i) = Encrypt(val)
                Else
                    Values(i) = val
                End If
            Else
                Values(i) = val
            End If


            Try
                Values(i) = Values(i).Replace("'", "''")
            Catch
            End Try
        Next

    End Sub

    Enum SaveState
        All
        Insert
        Update
    End Enum

    Public Function Save(ByVal ID() As String, ByVal IDValue() As String, Optional ByVal State As SaveState = SaveState.All) As Boolean
        ' DefineValues()
        If Not StopPro() Then Return False
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            'main.sqlConnection1.Open()
            MyCmd.CommandType = CommandType.Text
            MyCmd.Parameters.Clear()
            MyCmd.CommandText = ""
            If State <> SaveState.Insert Then
                MyCmd.CommandText = "IF Exists(Select *  From " & Table_Name & " T Where " & ID(0) & "='" & IDValue(0) & "'"
                For i As Integer = 1 To ID.Length - 1
                    MyCmd.CommandText &= " and " & ID(i) & "='" & IDValue(i) & "'"
                Next

                MyCmd.CommandText &= ")"
                MyCmd.CommandText &= " Update " & Table_Name & "  Set "

                For i As Integer = 0 To Fields.Length - 1
                    MyCmd.CommandText &= Fields(i) & "='" & Values(i)
                    If Not i = Fields.Length - 1 Then
                        MyCmd.CommandText &= "', "
                    Else : MyCmd.CommandText &= "'"
                    End If
                Next
                MyCmd.CommandText &= " Where " & ID(0) & "='" & IDValue(0) & "'"

                For i As Integer = 1 To ID.Length - 1
                    MyCmd.CommandText &= " and " & ID(i) & "='" & IDValue(i) & "'"
                Next
            End If

            If State = SaveState.All Then
                MyCmd.CommandText &= " Else "
            End If

            If State <> SaveState.Update Then
                MyCmd.CommandText &= "Insert into " & Table_Name & "(" & Table_Fields() & " ) Values ('"
                For i As Integer = 0 To Fields.Length - 1

                    MyCmd.CommandText &= Values(i)
                    If Not i = Fields.Length - 1 Then
                        MyCmd.CommandText &= "', '"
                    Else : MyCmd.CommandText &= "')"
                    End If
                Next
            End If

            MyCmd.ExecuteNonQuery()
            MyCmd.Connection.Close()
            Return True

        Catch ee As Exception

            Dim ss As String = ee.Message.ToString().Trim
            ShowMSG(ss)
            Return False

        Finally
            c = Nothing
        End Try
    End Function


    Public Function SaveGrid(ByVal Grid As DataGridView, ByVal TableName As String, ByVal KeysFields() As String, ByVal KeysValues() As String, ByVal Fields() As String, ByVal Indexes() As Integer, ByVal Type() As VariantType) As String
        'SaveGrid(New DataGridView, "", New Integer() {}, New VariantType() {VariantType.String})
        Dim Statement As String = "Delete From " & TableName & " where 1=1" & vbCrLf
        For i As Integer = 0 To KeysFields.Length - 1
            Statement &= " and " & KeysFields(i) & "='" & KeysValues(i) & "'" & vbCrLf
        Next

        Statement &= " Insert " & TableName & "("
        For i As Integer = 0 To KeysFields.Length - 1
            Statement &= KeysFields(i) & ","
        Next
        For i As Integer = 0 To Fields.Length - 1
            Statement &= Fields(i) & ","
        Next
        Statement = Mid(Statement, 1, Len(Statement) - 1)
        Statement &= ") values " & vbCrLf

        For i As Integer = 0 To Grid.Rows.Count - 1 - IIf(Grid.AllowUserToAddRows, 1, 0)
            Statement &= IIf(i = 0, "", ",")
            Statement &= "("
            For x As Integer = 0 To KeysValues.Length - 1
                Statement &= KeysValues(x) & ","
            Next
            For x As Integer = 0 To Indexes.Length - 1
                Statement &= GetCellValue(Type(x), Grid.Rows(i).Cells(Indexes(x)).Value)
                Statement &= IIf(x = Indexes.Length - 1, "", ",")
            Next
            Statement &= ")"
        Next

        Return Statement
    End Function

    Function GetCellValue(ByVal Type As VariantType, ByVal Value As String) As String
        Select Case Type
            Case VariantType.Date
                Return "'" & ToStrDate(DateTime.Parse(Value)) & "'"
            Case VariantType.Integer
                Return Value
            Case VariantType.Decimal
                Return Value
            Case VariantType.Boolean
                Return IIf(Value, 1, 0)
            Case VariantType.String
                Return "'" & Value & "'"
            Case Else
                Return "''"
        End Select
    End Function

    '___________________________Check IF Record Whith Condition is Exist__________________
    Public Function IF_Exists() As Boolean

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim dt As New DataTable
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            MyCmd.CommandText = ""
            MyCmd.CommandText = " IF Exists ( Select * From " & Table_Name & " " & WhereKeyFields() & " ) select '1' else select '0'"

            da.SelectCommand = MyCmd
            dt.Clear()
            da.Fill(dt)

            If dt.Rows(0)(0).ToString().Trim = "1" Then
                Return True
            Else : Return False
            End If

        Catch
            Return False
        End Try
    End Function
    Public Function IF_Exists(ByVal SqlStatment As String) As Boolean

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim dt As New DataTable
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = ""
            MyCmd.CommandText = " IF Exists (" & SqlStatment & ") select '1' else select '0'"

            da.SelectCommand = MyCmd
            dt.Clear()
            da.Fill(dt)

            If dt.Rows(0)(0).ToString().Trim = "1" Then

                Return True
            Else : Return False
            End If

        Catch
            Return False
        Finally
            c = Nothing
        End Try
    End Function

    Public Sub FillControls()

        Try

            Dim Type As String = ""
            For i As Integer = 0 To Fields.Length - 1
                Type = control(i).GetType().ToString().Trim
                Dim nn As String = control(i).Name
                If Type = "System.Windows.Forms.ComboBox" Then
                    Dim cbo As ComboBox = control(i)
                    If (cbo.ValueMember.Length < 2) Then    'for save integereger no. in DB. from combox
                        Dim index As String = ReturnedValues(0, i)
                        If Not (ReturnedValues(0, i).Trim() = "") Then
                            Try
                                cbo.SelectedIndex = Integer.Parse(ReturnedValues(0, i))
                            Catch ex As Exception
                            End Try
                        End If
                    Else
                        Try
                            cbo.SelectedValue = ReturnedValues(0, i)
                        Catch ex As Exception
                            cbo.SelectedIndex = -1
                        End Try

                    End If
                ElseIf (Type = "System.Windows.Forms.CheckBox") Then

                    Dim chk As CheckBox = control(i)
                    Dim val As String = ReturnedValues(0, i)
                    If (ReturnedValues(0, i) = "1" Or ReturnedValues(0, i) = "True") Then
                        chk.Checked = True
                    Else : chk.Checked = False
                    End If
                ElseIf (Type = "System.Windows.Forms.RadioButton") Then
                    Dim rbt As RadioButton = control(i)
                    If (ReturnedValues(0, i) = "1") Then
                        rbt.Checked = True

                    Else : rbt.Checked = False

                    End If
                ElseIf (Type = "System.Windows.Forms.TextBox") Then
                    Dim txt As TextBox = control(i)
                    If txt.PasswordChar <> Nothing Or (Table_Name = "PCs" And txt.Name = "txtName") Then
                        control(i).Text = Decrypt(ReturnedValues(0, i))
                    Else
                        control(i).Text = ReturnedValues(0, i)
                    End If
                Else : control(i).Text = ReturnedValues(0, i)
                End If
            Next

        Catch ex As Exception
        End Try
    End Sub
    Public Sub RetrieveAll(ByVal ID() As String, ByVal IDValue() As String, ByRef dt As DataTable)   ' k is the control focus if not exists

        '			if(!TestEmpt())
        '				return

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            dt.Clear()
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            MyCmd.Parameters.Clear()
            MyCmd.CommandText = ""
            MyCmd.CommandType = CommandType.Text
            If (SelectRetrieve = "") Then
                MyCmd.CommandText &= "Select " & Table_Fields() & " From " & Table_Name & " T "
            Else : MyCmd.CommandText &= SelectRetrieve
            End If
            MyCmd.CommandText &= " Where " & ID(0) & "='" & IDValue(0) & "'" & AppendWhere

            For i As Integer = 1 To ID.Length - 1
                MyCmd.CommandText &= " and " & ID(i) & "='" & IDValue(i) & "'"
            Next

            da.SelectCommand = MyCmd
            dt.Clear()
            dt = New DataTable
            da.Fill(dt)

            FillValues(dt)

        Catch ex As Exception
            Dim ss As String = ex.Message.ToString().Trim
            ShowMSG(ex.ToString().Trim)
        Finally
            c = Nothing
        End Try

    End Sub



    '___________________________________________________________________________________
    '___________________________Fill Retuened Values ___________________________________
    Public Sub FillValues(ByVal dt As DataTable)

        'Dim ReturnedValues(0, Fields.Length) As String
        If (dt.Rows.Count = 0) Then
            Return
        End If

        ReDim ReturnedValues(dt.Rows.Count, Fields.Length)
        For j As Integer = 0 To dt.Rows.Count - 1
            For i As Integer = 0 To Fields.Length - 1
                ReturnedValues(j, i) = dt.Rows(j)(i).ToString().Trim()
            Next
        Next
    End Sub             '___________________________________________________________________________________
    '---------------------------Delete ON Condition-------------------------------------
    'Public Sub Delete(ByVal Msg As Boolean)

    '    If (Msg) Then

    '        returne = True
    '        Dim mm = New Message(DeleteMsg, True)
    '        mm.ShowDialog()
    '        If Not mm.Ok Then
    '            returne = False
    '            Return
    '        End If
    '    End If
    '    DefineValues()
    '    Try
    '                                                   ()
    '        MyCmd.CommandText = ""
    '        MyCmd.CommandText &= " Delete From " & Table_Name & " " & WhereKeyFields()
    '        MyCmd.ExecuteNonQuery()

    '    Catch ex As Exception
    '    End Try
    'End Sub
    'Public Sub Delete(ByVal Msg As Boolean, ByVal ID As String, ByVal IDValue As String)
    '    If (Msg) Then
    '        returne = True
    '        dim mm as New Message(DeleteMsg, True))
    '        mm.ShowDialog()
    '        If Not mm.Ok Then
    '            returne = False
    '            Return
    '        End If
    '    End If
    '    DefineValues()
    '    Try
    '        MyCmd.CommandText = ""
    '        MyCmd.CommandText &= " Delete From " & Table_Name & " Where " & ID & "=" & IDValue & TableWhere

    '        MyCmd.ExecuteNonQuery()
    '    Catch ex As Exception
    '    End Try
    'End Sub

    'public sub Delete2(bool Msg,string ID,string IDValue,string ID2,string IDValue2)
    '{
    '      If (Msg) Then
    '	{
    '		Return=true
    '          Message(mm = New Message(DeleteMsg, True))
    '          mm.ShowDialog()
    '          If (!mm.Ok) Then
    '		{
    '			Return=false
    '              Return
    '		}

    '	}
    '              DefineValues()
    '              Try
    '	{				
    '                                                                                              ()
    '                  MyCmd.CommandText = ""
    '                  MyCmd.CommandText &= " Delete From " & Table_Name & " Where " & ID & "='" & IDValue & "' and " & ID2 & "='" & IDValue2 & "' "

    '                  If (Global.sqlConnection1.State = ConnectionState.Closed) Then
    '                      Global.sqlConnection1.Open()
    '                      MyCmd.ExecuteNonQuery()
    '	}
    '	catch{}

    '}
    '                  '-----------------------------------------------------------------------------------
    '-------------------------Retrieve First--------------------------------------------
    Public Sub FirstLast(ByVal ID() As String, ByVal MaxOrMin As String, ByRef dt As DataTable)
        DefineValues()
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            dt.Clear()
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            MyCmd.CommandType = CommandType.Text
            MyCmd.Parameters.Clear()
            MyCmd.CommandText = ""
            If (SelectRetrieve = "") Then
                MyCmd.CommandText &= "Select " & Table_Fields() & " From " & Table_Name & " T "
            Else : MyCmd.CommandText &= SelectRetrieve
            End If

            Dim MyKeys As String = ""

            DiscountKeyFiels = 1
            MyCmd.CommandText &= WhereKeyFields()

            Dim KeyField As String = ID(KeyFields.Length - 1).Replace("T.", "") 'KeyFields[KeyFields.Length-1]
            MyCmd.CommandText &= " " & AppendWhere & " and " & KeyField & " =(select " & MaxOrMin & "(" & KeyField & ") From " & Table_Name & " " & WhereKeyFields() & " " & AppendWhere & " ) "

            DiscountKeyFiels = 0
            da.SelectCommand = MyCmd
            dt.Clear()
            dt = New DataTable
            da.Fill(dt)
            FillValues(dt)

        Catch ex As Exception
            Dim ss As String = ex.Message.ToString().Trim
        Finally
            c = Nothing
        End Try
    End Sub

    Public Sub NextPrevious(ByVal ID() As String, ByVal Value() As String, ByVal NextOrBack As String, ByRef dt As DataTable)
        DefineValues()
        dt.Clear()
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Dim MyCmd As SqlCommand = c.CreateCommand()
        If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

        MyCmd.Parameters.Clear()
        MyCmd.CommandType = CommandType.Text
        Dim MaxOrMin As String = "min"
        Dim Compare As String = ">"
        If (NextOrBack = "Next") Or (NextOrBack = "next") Then

            MaxOrMin = "min"
            Compare = " > "
        ElseIf (NextOrBack = "Back") Or (NextOrBack = "back") Then
            MaxOrMin = "max"
            Compare = " < "

        Else : Return
        End If
        Try

            dt.Clear()
            MyCmd.CommandText = ""
            If (SelectRetrieve = "") Then
                MyCmd.CommandText &= "Select " & Table_Fields() & " From " & Table_Name & " T "
            Else : MyCmd.CommandText &= SelectRetrieve
            End If
            Dim KeyField As String = ID(KeyFields.Length - 1)
            KeyField = KeyField.Replace("T.", "")

            DiscountKeyFiels = 1
            MyCmd.CommandText &= WhereKeyFields() & AppendWhere

            MyCmd.CommandText &= " and " & KeyFields(KeyFields.Length - 1) & "=("
            MyCmd.CommandText &= "Select " & MaxOrMin & "(" & KeyField & ") From " & Table_Name
            MyCmd.CommandText &= WhereKeyFields()
            DiscountKeyFiels = 0

            If (Value(KeyFields.Length - 1) = "") Then
                Compare = " > "
            End If
            MyCmd.CommandText &= " and " & KeyFields(KeyFields.Length - 1) & Compare & "'" & Value(KeyFields.Length - 1) & "' " & AppendWhere & ")"

            da.SelectCommand = MyCmd
            dt.Clear()
            da.Fill(dt)
            'main.sqlConnection1.Close()
            If (dt.Rows.Count = 0) Then
                Return
            End If
            FillValues(dt)

        Catch ex As Exception
        Finally
            c = Nothing
        End Try
    End Sub

    Public Sub ClearControls()

        Dim type As String = ""
        For i As Integer = 0 To control.Length - 1
            Dim s As String = control(i).Name

            For n As Integer = 0 To KeyFields.Length - 1
                If KeyFields(n) = Fields(i) Then
                    GoTo A
                End If
            Next

            type = control(i).GetType().ToString().Trim
            If (type = "System.Windows.Forms.ComboBox") Then

                Dim cbo As ComboBox = control(i)
                Try
                    cbo.SelectedIndex = 0
                    cbo.SelectedValue = 0
                Catch
                End Try
            ElseIf (type = "System.Windows.Forms.CheckBox") Then

                Dim chk As CheckBox = control(i)
                chk.Checked = False

            ElseIf (type = "System.Windows.Forms.PictureBox") Then

                Dim chk As PictureBox = control(i)
                chk.Image = Image.FromFile(Application.Current.StartupUri.ToString & "\Images\no.photo.gif")

            ElseIf (type = "System.Windows.Forms.RadioButton") Then

                Dim chk As RadioButton = control(i)
                chk.Checked = False

            Else : control(i).Text = ""
            End If
A:
        Next
        control(0).Focus()
    End Sub

    Public Sub ClearControls(ByVal Focus As Boolean)

        Dim type As String = ""
        For i As Integer = 0 To control.Length - 1

            type = control(i).GetType().ToString().Trim
            If (type = "System.Windows.Forms.ComboBox") Then

                Dim cbo As ComboBox = control(i)
                cbo.SelectedValue = 0

            Else : control(i).Text = ""
            End If
        Next

    End Sub
    '----------------------Select Items ON Conditions-----------------------------------
    Private Function WhereKeyFields() As String

        GeneralString = ""
        GeneralString &= " Where 1=1 "
        For j As Integer = 0 To KeyFields.Length - DiscountKeyFiels - 1
            GeneralString &= " and " & KeyFields(j) & "='"
            For i As Integer = 0 To Fields.Length - 1
                If Fields(i) = KeyFields(j) Then
                    GeneralString &= Values(i)
                End If
            Next
            GeneralString &= "' "
        Next
        Return GeneralString
    End Function


    Private Function WhereLastKeyFields() As String

        GeneralString = ""

        GeneralString &= KeyFields(KeyFields.Length - 1) & "='"
        For i As Integer = 0 To Fields.Length - 1
            If Fields(i) = KeyFields(KeyFields.Length - 1) Then
                GeneralString &= Values(i)
            End If
        Next
        GeneralString &= "'"

        Return GeneralString
    End Function


    '----------------------------------------------------------------------------------
    ' still not used
    Private Function ValueFields() As String
        GeneralString = ""
        For k As Integer = 0 To KeyFields.Length - 1
            GeneralString &= Values(k)
            If (k = KeyFields.Length - 2) Then
                GeneralString &= ""
            Else : GeneralString &= "+"
            End If
        Next
        Return GeneralString
    End Function
    '--------------------------Select Items From Table---------------------------------
    Private Function Table_Fields() As String

        GeneralString = ""
        'GeneralString&="Select "
        For i As Integer = 0 To Fields.Length - 1

            GeneralString &= Fields(i)
            If Not i = Fields.Length - 1 Then
                GeneralString &= " , "
            End If
        Next
        Return GeneralString
    End Function

    Public Sub ShowDialogMessage(ByVal msg As String, ByVal txt As String)
        'Dim mm As New Message(msg, False)
        'mm.Text = txt
        'mm.ShowDialog()
    End Sub
    '  Public Sub ShowDialogHelp(ByVal SelectStatment As String, ByVal HelpTitle As String)

    '      Result = New String(2)
    '      Dim hh As New Help()
    '      hh.Text = HelpTitle

    '      hh.No_Of_Parameters = 1
    '      hh.NoOfCol = 2

    '      hh.Para1_Name = "@sql"
    '      hh.Para1_Type = SqlDbType.VarChar
    '      hh.Para1_Value = SelectStatment

    '      hh.First_Col_Name = "الكود"
    '      hh.Second_Col_Name = "الاســـم"
    '      hh.tableName = "Help"
    '      hh.ShowDialog()
    '	Result[0]=hh.Result1
    '	Result[1]=hh.Result2

    '}
    Public Function RetrieveNameOnly(ByVal SqlStatment As String) As String

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            MyCmd.CommandText = SqlStatment
            Dim ss As String = MyCmd.ExecuteScalar().ToString().Trim
            If ss.Trim() = "" Then

                Return -1
            End If
            Return ss.ToString().Trim

        Catch
            Return -1
        Finally
            c = Nothing
        End Try
    End Function

    Public Function GetMax(ByVal Length As Integer) As String

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            DefineValues()
            Dim kf As Integer = KeyFields.Length

            DiscountKeyFiels = 1
            If (kf > 1) Then
                MyCmd.CommandText = " select max(" & KeyFields(kf - 1) & ") from " & Table_Name & " T " & WhereKeyFields()
            Else : MyCmd.CommandText = " select max(" & KeyFields(kf - 1) & ") from " & Table_Name
            End If

            Dim ss1 As String = MyCmd.ExecuteScalar().ToString().Trim

            Dim ss As Integer = 0

            DiscountKeyFiels = 0
            If (ss1.Trim() = "") Then
                ss = 1
            Else

                ss = Integer.Parse(ss1)
                ss += 1
            End If
            Return ss.ToString().Trim.PadLeft(Length, "0")

        Catch ex As Exception
            Return ex.Message.ToString()
        Finally
            c = Nothing
        End Try
    End Function

    Public Function LoadConnString(ByVal file As String) As String

        Dim ConnectionString As String = ""
        Dim sr As StreamReader = New StreamReader(file)
        sr.ReadLine()
        sr.ReadLine()
        For i As Integer = 0 To 19
            sr.Read()
        Next
        ConnectionString = sr.ReadToEnd()
        sr.Close()
        Return ConnectionString
    End Function

    Public Sub ConfirmAppExit(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)

        'If Not e.Cancel Then
        '    Dim mm As New Message("هل أنت متأكد الخروج من البرنامج", True)
        '    mm.Text = "تأكيد"
        '    mm.ShowDialog()
        '    If Not mm.Ok Then

        '        e.Cancel = True
        '        Return
        '    End If
        'End If
    End Sub

    Public Function FillDate(ByVal dd As DateTime) As String

        Dim dd1 As String = dd.Month.ToString().Trim
        dd1 &= "/" & dd.Day.ToString() & "/" & dd.Year.ToString()
        Return dd1
    End Function

    Public Sub FillCombo(ByVal TableName As String, ByVal cbo As ComboBox, ByVal Condition As String)

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "select 0 Id, '-' Name union Select Id,Name From " & TableName & "  " & Condition
            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim ds As New DataSet
            da.SelectCommand = MyCmd
            da.Fill(ds, TableName)
            Dim d As Integer = ds.Tables(0).Rows.Count
            cbo.DataSource = ds.Tables(TableName)
            cbo.DisplayMember = "NAME"
            cbo.ValueMember = "ID"

            cbo.SelectedValue = 0
            cbo.SelectedIndex = 0
        Catch ex As Exception
            Dim s As String = ex.Message
        Finally
            c = Nothing
        End Try
    End Sub


    Public Sub FillCombo(ByVal SqlStatment As String, ByVal cbo As ComboBox)

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = SqlStatment
            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim ds As New DataSet()
            da.SelectCommand = MyCmd
            da.Fill(ds, "Table1")
            Dim d As Integer = ds.Tables(0).Rows.Count
            cbo.DataSource = ds.Tables("Table1")
            cbo.DisplayMember = "NAME"
            cbo.ValueMember = "Id"
            cbo.SelectedValue = 0
            cbo.SelectedIndex = 0
        Catch ex As Exception
            Dim s As String = ex.Message
        Finally
            c = Nothing
        End Try
    End Sub

    Public Sub FillListBox(ByVal SqlStatment As String, ByVal Lst As ListBox)

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = SqlStatment
            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim ds As New DataSet()
            da.SelectCommand = MyCmd
            da.Fill(ds, "Table1")
            Dim d As Integer = ds.Tables(0).Rows.Count
            Lst.DataSource = ds.Tables("Table1")
            Lst.DisplayMember = "NAME"
            Lst.ValueMember = "Id"
            Lst.SelectedIndex = 0
        Catch ex As Exception
            Dim s As String = ex.Message
        Finally
            c = Nothing
        End Try
    End Sub

    Public Sub FillListBox(ByVal DT As DataTable, ByVal Lst As ListBox)

        Try
            Lst.DataSource = DT
            Lst.DisplayMember = "NAME"
            Lst.ValueMember = "Id"
            Lst.SelectedIndex = 0
        Catch ex As Exception
            Dim s As String = ex.Message
        End Try
    End Sub

    Public Sub FillCombo(ByVal SqlStatment As String, ByVal cbo As DataGridViewComboBoxColumn)

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = SqlStatment
            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim ds As New DataSet()
            da.SelectCommand = MyCmd
            da.Fill(ds, "Table1")
            Dim d As Integer = ds.Tables(0).Rows.Count
            cbo.DataSource = ds.Tables("Table1")
            cbo.ValueMember = "Id"
            cbo.DisplayMember = "NAME"
        Catch ex As Exception
            Dim s As String = ex.Message
        Finally
            c = Nothing
        End Try
    End Sub

    Public Sub FillCombo(ByVal SqlStatment As String, ByVal cbo As DataGridViewComboBoxCell)

        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = SqlStatment
            Dim da As SqlDataAdapter = New SqlDataAdapter()
            Dim ds As New DataSet()
            da.SelectCommand = MyCmd
            da.Fill(ds, "Table1")
            Dim d As Integer = ds.Tables(0).Rows.Count
            cbo.DataSource = ds.Tables("Table1")
            cbo.DisplayMember = "NAME"
            cbo.ValueMember = "Id"
        Catch ex As Exception
            Dim s As String = ex.Message
        Finally
            c = Nothing
        End Try
    End Sub

    Public Function ExcuteNonQuery(ByVal sqlstatment As String) As Boolean

        If Not StopPro() Then Return False
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "SET XACT_ABORT ON BEGIN TRAN " & vbCrLf & sqlstatment & vbCrLf & " COMMIT"

            MyCmd.ExecuteNonQuery()
            MyCmd.Connection.Close()
        Catch ex As Exception
            Dim s As String = ex.Message.ToString()
            c = Nothing
            Return False
        Finally
            c = Nothing
        End Try
        Return True

    End Function
    Public Function StopPro() As Boolean
        If Not Md.StopProfiler Then Return True
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = "StopPro"
            MyCmd.ExecuteNonQuery()
            MyCmd.Connection.Close()
        Catch ex As Exception
            Dim s As String = ex.Message.ToString()
            c = Nothing
            Return False
        Finally
            c = Nothing
        End Try
        Return True

    End Function

    Public Function ExcuteAdapter(ByVal sqlstatment As String) As DataTable

        Dim dt As New DataTable()
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            dt.Clear()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = sqlstatment
            da.SelectCommand = MyCmd

            dt.Clear()
            da.Fill(dt)

            Dim x As Integer = dt.Rows.Count
            Return dt

        Catch ex As Exception
            Dim SS As String = ex.Message.ToString()
            Return dt
        Finally
            c = Nothing
        End Try
    End Function

    Public Function ExcuteAdapter(ByVal StoredName As String, ByVal ParaName() As String, ByVal ParaValue() As String) As DataTable

        Dim dt As New DataTable()
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.StoredProcedure
            MyCmd.CommandText = StoredName
            da.SelectCommand = MyCmd

            MyCmd.Parameters.Clear()
            For i As Integer = 0 To ParaName.Length - 1
                MyCmd.Parameters.Add("@" & ParaName(i), SqlDbType.VarChar)
                MyCmd.Parameters("@" & ParaName(i)).Value = ParaValue(i)
            Next
            dt.Clear()
            da.Fill(dt)

            Return dt

        Catch ex As Exception
            Dim ss As String = ex.Message.ToString()
            Return dt
        Finally
            c = Nothing
        End Try
    End Function





    Public Sub SendEMail(ByVal Bath As String)
        Dim demoThread As New System.ComponentModel.BackgroundWorker
        AddHandler demoThread.DoWork, AddressOf MyDoWrok
        ' AddHandler demoThread.RunWorkerCompleted, AddressOf MyWorkerCompleted
        MyBath = Bath
        demoThread.RunWorkerAsync()
    End Sub
    Dim MyBath As String = ""
    Sub MyDoWrok(ByVal sender As Object, ByVal ev As EventArgs)
        SendEMail2(MyBath)
    End Sub

    Public Sub SendEMail2(ByVal Bath As String)
        Dim mailMsg As New MailMessage()
        mailMsg.From = New MailAddress("PhysicsAdvansed@gmail.com")
        mailMsg.To.Add("Mr_Aymans@Yahoo.com")
        mailMsg.To.Add("AsebBackups@yahoo.com")
        mailMsg.To.Add("PhysicsAdvansed@Yahoo.com")
        mailMsg.To.Add("Mr_Aymans@Hotmail.com")
        mailMsg.Subject = Bath.Replace(Application.Current.StartupUri.ToString & "\", "") & " " & Now.ToShortDateString & " " & Now.ToShortTimeString
        mailMsg.IsBodyHtml = True
        mailMsg.BodyEncoding = Encoding.UTF8
        mailMsg.Attachments.Add(New System.Net.Mail.Attachment(Bath))
        mailMsg.Body = "Thanks, Mr. Physics"
        mailMsg.Priority = MailPriority.High
        ' Smtp configuration
        Dim client As New SmtpClient()
        client.Credentials = New NetworkCredential("PhysicsAdvansed@gmail.com", "0000000000")
        client.Port = 587 'or use 465
        client.Host = "smtp.gmail.com"
        client.EnableSsl = True
        client.Timeout = 72000000
        Dim userState = mailMsg
        Try
            'you can also call client.Send(msg)
            client.SendAsync(mailMsg, userState)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub ExcuteNonQuery(ByVal StoredName As String, ByVal ParaName() As String, ByVal ParaValue() As String)

        If Not StopPro() Then Return
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.StoredProcedure
            MyCmd.CommandText = StoredName
            For i As Integer = 0 To ParaName.Length - 1

                MyCmd.Parameters.Add("@" & ParaName(i), SqlDbType.VarChar)
                MyCmd.Parameters("@" & ParaName(i)).Value = ParaValue(i)
            Next
            MyCmd.ExecuteNonQuery()
            MyCmd.Connection.Close()
        Catch ee As Exception
            Dim ss As String = ee.Message.ToString()
            ShowMSG(ee.Message.ToString())
        Finally
            c = Nothing
        End Try
    End Sub

    Public Function ExcuteAdapter(ByVal StoredName As String, ByVal ParaName() As String, ByVal ParaValue() As String, ByVal Type() As SqlDbType) As DataTable

        Dim dt As New DataTable()
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.StoredProcedure
            MyCmd.CommandText = StoredName
            da.SelectCommand = MyCmd

            For i As Integer = 0 To ParaName.Length - 1

                MyCmd.Parameters.Add("@" & ParaName(i), Type(i))
                MyCmd.Parameters("@" & ParaName(i)).Value = ParaValue(i)
            Next
            dt.Clear()
            da.Fill(dt)

            Return dt
        Catch ee As Exception
            Dim ss As String = ee.Message.ToString()
            Return dt
        Finally
            c = Nothing
        End Try
    End Function

    Public Sub ExcuteNonQuery(ByVal StoredName As String, ByVal ParaName() As String, ByVal ParaValue() As String, ByVal Type() As SqlDbType)

        If Not StopPro() Then Return
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.StoredProcedure
            MyCmd.CommandText = StoredName

            For i As Integer = 0 To ParaName.Length - 1

                MyCmd.Parameters.Add("@" & ParaName(i), Type(i))
                MyCmd.Parameters("@" & ParaName(i)).Value = ParaValue(i)
            Next
            MyCmd.ExecuteNonQuery()
            MyCmd.Connection.Close()
        Catch ee As Exception
            Dim ss As String = ee.Message.ToString()
        Finally
            c = Nothing
        End Try
    End Sub

    Public Function ExecuteScalar(ByVal sqlstatment As String) As String

        If Not StopPro() Then Return ""
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.Text
            MyCmd.CommandText = sqlstatment

            Return MyCmd.ExecuteScalar().ToString().Trim
        Catch ex As Exception
            Dim s As String = ex.Message.ToString()
            Return ""
        Finally
            c = Nothing
        End Try
    End Function


    Public Function ExecuteScalar(ByVal StoredName As String, ByVal ParaName() As String, ByVal ParaValue() As String) As String

        If Not StopPro() Then Return ""
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.StoredProcedure
            MyCmd.CommandText = StoredName
            For i As Integer = 0 To ParaName.Length - 1

                MyCmd.Parameters.Add("@" & ParaName(i), SqlDbType.VarChar)
                MyCmd.Parameters("@" & ParaName(i)).Value = ParaValue(i)
            Next
            Return MyCmd.ExecuteScalar().ToString().Trim

        Catch ee As Exception
            Dim ss As String = ee.Message.ToString()
            Return ""
        Finally
            c = Nothing
        End Try
    End Function


    Public Function ToStrDate(ByVal dd As DateTime) As String

        Return dd.Month.ToString().PadLeft(2, "0") & "/" & dd.Day.ToString().PadLeft(2, "0") & "/" & dd.Year.ToString()
    End Function


    Public Function ToStrDateTime(ByVal dd As DateTime) As String

        Return (dd.Month.ToString().PadLeft(2, "0") & "/" & dd.Day.ToString().PadLeft(2, "0") & "/" & dd.ToString.Substring(6)).Replace("ص", "am").Replace("م", "pm")
    End Function

    'public Function  ToStrDate(ByVal date as AMS.TextBox.DateTextBox )as string

    '    dim dd As DateTime = Convert.ToDateTime             (Date.Value.ToString()))
    '    Return dd.Month.ToString().PadLeft(0, "0") & "/" & dd.Day.ToString().PadLeft(0, "0") & "/" & dd.Year.ToString()
    'End Function

    'public sub ShowHelp(TwoCol As Boolean  ,sender As AcceptRejectRule          object ,sender2 As object ,ProcName As string ,TitleTxt As string ,                     paraname() As String , paravalue() As String )

    '      Try

    '          Dim hh As New Help
    '          If (TwoCol) Then

    '              hh.NoOfCol = 2
    '              hh.First_Col_Name = FirstColName
    '              hh.Second_Col_Name = SecondColName

    '          Else

    '              hh.NoOfCol = 1
    '              hh.First_Col_Name = ColName1
    '          End If
    '          hh.Text = TitleTxt
    '          hh.No_Of_Parameters = 1
    '          hh.paraname = paraname
    '          hh.paravalue = paravalue
    '          hh.tableName = ProcName
    '          hh.ShowDialog()
    '		Dim  c1 As New Control=sender
    '		dim c2 As New Control=sender2

    '          If Not sender = sender2 Then

    '              c2.Text = hh.Result2.Trim()
    '              c1.Text = hh.Result1.Trim()

    '          Else

    '              c1.Text = hh.Result1.Trim()
    '          End If
    '          'selectnex

    '      Catch ex As Exception

    '          Dim ss As String = ex.Message.ToString()
    '      End Try
    '  End Sub
    '  Public Sub ShowHelp(ByVal TwoCol As Boolean, ByVal sender As Object, ByVal sender2 As Object, ByVal TableName As String, ByVal TitleTxt As String)

    '      Try

    '          'AppendWhere=" Where CompanyID='"&GLobal.CompanyID&"'"
    '          Dim SqlStatment As String = ""
    '          If (stat = "") Then
    '              SqlStatment = "Select ID,Name From " & TableName & " " & AppendWhere
    '          Else
    '              SqlStatment = TableName
    '          End If

    '          AppendWhere = ""
    '          Dim hh As New Help()
    '          If (TwoCol) Then

    '              hh.NoOfCol = 2
    '              hh.First_Col_Name = FirstColName
    '              hh.Second_Col_Name = SecondColName

    '          Else

    '              hh.NoOfCol = 1
    '              hh.First_Col_Name = ColName1
    '          End If

    '          hh.Text = TitleTxt
    '          hh.paraname = New String() {"Sql"}
    '          hh.paravalue = New String() {SqlStatment}
    '          hh.tableName = "Help"
    '          hh.ShowDialog()
    '          Dim c1 As new Control = sender 
    '		dim c2 As new Control =sender2

    '          If Not sender = sender2 Then

    '              c2.Text = hh.Result2.Trim()
    '              c1.Text = hh.Result1.Trim()

    '          Else

    '              c1.Text = hh.Result1.Trim()
    '          End If
    '          'selectnex

    '      Catch ex As Exception

    '          Dim ss As String = ex.Message.ToString()
    '      End Try
    '  End Sub

    'public sub ShowHelp(object sender,object sender2,string SqlStatment,string TitleTxt,bool TwoCol)
    '{
    '      Try
    '	{
    '          Help(hh = New Help())
    '          If (TwoCol) Then
    '		{
    '              hh.NoOfCol = 2
    '              hh.First_Col_Name = FirstColName
    '              hh.Second_Col_Name = SecondColName
    '		}
    '          Else
    '		{
    '              hh.NoOfCol = 1
    '              hh.First_Col_Name = ColName1
    '		}
    '              hh.Text = TitleTxt
    '		hh.paraname=new string[]{"Sql"}
    '		hh.paravalue=new string[]{SqlStatment}
    '              hh.tableName = HelpProc

    '              hh.ShowDialog()
    '		Control c1=(Control)sender
    '		Control c2=(Control)sender2

    '              If (sender! = sender2) Then
    '		{
    '                  c2.Text = hh.Result2.Trim()
    '                  c1.Text = hh.Result1.Trim()
    '		}
    '              Else
    '		{
    '                  c1.Text = hh.Result1.Trim()
    '		}
    '                  'selectnex
    '	}
    '	catch(Exception ex)
    '	{
    '		string ss=ex.Message.ToString()				
    '	}
    '}
    'public sub ShowHelpOneCol(object sender,string SqlStatment,string TitleTxt)
    '{
    '      Try
    '	{
    '          Help(hh = New Help())

    '          hh.NoOfCol = 1
    '          hh.First_Col_Name = ColName1

    '          hh.Text = TitleTxt
    '		hh.paraname=new string[]{"Sql"}
    '		hh.paravalue=new string[]{SqlStatment}
    '          hh.tableName = HelpProc

    '          hh.ShowDialog()
    '		Control c1=(Control)sender

    '          c1.Text = hh.Result1.Trim()

    '	}
    '	catch(Exception ex)
    '	{
    '		string ss=ex.Message.ToString()				
    '	}
    '}

    '  Public Sub ResetHelpCols()
    '{
    '      FirstColName = "الكــود"
    '      SecondColName = "الاســــم"
    '      ColName1 = "الرقـــم"
    '}
    'public sub ShowMessage(string msg,string txt)
    '{
    '      Message(mm = New Message(msg, False))
    '      mm.Text = txt
    '      mm.ShowDialog()
    '}
    'public sub ShowMessage(string msg)
    '{
    '      Message(mm = New Message(msg, False))
    '      mm.Text = "تنبيه"
    '      mm.ShowDialog()
    '}
    'public sub LoadSearch(string table,ComboBox cboSearchResult,string txtSearch ,short start)
    '{
    '      Try
    '	{
    '                                                                                                                                                                                                                                                                          SqlCommand(())
    '          MyCmd.CommandText = "Select ID,Name From " & table & " Where "
    '		if(txtSearch !="")
    '		{
    '              If (start = 1) Then
    '                  MyCmd.CommandText &= "   Name Like '" & txtSearch & "%'"
    '              ElseIf (start = 2) Then
    '                  MyCmd.CommandText &= "   Name Like '%" & txtSearch & "'"
    '              Else
    '                  MyCmd.CommandText &= "   Name Like '%" & txtSearch & "%'"
    '		}
    '                  DataTable(dt = New DataTable("Items"))
    '                  da.SelectCommand = MyCmd
    '                  dt.Clear()
    '                  da.Fill(dt)
    '                  cboSearchResult.DataSource = dt
    '                  cboSearchResult.DisplayMember = "Name"
    '                  cboSearchResult.ValueMember = "ID"
    '	}
    '	catch{}
    '}
    'public sub LoadSearch(string table,string ColName,ComboBox cboSearchResult,string txtSearch ,short start)
    '{
    '      Try
    '	{
    '		string ColID="ID"				
    '                                                                                                                                                                                                                                                                              SqlCommand(())
    '          MyCmd.CommandText = "Select " & ColID & " ID," & ColName & " Name From " & table & " Where "
    '		if(txtSearch !="")
    '		{
    '              If (start = 1) Then
    '                  MyCmd.CommandText &= "   " & ColName & " Like '" & txtSearch & "%'"
    '              ElseIf (start = 2) Then
    '                  MyCmd.CommandText &= "   " & ColName & " Like '%" & txtSearch & "'"
    '              Else
    '                  MyCmd.CommandText &= "   " & ColName & " Like '%" & txtSearch & "%'"
    '		}
    '                  DataTable(dt = New DataTable("Items"))
    '                  da.SelectCommand = MyCmd
    '                  dt.Clear()
    '                  da.Fill(dt)
    '                  cboSearchResult.DataSource = dt
    '                  cboSearchResult.DisplayMember = "Name"
    '                  cboSearchResult.ValueMember = "ID"
    '	}
    '	catch{}
    '}
    '    Public Function GetMacAddress() As String
    '        Dim SNo As String = ""
    '        ManagementScope(theScope = New ManagementScope("\\\\" & Environment.MachineName & "\\root\\cimv2"))

    '        StringBuilder(theQueryBuilder = New StringBuilder())
    '        theQueryBuilder.Append("SELECT Macddress FROM Win32_NetworkAdapter")
    '        ObjectQuery(theQuery = New ObjectQuery(theQueryBuilder.ToString()))
    '        ManagementObjectSearcher(theSearcher = New ManagementObjectSearcher(theScope, theQuery))
    '        ManagementObjectCollection(theCollectionOfResults = theSearcher.Get())

    '			string macAdd=""
    '                                                                                                                                                                                                                                        integer(c = 0)
    '			foreach (ManagementObject theCurrentObject in theCollectionOfResults)
    '			{
    '        Try
    '				{
    '					macAdd =theCurrentObject["MACAddress"].ToString()
    '					c++
    '				}
    '				catch{}
    '			}
    '*/
    '            AdapterClass(adpt = New AdapterClass())
    '            Return adpt.ReturnMacAddress()
    '		}



    '}

    '''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Sub FillCombo(ByVal cbo As ComboBox, ByVal tbl As String, ByVal SubId As String, ByVal SubName As String)
        Try
            Dim dt As DataTable = ExcuteAdapter("select " & SubId & " 'Id'," & SubName & " 'Name From " & tbl & " union select 0 Id,'-' Name")
            cbo.DataSource = dt
            cbo.ValueMember = "Id"
            cbo.DisplayMember = "Name"
            cbo.SelectedValue = 0
            cbo.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub

    Public Sub FillCombo(ByVal cbo As ComboBox, ByVal tbl As String, ByVal SubId As String, ByVal SubName As String, ByVal where As String)
        Try
            Dim dt As DataTable = ExcuteAdapter("select " & SubId & " 'Id'," & SubName & " 'Name' From " & tbl & " " & where & " union select 0 Id,'-' Name")
            cbo.DataSource = dt
            cbo.ValueMember = "Id"
            cbo.DisplayMember = "Name"
            cbo.SelectedValue = 0
            cbo.SelectedIndex = 0
        Catch ex As Exception
        End Try
    End Sub
    Public Sub FillCombo(ByVal cbo As ComboBox, ByVal StoredName As String, ByVal ParaName() As String, ByVal ParaValue() As String)

        Dim dt As New DataTable()
        Dim c As New SqlConnection(cmd.Connection.ConnectionString)
        Try
            Dim MyCmd As SqlCommand = c.CreateCommand()
            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()

            If MyCmd.Connection.State = ConnectionState.Closed Then MyCmd.Connection.Open()
            MyCmd.CommandTimeout = 72000000
            MyCmd.Parameters.Clear()
            MyCmd.CommandType = CommandType.StoredProcedure
            MyCmd.CommandText = StoredName
            da.SelectCommand = MyCmd

            MyCmd.Parameters.Clear()
            For i As Integer = 0 To ParaName.Length - 1

                MyCmd.Parameters.Add("@" & ParaName(i), SqlDbType.VarChar)
                MyCmd.Parameters("@" & ParaName(i)).Value = ParaValue(i)
            Next
            dt.Clear()
            da.Fill(dt)

            cbo.DataSource = dt
            cbo.ValueMember = "Id"
            cbo.DisplayMember = "Name"
            cbo.SelectedValue = 0
            cbo.SelectedIndex = 0

        Catch ex As Exception
        Finally
            c = Nothing
        End Try
    End Sub

    Public Function GetDelayPeriod() As Integer
        Dim s As String = ExecuteScalar("select dbo.GetDelayPeriod()")
        If s = "" Then s = "0"
        Return Val(s)
    End Function

    Public Sub SaveImage(ByVal tbl As String, ByVal KeyId As String, ByVal KeyValue As String, ByVal Field As String, ByVal PicPox As PictureBox)
        Try

            Dim sampleImage As Bitmap
            sampleImage = PicPox.Image
            Dim mStream As New System.IO.MemoryStream
            Dim ImageBytes As Byte()
            sampleImage.Save(mStream, Imaging.ImageFormat.Jpeg)
            ImageBytes = mStream.ToArray
            Dim com As New SqlClient.SqlCommand("update " & tbl & " set " & Field & "=@MyImage where " & KeyId & "='" & KeyValue & "'", con)
            com.Parameters.Add("@MyImage", SqlDbType.Image)
            com.Parameters("@MyImage").Value = ImageBytes
            com.ExecuteNonQuery()

        Catch ex As Exception

        End Try
    End Sub


    Public Sub SaveFile(ByVal tbl As String, ByVal KeyId As String, ByVal KeyValue As String, ByVal KeyId2 As String, ByVal KeyValue2 As String, ByVal Field As String, ByVal Path As String)
        Try
            Dim ImageBytes As Byte() = File.ReadAllBytes(Path)
            Dim com As New SqlClient.SqlCommand("if exists (select * from " & tbl & " where " & KeyId & "='" & KeyValue & "' and " & KeyId2 & "='" & KeyValue2 & "') update " & tbl & " set " & Field & "=@MyImage,LastUpdate=getdate() where " & KeyId & "='" & KeyValue & "' and " & KeyId2 & "='" & KeyValue2 & "' else insert into " & tbl & "( " & KeyId & "," & KeyId2 & "," & Field & ",LastUpdate) select '" & KeyValue & "','" & KeyValue2 & "',@MyImage,getdate()", con)
            com.Parameters.Add("@MyImage", SqlDbType.Image)
            com.Parameters("@MyImage").Value = ImageBytes
            com.ExecuteNonQuery()

        Catch ex As Exception
            ShowMSG(ex.Message)
        End Try
    End Sub

    Public Sub AllowDorp(ByVal picBox As PictureBox)
        picBox.AllowDrop = True
        AddHandler picBox.DragDrop, AddressOf pictureBox_DragDrop
        AddHandler picBox.DragEnter, AddressOf pictureBox_DragEnter
    End Sub
    Private Sub pictureBox_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs)
        e.Effect = e.AllowedEffect
    End Sub

    Private Sub pictureBox_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        Dim ss() As String = e.Data.GetData(DataFormats.FileDrop)
        CType(sender, PictureBox).ImageLocation = ss(0)
    End Sub
    Public Sub ShowMSG(ByVal MSG As String)
        Dim mm As New MSG
        mm.MSG = MSG
        mm.DelMsg = False
        mm.ShowDialog()
    End Sub
    Public Function ShowDeleteMSG(ByVal MSG As String) As Boolean
        Dim mm As New MSG
        mm.MSG = MSG
        mm.DelMsg = True
        mm.ShowDialog()
        Return mm.Ok
    End Function

    Public Function MistakeOrUpdate() As Boolean
        Dim mm As New MSG
        mm.MistakeOrUpdate = True
        mm.MSG = "Please, Select Type of Editting ..."
        mm.DelMsg = True
        mm.ShowDialog()
        Return mm.Ok
    End Function

    Public Sub TestMistakeOrUpdate(ByVal OldUser As String, ByVal Tbl As String)
        If MistakeOrUpdate() Then
            ExcuteNonQuery("insert into Mistakes(OldUser,NewUser,Tbl,DayDate) select '" & OldUser & "','" & Md.UserName & "','" & Tbl & "',GETDATE()")
        End If
    End Sub

    Public Function Encrypt(ByVal plainText As String) As String

        Dim passPhrase As String
        Dim saltValue As String
        Dim hashAlgorithm As String
        Dim passwordIterations As Integer
        Dim initVector As String
        Dim keySize As Integer
        passPhrase = "Pa5pr@se"        ' can be any string
        saltValue = "s@tValue"        ' can be any string
        hashAlgorithm = "SHA1"             ' can be "MD5"
        passwordIterations = 2                  ' can be any number
        initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
        keySize = 256                ' can be 192 or 128

        Dim initVectorBytes As Byte()
        initVectorBytes = Encoding.ASCII.GetBytes(initVector)

        Dim saltValueBytes As Byte()
        saltValueBytes = Encoding.ASCII.GetBytes(saltValue)
        Dim plainTextBytes As Byte()
        plainTextBytes = Encoding.UTF8.GetBytes(plainText)
        Dim password As PasswordDeriveBytes
        password = New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)

        Dim keyBytes As Byte()
        keyBytes = password.GetBytes(keySize / 8)

        Dim symmetricKey As RijndaelManaged
        symmetricKey = New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC
        Dim encryptor As ICryptoTransform
        encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)
        Dim memoryStream As MemoryStream
        memoryStream = New MemoryStream()
        Dim cryptoStream As CryptoStream
        cryptoStream = New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)
        cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)

        cryptoStream.FlushFinalBlock()

        Dim cipherTextBytes As Byte()
        cipherTextBytes = memoryStream.ToArray()

        memoryStream.Close()
        cryptoStream.Close()
        Return Convert.ToBase64String(cipherTextBytes)

    End Function
    Public Function Decrypt(ByVal cipherText As String) As String

        Dim passPhrase As String
        Dim saltValue As String
        Dim hashAlgorithm As String
        Dim passwordIterations As Integer
        Dim initVector As String
        Dim keySize As Integer

        passPhrase = "Pa5pr@se"        ' can be any string
        saltValue = "s@tValue"        ' can be any string
        hashAlgorithm = "SHA1"             ' can be "MD5"
        passwordIterations = 2                  ' can be any number
        initVector = "@1B2c3D4e5F6g7H8" ' must be 16 bytes
        keySize = 256                ' can be 192 or 128


        Dim initVectorBytes As Byte()
        initVectorBytes = Encoding.ASCII.GetBytes(initVector)

        Dim saltValueBytes As Byte()
        saltValueBytes = Encoding.ASCII.GetBytes(saltValue)

        Dim cipherTextBytes As Byte()
        cipherTextBytes = Convert.FromBase64String(cipherText)
        Dim password As PasswordDeriveBytes
        password = New PasswordDeriveBytes(passPhrase, saltValueBytes, hashAlgorithm, passwordIterations)
        Dim keyBytes As Byte()
        keyBytes = password.GetBytes(keySize / 8)
        Dim symmetricKey As RijndaelManaged
        symmetricKey = New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC
        Dim decryptor As ICryptoTransform
        decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)
        Dim memoryStream As MemoryStream
        memoryStream = New MemoryStream(cipherTextBytes)
        Dim cryptoStream As CryptoStream
        cryptoStream = New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)

        Dim plainTextBytes As Byte()
        ReDim plainTextBytes(cipherTextBytes.Length)
        Dim decryptedByteCount As Integer
        decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)

        memoryStream.Close()
        cryptoStream.Close()
        Return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)

    End Function


    Public Function TestDublicatedInvoices(ByVal BranchId As String, ByVal InvoiceNo As String, ByVal RealInvoiceNo As String, ByVal Type As String) As Boolean
        If RealInvoiceNo.Trim = "" Then Return True
        Dim dt As DataTable = ExcuteAdapter("GetDublicatedRealInvoices", New String() {"Branch", "InvoiceNo", "RealInvoiceNo", "Type"}, New String() {BranchId, InvoiceNo, RealInvoiceNo, Type})
        If dt.Rows.Count = 0 Then
            Return False
        ElseIf dt.Rows.Count <> 1 Then
            Dim s As String = ""
            Dim i As Integer = 0
            If dt.Rows(0)(0).ToString = "X" Then i = 1

            Select Case dt.Rows(i)(0).ToString
                Case "0"
                    s = "Detail"
                Case "1"
                    s = "Memberships"
                Case "2"
                    s = "Bundle"
                Case "R"
                    s = "Remaining"
                Case "I"
                    s = "Installments"
                Case "O"
                    s = "Outcome"
            End Select
            ShowMSG("Real Invoice Exists in ( " & s & " )" & vbCrLf & "Invoice ( " & dt.Rows(i)(1).ToString & " )")
            Return False
        End If
        Return True
    End Function

End Class
