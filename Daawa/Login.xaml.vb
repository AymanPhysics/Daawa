Imports System.Data

Public Class Login

    Dim bm As New BasicMethods
    Public LogedIn As Boolean = False
    Public Flag As Integer = 1
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click
        If Username.Text.Trim = "" Or Username.SelectedIndex < 1 Then
            Username.Focus()
            Return
        End If

        If Password.Password.Trim = "" Then
            Password.Focus()
            Return
        End If

        Dim dt As DataTable

        Md.CountryId = 0 ' dt.Rows(0)("CountryId").ToString
        Md.CityId = 0 ' dt.Rows(0)("CityId").ToString

        If Not bm.StopPro() Then Return
        Dim paraname() As String = {"Id", "Password"}
        Dim paravalue() As String = {Username.SelectedValue.ToString, bm.Encrypt(Password.Password)}
        dt = bm.ExcuteAdapter("TestLogin", paraname, paravalue)
        If dt.Rows.Count = 0 Then
            bm.ShowMSG("كلمة المرور غير صحيحة ...")
            Exit Sub
        End If
        Md.UserName = Username.SelectedValue.ToString
        Md.ArName = dt.Rows(0)("ArName").ToString
        Md.EnName = dt.Rows(0)("EnName").ToString
        Md.JobId = dt.Rows(0)("JobId").ToString
        Md.Manager = dt.Rows(0)("Manager").ToString
        Md.LevelId = dt.Rows(0)("LevelId").ToString
        Md.Password = bm.Decrypt(dt.Rows(0)("Password").ToString)
        Md.GeneralManager = dt.Rows(0)("GeneralManager").ToString
        Md.Board = dt.Rows(0)("Board").ToString
        Md.Accountant = dt.Rows(0)("Accountant").ToString
        Md.Cashier = dt.Rows(0)("Cashier").ToString
        Md.SystemAdmin = dt.Rows(0)("SystemAdmin").ToString
        Md.CompanyName = dt.Rows(0)("CompanyName").ToString

        Dim m As MainWindow = Application.Current.MainWindow
        If MyProjectType = ProjectType.PCs Then
            Dim frm As New BasicForm
            frm.TableName = "PCs"
            frm.txtName.MaxLength = 1000
            m.TabControl1.Items.Clear()
            m.AddTab(New MenuItem With {.Header = "PCs"}, frm)
        ElseIf Flag = 1 Then
            m.LoadTabs()
        End If
        LogedIn = True
    End Sub

    Private Sub Login_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If bm.TestIsLoaded Then Return
        bm.FillCombo("select Id,EnName Name from Employees where SystemUser='1' and Stopped='0' union select 0 Id,'-' Name order by Name", Username)
        'bm.ApplyKeyDown(New Control() {Stores, Username, Password})
        Username.Focus()
    End Sub
End Class
