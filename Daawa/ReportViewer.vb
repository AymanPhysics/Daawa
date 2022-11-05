Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data
Imports CrystalDecisions.Shared

Public Class ReportViewer

    Dim ReportDoc As New ReportDocument
    Dim bm As New BasicMethods
    Public RptPath As String = "", Header As String = ""
    Public paraname() As String = {}
    Public paravalue() As String = {}
    Public SubReports() As String = {}

    Private Sub ReportViewer_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ReportDoc.Dispose()
    End Sub


    Private Sub ReportViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Dim ServerName As String = con.DataSource
            Dim DataBase As String = con.Database

            ReportDoc.Load(System.Windows.Forms.Application.StartupPath & "\Rpts\" & RptPath)

            Dim stb As New SqlClient.SqlConnectionStringBuilder
            stb.ConnectionString = con.ConnectionString
            'ReportDoc.SetDatabaseLogon("Sa", "123", ServerName, DataBase)
            ReportDoc.SetDatabaseLogon(stb.UserID, stb.Password, ServerName, DataBase)

            Dim Table_LogOn_Info As New TableLogOnInfo()

            'Table_LogOn_Info.ConnectionInfo.UserID = "sa"
            'Table_LogOn_Info.ConnectionInfo.Password = "123"
            Table_LogOn_Info.ConnectionInfo.UserID = stb.UserID
            Table_LogOn_Info.ConnectionInfo.Password = stb.Password

            Table_LogOn_Info.ConnectionInfo.ServerName = ServerName
            Table_LogOn_Info.ConnectionInfo.DatabaseName = DataBase

            Dim TableServer() As String
            For Each Table_In_Report As CrystalDecisions.CrystalReports.Engine.Table In ReportDoc.Database.Tables
                Table_In_Report.ApplyLogOnInfo(Table_LogOn_Info)
                TableServer = Table_In_Report.Location.Split(".")
                Try
                    Table_In_Report.Location = DataBase & "." & TableServer(1) & "." & TableServer(2)
                Catch ex As Exception
                    Table_In_Report.Location = DataBase & ".dbo." & TableServer(0)
                End Try
            Next

            For i As Integer = 0 To SubReports.Length - 1

                Try

                    For Each Table_In_Report As CrystalDecisions.CrystalReports.Engine.Table In ReportDoc.OpenSubreport(SubReports(i)).Database.Tables
                        Table_In_Report.ApplyLogOnInfo(Table_LogOn_Info)
                        TableServer = Table_In_Report.Location.Split(".")
                        Try
                            Table_In_Report.Location = DataBase & "." & TableServer(1) & "." & TableServer(2)
                        Catch
                            Table_In_Report.Location = DataBase & ".dbo." & TableServer(0)
                        End Try
                    Next
                Catch
                End Try
            Next

            SetParamValue("DataBase", DataBase)
            SetParamValue("UserName", Md.UserName)
            SetParamValue("EnName", Md.EnName)
            Try
                SetParamValue("Header", IIf(Header = "", CType(CType(Application.Current.Windows(0), MainWindow).TabControl1.SelectedItem, TabItem).Header, Header))
            Catch
            End Try

            For i As Integer = 0 To paraname.Length - 1
                SetParamValue(paraname(i), paravalue(i))
            Next

            CrystalReportViewer1.ReportSource = ReportDoc

        Catch ex As Exception

        End Try

    End Sub

    Public Sub Print(Optional ByVal ServerName As String = "", Optional ByVal PrinterName As String = "", Optional ByVal NoOfCopies As Integer = 1)
        ReportViewer_Load(Nothing, Nothing)
        Try
            If PrinterName <> "" Then ReportDoc.PrintOptions.PrinterName = "\\" & ServerName & "\" & PrinterName
            For i As Integer = 1 To NoOfCopies
                ReportDoc.PrintToPrinter(1, False, 1, 5)
            Next
        Catch ex As Exception
            bm.ShowMSG(ex.Message)
        End Try

    End Sub
    Private Sub SetParamValue(ByVal paramName As String, ByVal paramValue As String)

        For i As Integer = 0 To ReportDoc.DataDefinition.ParameterFields.Count - 1
            If ReportDoc.DataDefinition.ParameterFields(i).ParameterFieldName = paramName Then
                Dim PFD As ParameterFieldDefinition = ReportDoc.DataDefinition.ParameterFields(i)
                Dim PValues As New ParameterValues()
                Dim Parm As New ParameterDiscreteValue()
                Parm.Value = paramValue
                PValues.Add(Parm)
                Try
                    PFD.ApplyCurrentValues(PValues)
                Catch ex As Exception
                End Try
                'Exit For
            End If
        Next

        For i As Integer = 0 To SubReports.Length - 1
            Try
                For i2 As Integer = 0 To ReportDoc.OpenSubreport(SubReports(i)).DataDefinition.ParameterFields.Count - 1
                    If (ReportDoc.OpenSubreport(SubReports(i)).DataDefinition.ParameterFields(i2).ParameterFieldName.ToLower() = paramName.ToLower()) Then
                        Dim PFD As ParameterFieldDefinition = ReportDoc.OpenSubreport(SubReports(i)).DataDefinition.ParameterFields(i2)
                        Dim PValues As ParameterValues = New ParameterValues()
                        Dim Parm As ParameterDiscreteValue = New ParameterDiscreteValue()
                        Parm.Value = paramValue.Trim()
                        PValues.Add(Parm)
                        PFD.ApplyCurrentValues(PValues)
                        'Exit For
                    End If
                Next
            Catch
            End Try
        Next
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Close()
    End Sub
End Class