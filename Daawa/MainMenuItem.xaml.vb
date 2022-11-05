Imports System.Data
Imports System.Windows.Threading

Public Class MainMenuItem

    Public NLevel As Boolean = False
    Dim m As MainWindow = Application.Current.MainWindow
    Dim bm As New BasicMethods

    WithEvents t As New DispatcherTimer With {.IsEnabled = True, .Interval = New TimeSpan(0, 0, 1)}

    Sub GetCurrent(ByVal sender As Object, ByVal e As EventArgs) Handles t.Tick
        bm.GetCurrent()
    End Sub


    Public Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem16.Click
        m.TabControl1.Items.Clear()
        m.AddTAB(sender, New Login, False)
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem2.Click
        Dim frm As New BasicForm With {.TableName = "Seasons"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem3.Click
        Dim frm As New BasicForm With {.TableName = "Lagna", .ReLoadMenue = True}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem4.Click
        Dim frm As New BasicForm2
        frm.ReLoadMenue = True
        
        frm.MainTableName = "Lagna"
        frm.MainSubId = "Id"
        frm.MainSubName = "Name"
        frm.lblMain.Content = "اللجنة"

        frm.TableName = "LagnaOperations"
        frm.MainId = "LagnaId"
        frm.SubId = "Id"
        frm.SubName = "Name"
        m.AddTab(sender, frm)
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem5.Click
        Dim frm As New BasicForm With {.TableName = "AttachmentTypes"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem88_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem88.Click
        Dim frm As New BasicForm With {.TableName = "IncomeTypes"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem6.Click
        Dim frm As New BasicForm With {.TableName = "Needs"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem8.Click
        Dim frm As New BasicForm With {.TableName = "Cities"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem9.Click
        Dim frm As New BasicForm2

        frm.MainTableName = "Cities"
        frm.MainSubId = "Id"
        frm.MainSubName = "Name"
        frm.lblMain.Content = "المحافظة"

        frm.TableName = "Areas"
        frm.MainId = "CityId"
        frm.SubId = "Id"
        frm.SubName = "Name"

        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem10.Click
        Dim frm As New Towns
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem137_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem137.Click
        Dim frm As New BasicForm4
        m.AddTAB(sender, frm)
    End Sub


    Private Sub MenuItem12_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem12.Click
        Dim frm As New BasicForm With {.TableName = "Religions"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem50_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem50.Click
        Dim frm As New Costs
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem14.Click
        Dim frm As New BasicForm With {.TableName = "EmpJobs"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem84_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem84.Click
        Dim frm As New BasicForm With {.TableName = "GuideJobs"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem86_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem86.Click
        Dim frm As New BasicForm With {.TableName = "CaseJobs"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub UserControl_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        If bm.TestIsLoaded Then Return
        LoadMenuitem()
        LoadOperations()
    End Sub

    Public Sub LoadOperations()
        MenuItem35.Items.Clear()
        Dim dt As DataTable = bm.ExcuteAdapter("select * from Lagna")
        Dim dt2 As DataTable = bm.ExcuteAdapter("select * from LagnaOperations")
        For i As Integer = 0 To dt.Rows.Count - 1
            If i <> 0 Then MenuItem35.Items.Add(New Separator)
            Dim nn As New MenuItem With {.Name = "menuitem_" & dt.Rows(i)("Id"), .Tag = dt.Rows(i)("Id"), .Header = dt.Rows(i)("Name")}
            MenuItem35.Items.Add(nn)
            Dim dr() As DataRow = dt2.Select("LagnaId=" & dt.Rows(i)("Id"))
            For i2 As Integer = 0 To dr.Length - 1
                If i2 <> 0 Then nn.Items.Add(New Separator)
                Dim nn2 As New MenuItem With {.Name = "menuitem_" & dt.Rows(i)("Id") & "_" & dr(i2)("Id"), .Tag = dr(i2)("Id"), .Header = dr(i2)("Name")}
                nn.Items.Add(nn2)
                AddHandler nn2.Click, AddressOf MenuItem_Click
            Next
        Next

    End Sub

    Private Sub MenuItem_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim frm As New Tasks With {.lagnaId = CType(sender, MenuItem).Tag, .OperationId = CType(CType(sender, MenuItem).Parent, MenuItem).Tag}
        m.AddTAB(sender, frm)
    End Sub

    Sub LoadMenuitem()
        Dim dt As DataTable = bm.ExcuteAdapter("Select * From NLevels Where Id='" & Md.LevelId & "'")
        If dt.Rows.Count = 0 Then
            Application.Current.Shutdown()
            Exit Sub
        End If

        For i As Integer = 0 To Menu1.Items.Count - 1
            Try
                Dim item As MenuItem
                item = Menu1.Items(i)
                item.Visibility = Visibility.Visible
                If Not item.Tag Is Nothing And Not item.Tag = "" Then Continue For
                item.Visibility = IIf(dt.Rows(0)(item.Name), Visibility.Visible, Visibility.Collapsed)
                LoadSub(item, dt)
            Catch
            End Try
            Try
                Dim item As Separator
                item = Menu1.Items(i)
                item.Visibility = Visibility.Visible
                If Not item.Tag Is Nothing And Not item.Tag = "" Then Continue For
                item.Visibility = IIf(dt.Rows(0)(item.Name), Visibility.Visible, Visibility.Collapsed)
            Catch
            End Try
        Next

    End Sub

    Sub LoadSub(ByVal item2 As MenuItem, ByVal dt As DataTable)
        For i As Integer = 0 To item2.Items.Count - 1
            Try
                Dim item As MenuItem
                item = item2.Items(i)
                item.Visibility = Visibility.Visible
                If Not item.Tag Is Nothing And Not item.Tag = "" Then Continue For
                item.Visibility = IIf(dt.Rows(0)(item.Name), Visibility.Visible, Visibility.Collapsed)
                LoadSub(item, dt)
            Catch
            End Try
            Try
                Dim item As Separator
                item = item2.Items(i)
                item.Visibility = Visibility.Visible
                If Not item.Tag Is Nothing And Not item.Tag = "" Then Continue For
                item.Visibility = IIf(dt.Rows(0)(item.Name), Visibility.Visible, Visibility.Collapsed)
            Catch
            End Try
        Next
    End Sub


    Private Sub MenuItem21_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem21.Click
        m.AddTAB(sender, New Suppliers)
    End Sub

    Private Sub MenuItem20_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem20.Click
        m.AddTAB(sender, New Customers)
    End Sub


    Private Sub MenuItem23_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem23.Click
        Dim frm As New CreditsDebits With {.TableName = "Debits", .LinkFile = 3}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem24.Click
        Dim frm As New CreditsDebits With {.TableName = "Credits", .LinkFile = 4}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem26_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem26.Click
        Dim frm As New CreditsDebits With {.TableName = "Saves", .LinkFile = 5}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem27_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem27.Click
        Dim frm As New CreditsDebits With {.TableName = "Banks", .LinkFile = 6}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem18_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem18.Click
        Dim frm As New Chart
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem52_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem52.Click
        Dim frm As New Attachments
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem37_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem37.Click
        Dim frm As New Employees
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem56_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem56.Click
        Dim frm As New EmployeesTemp
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem39_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem39.Click
        Dim frm As New Levels
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem32_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem32.Click, MenuItem34.Click
        Dim frm As New Cases
        If sender Is MenuItem32 Then
            frm.lblSeasonId.Visibility = Visibility.Hidden
            frm.SeasonId.Visibility = Visibility.Hidden
            frm.SeasonName.Visibility = Visibility.Hidden
        End If
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem53_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem53.Click
        Dim frm As New BasicForm With {.TableName = "IncomeReasons"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem89_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem89.Click
        Dim frm As New BasicForm With {.TableName = "NeedPeriod"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem54_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem54.Click
        Dim frm As New BasicForm With {.TableName = "OutcomeReasons"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem58_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem58.Click
        Dim frm As New BasicForm With {.TableName = "IllTypes"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem59_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem59.Click
        Dim frm As New BasicForm With {.TableName = "ProblemTypes"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem90_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem90.Click
        Dim frm As New BasicForm With {.TableName = "CaseLevels"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem91_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem91.Click
        Dim frm As New BasicForm With {.TableName = "SelaKaraba"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem60_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem60.Click
        Dim frm As New BasicForm With {.TableName = "TownTasksGoals"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem61_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem61.Click
        Dim frm As New BasicForm With {.TableName = "TownTasksSteps"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem62_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem62.Click
        Dim frm As New PreCosts
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem63_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem63.Click
        Dim frm As New TaskTown With {.LagnaId = 1, .OperationId = 1}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem68_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem68.Click
        Dim frm As New BasicForm With {.TableName = "PrintingTypes"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem66_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem66.Click
        Dim frm As New BasicForm With {.TableName = "Authors"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem67_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem67.Click
        Dim frm As New BasicForm With {.TableName = "Reviewers"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem72_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem72.Click
        Dim frm As New BasicForm With {.TableName = "Groups"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem87_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem87.Click
        Dim frm As New BasicForm With {.TableName = "CaseTypes"}
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem71_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem71.Click
        Dim frm As New Items
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem73_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem73.Click
        Dim frm As New BasicForm2

        frm.MainTableName = "Groups"
        frm.MainSubId = "Id"
        frm.MainSubName = "Name"
        frm.lblMain.Content = "المجموعة"

        frm.TableName = "Types"
        frm.MainId = "GroupId"
        frm.SubId = "Id"
        frm.SubName = "Name"
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem75_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem75.Click
        Dim frm As New Reviewing
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem79_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem79.Click
        Dim frm As New StoreMovesIn
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem81_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem81.Click
        Dim frm As New StoreMovesOut
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem83_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem83.Click
        Dim frm As New Sales
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem77_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem77.Click
        Dim frm As New Printing
        m.AddTAB(sender, frm)
    End Sub

    Private Sub PrintTbl(ByVal Header As String, ByVal tbl As String, Optional ByVal maintbl As String = "", Optional ByVal mainfield As String = "")
        Dim frm As New ReportViewer
        frm.RptPath = IIf(maintbl = "", "PrintTbl.rpt", "PrintTbl2.rpt")
        frm.paraname = {"Header", "@tbl", "@maintbl", "@mainfield"}
        frm.paravalue = {Header, tbl, maintbl, mainfield}
        frm.ShowDialog()
    End Sub


    Private Sub MenuItem93_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem93.Click
        PrintTbl("المواسم", "Seasons")
    End Sub

    Private Sub MenuItem94_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem94.Click
        PrintTbl("اللجان", "Lagna")
    End Sub

    Private Sub MenuItem95_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem95.Click
        PrintTbl("أنواع المرفقات", "AttachmentTypes")
    End Sub

    Private Sub MenuItem96_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem96.Click
        PrintTbl("أنواع الاحتياجات", "Needs")
    End Sub

    Private Sub MenuItem97_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem97.Click
        PrintTbl("أنواع فترات الاحتياجات", "NeedPeriod")
    End Sub
    Private Sub MenuItem98_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem98.Click
        PrintTbl("أهداف عمليات القرى", "TownTasksGoals")
    End Sub
    Private Sub MenuItem99_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem99.Click
        PrintTbl("المراحل التنفيذية", "TownTasksSteps")
    End Sub
    Private Sub MenuItem100_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem100.Click
        PrintTbl("المؤلفين", "Authors")
    End Sub
    Private Sub MenuItem101_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem101.Click
        PrintTbl("المراجعين", "Reviewers")
    End Sub
    Private Sub MenuItem102_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem102.Click
        PrintTbl("المحافظات", "Cities")
    End Sub
    Private Sub MenuItem103_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem103.Click
        PrintTbl("الديانات", "Religions")
    End Sub
    Private Sub MenuItem104_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem104.Click
        PrintTbl("أنواع الدخل للحالات", "IncomeTypes")
    End Sub
    Private Sub MenuItem105_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem105.Click
        PrintTbl("أنواع الحالات", "CaseTypes")
    End Sub
    Private Sub MenuItem106_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem106.Click
        PrintTbl("درجات الحالات", "CaseLevels")
    End Sub
    Private Sub MenuItem111_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem111.Click
        PrintTbl("وظائف خاصة بالدليل", "GuideJobs")
    End Sub
    Private Sub MenuItem110_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem110.Click
        PrintTbl("وظائف خاصة بالقائمين على العمل", "EmpJobs")
    End Sub
    Private Sub MenuItem109_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem109.Click
        PrintTbl("أنواع الشكاوى", "ProblemTypes")
    End Sub
    Private Sub MenuItem108_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem108.Click
        PrintTbl("أنواع الأمراض", "IllTypes")
    End Sub
    Private Sub MenuItem107_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem107.Click
        PrintTbl("صلات القرابة", "SelaKaraba")
    End Sub

    Private Sub MenuItem115_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem115.Click
        PrintTbl("مجموعات الأصناف", "Groups")
    End Sub
    Private Sub MenuItem114_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem114.Click
        PrintTbl("بنود المصروفات", "OutcomeReasons")
    End Sub
    Private Sub MenuItem113_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem113.Click
        PrintTbl("بنود الإيرادات", "IncomeReasons")
    End Sub
    Private Sub MenuItem112_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem112.Click
        PrintTbl("وظائف خاصة بالحالات", "CaseJobs")
    End Sub


    Private Sub MenuItem116_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem116.Click
        PrintTbl("عمليات اللجان", "LagnaOperations", "Lagna", "LagnaId")
    End Sub
    Private Sub MenuItem117_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem117.Click
        PrintTbl("أنواع الأصناف", "Types", "Groups", "GroupId")
    End Sub
    Private Sub MenuItem118_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem118.Click
        PrintTbl("المراكز", "Areas", "Cities", "CityId")
    End Sub
    Private Sub MenuItem141_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem141.Click
        Dim frm As New ReportViewer
        frm.RptPath = "Towns.rpt"
        frm.paraname = {"Header"}
        frm.paravalue = {"القرى"}
        frm.ShowDialog()
    End Sub


    Private Sub MenuItem133_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem133.Click
        Dim frm As New RPT1
        frm.Flag = 1
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem132_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem132.Click
        Dim frm As New RPT1
        frm.Flag = 2
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem138_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem138.Click
        Dim frm As New RPT1
        frm.Flag = 5
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem139_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem139.Click
        Dim frm As New RPT1
        frm.Flag = 6
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem140_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem140.Click
        Dim frm As New RPT1
        frm.Flag = 7
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem134_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem134.Click
        Dim frm As New RPT1
        frm.Flag = 4
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem136_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem136.Click
        Dim frm As New RPT1
        frm.Flag = 3
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem135_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem135.Click
        Dim frm As New RPT2
        frm.Flag = 1
        m.AddTAB(sender, frm)
    End Sub

    Private Sub MenuItem142_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MenuItem142.Click
        Dim frm As New RPT3
        frm.Flag = 1
        m.AddTAB(sender, frm)
    End Sub
End Class
