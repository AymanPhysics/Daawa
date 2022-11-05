Public Class Form1

    Public Password As String = ""

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Return
        If Not Exists Then
            Dim p As New PCs
            p.TextBox1.Text = s
            p.TextBox1.SelectAll()
            p.TextBox1.Focus()
            p.ShowDialog()
            Application.Current.Shutdown()
        End If
    End Sub
    Dim s As String = ""
    Dim Exists As Boolean = False
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim bm As New BasicMethods
        s = bm.ProtectionSerial()
        Exists = bm.IF_Exists("select * from PCs where Name='" & bm.Encrypt(s) & "'")
    End Sub
End Class
