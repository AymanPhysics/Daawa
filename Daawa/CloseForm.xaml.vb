Public Class CloseForm
    Dim bm As New BasicMethods

    Public State As BasicMethods.CloseState = BasicMethods.CloseState.Cancel

    Private Sub Window_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
    End Sub

    Private Sub btnYes_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnYes.Click
        State = BasicMethods.CloseState.Yes
    End Sub

    Private Sub btnNo_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnNo.Click
        State = BasicMethods.CloseState.No
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        State = BasicMethods.CloseState.Cancel
    End Sub
End Class
