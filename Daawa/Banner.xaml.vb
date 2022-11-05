' Copyright © Microsoft Corporation.  All Rights Reserved.
' This code released under the terms of the 
' Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Navigation
Imports System.Windows.Shapes
Imports System.Windows.Threading

Namespace EmployeeTracker
    ''' <summary>
    ''' Interaction logic for Banner.xaml
    ''' </summary>
    Partial Public Class Banner
        Inherits UserControl
        Dim t As New DispatcherTimer With {.IsEnabled = True, .Interval = New TimeSpan(0, 0, 0, 0, 100)}

        Public Header As String = ""
        Public StopTimer As Boolean = False
        Public Sub New()
            InitializeComponent()
            AddHandler t.Tick, AddressOf Tick
        End Sub

        Public Sub Tick()
            If StopTimer Then
                t.Stop()
                lblMain.Text = Header
                Return
            End If
            lblMain.FlowDirection = FlowDirection.LeftToRight
            lblMain.Text = "O M E G A                                     Username: " & Md.EnName
        End Sub

        Private Sub UserControl_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        End Sub
    End Class
End Namespace
