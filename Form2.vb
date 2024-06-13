Imports System.ComponentModel

Public Class Form2
    Private Sub Form2_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        e.Cancel = True
    End Sub
End Class