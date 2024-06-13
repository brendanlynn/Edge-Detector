Public Class Form1
    Private workingBitmap As Bitmap
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveFileDialog1.ShowDialog()
    End Sub
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        workingBitmap = Bitmap.FromFile(OpenFileDialog1.FileName)
        PictureBox1.Image = workingBitmap
    End Sub
    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        If Not IsNothing(workingBitmap) Then workingBitmap.Save(SaveFileDialog1.FileName)
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If IsNothing(workingBitmap) Then Return
        If workingBitmap.Width < 3 Or workingBitmap.Height < 3 Then Return
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Form2.ProgressBar1.Value = 0
        Form2.Show()
        Dim computeThread As New Threading.Thread(Sub(usingBitmap As Bitmap)
                                                      Dim newBitmap As New Bitmap(usingBitmap.Width, usingBitmap.Height)
                                                      Dim length As UInt64 = newBitmap.Width * newBitmap.Height
                                                      Dim lastFraction As Integer
                                                      For y As UInt64 = 1 To usingBitmap.Height - 2
                                                          For x As UInt64 = 1 To usingBitmap.Width - 2
                                                              Dim r As Double = Math.Abs(0.5 * (CType(usingBitmap.GetPixel(x + 1, y - 1).R, Double) + CType(usingBitmap.GetPixel(x + 1, y + 1).R, Double) - CType(usingBitmap.GetPixel(x - 1, y - 1).R, Double) - CType(usingBitmap.GetPixel(x - 1, y + 1).R, Double)) - CType(usingBitmap.GetPixel(x - 1, y).R, Double) + CType(usingBitmap.GetPixel(x + 1, y).R, Double)) + Math.Abs(0.5 * (CType(usingBitmap.GetPixel(x - 1, y + 1).R, Double) + CType(usingBitmap.GetPixel(x + 1, y + 1).R, Double) - CType(usingBitmap.GetPixel(x - 1, y - 1).R, Double) - CType(usingBitmap.GetPixel(x + 1, y - 1).R, Double)) - CType(usingBitmap.GetPixel(x, y - 1).R, Double) + CType(usingBitmap.GetPixel(x, y - 1).R, Double))
                                                              Dim g As Double = Math.Abs(0.5 * (CType(usingBitmap.GetPixel(x + 1, y - 1).G, Double) + CType(usingBitmap.GetPixel(x + 1, y + 1).G, Double) - CType(usingBitmap.GetPixel(x - 1, y - 1).G, Double) - CType(usingBitmap.GetPixel(x - 1, y + 1).G, Double)) - CType(usingBitmap.GetPixel(x - 1, y).G, Double) + CType(usingBitmap.GetPixel(x + 1, y).G, Double)) + Math.Abs(0.5 * (CType(usingBitmap.GetPixel(x - 1, y + 1).G, Double) + CType(usingBitmap.GetPixel(x + 1, y + 1).G, Double) - CType(usingBitmap.GetPixel(x - 1, y - 1).G, Double) - CType(usingBitmap.GetPixel(x + 1, y - 1).G, Double)) - CType(usingBitmap.GetPixel(x, y - 1).G, Double) + CType(usingBitmap.GetPixel(x, y - 1).G, Double))
                                                              Dim b As Double = Math.Abs(0.5 * (CType(usingBitmap.GetPixel(x + 1, y - 1).B, Double) + CType(usingBitmap.GetPixel(x + 1, y + 1).B, Double) - CType(usingBitmap.GetPixel(x - 1, y - 1).B, Double) - CType(usingBitmap.GetPixel(x - 1, y + 1).B, Double)) - CType(usingBitmap.GetPixel(x - 1, y).B, Double) + CType(usingBitmap.GetPixel(x + 1, y).B, Double)) + Math.Abs(0.5 * (CType(usingBitmap.GetPixel(x - 1, y + 1).B, Double) + CType(usingBitmap.GetPixel(x + 1, y + 1).B, Double) - CType(usingBitmap.GetPixel(x - 1, y - 1).B, Double) - CType(usingBitmap.GetPixel(x + 1, y - 1).B, Double)) - CType(usingBitmap.GetPixel(x, y - 1).B, Double) + CType(usingBitmap.GetPixel(x, y - 1).B, Double))
                                                              Dim change As Double = (r + g + b) / 6
                                                              newBitmap.SetPixel(x, y, Color.FromArgb(change, change, change))
                                                              Dim fraction As Integer = Math.Round((x - 1 + (y - 1) * (newBitmap.Width - 2)) / length * 1000)
                                                              If fraction <> lastFraction Then
                                                                  lastFraction = fraction
                                                                  Me.Invoke(Sub()
                                                                                Form2.ProgressBar1.Value = fraction
                                                                            End Sub)
                                                              End If
                                                          Next
                                                      Next
                                                      Me.Invoke(Sub()
                                                                    workingBitmap = newBitmap
                                                                    PictureBox1.Image = workingBitmap
                                                                    Button1.Enabled = True
                                                                    Button2.Enabled = True
                                                                    Button3.Enabled = True
                                                                    Form2.Hide()
                                                                End Sub)
                                                  End Sub)
        computeThread.Start(New Bitmap(workingBitmap))
    End Sub
End Class
