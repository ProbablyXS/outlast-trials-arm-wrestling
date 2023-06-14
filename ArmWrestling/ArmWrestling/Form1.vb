Imports System.Runtime.InteropServices
Imports WindowsInput

Public Class Form1

    Dim NPoint, NColor
    Dim InputSim As New InputSimulator
    Dim Rval = "0", Gval = "0", Bval = "0"

    Dim t1 As Integer
    Dim t2 As Integer

    Dim bitmap As Bitmap

    Private Function loopbitmap() As Task

        Try

            hWnd = FindWindow(TextBox1.Text, vbNullString)

            If hWnd <> 0 Then

                Dim wr As New RECT

                DwmGetWindowAttribute(hWnd, DWMWA_EXTENDED_FRAME_BOUNDS, wr, Marshal.SizeOf(wr))


                OL = wr.left + 1
                OT = wr.top + 31
                RR = wr.right - 1
                OB = wr.bottom - 1

                window_h = OB - OT
                window_w = RR - OL

                t1 = window_w / 2
                t2 = window_h / 2


                bitmap = New Bitmap(t1 - Convert.ToInt32(NumericUpDown4.Value), t2 - Convert.ToInt32(NumericUpDown3.Value))
                g = Graphics.FromImage(bitmap)
                g.CopyFromScreen(OL + t1 / 2 + NumericUpDown1.Value, OT + t2 / 2 + NumericUpDown2.Value, 0, 0, bitmap.Size)

            End If

        Catch ex As Exception

        End Try

    End Function

    Private Async Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try

            While True

                Await Task.Delay(10)

                loopbitmap()

                If Button1.Text = "Start" Then

                    Exit While

                End If

                'Try

                If Me.PictureBox1.Image Is Nothing Then

                Else

                    PictureBox1.Image.Dispose()

                End If

                'Dim r As Long = 70 'radius
                'Dim ox As Long = 120 'starting point X (middle)
                'Dim oy As Long = 90 'starting point Y (middle)
                'Dim ar As Double ' angle en radians
                'Dim i As Long
                'Dim sinus As Double
                'Dim cosinus As Double
                'Dim x As Long
                'Dim y As Long
                'Dim xy As PointF() = New PointF(360) {} ' nombre degrés boucle : matrice v(x, y)
                Dim c As New Pen(Color.Red, 1) ' couleur rgb 2) = taille trait

                'For i = 0 To 360 Step 6
                '    ar = i * 3.14 / 180
                '    cosinus = Math.Cos(ar)
                '    x = r * cosinus + ox
                '    xy(i).X = x
                '    sinus = Math.Sin(ar)
                '    y = oy - r * sinus
                '    xy(i).Y = y

                '    g.DrawLine(c, ox, 90, x, y)
                'Next i

                While True

                    Await Task.Delay(1)

                    For a = 0 To bitmap.Width - 1 Step 4
                        For b = 0 To bitmap.Height - 1 Step 4

                            NColor = bitmap.GetPixel(a, b)

                            If NColor.R >= 210 And NColor.R <= 225 And NColor.G >= 83 And NColor.G <= 90 And NColor.B >= 48 And NColor.B <= 58 Then 'rouge
                                'NColor.R >= 250 And NColor.R <= 256 And NColor.G >= 250 And NColor.G <= 256 And NColor.B >= 250 And NColor.B <= 256 Then 'blanc

                                Rval = a.ToString
                                Gval = b.ToString
                                Bval = "Waiting"

                                PictureBox1.Image = bitmap.Clone

                                While True

                                    If Button1.Text = "Start" Then

                                        Exit While

                                    End If

                                    Await Task.Delay(10)

                                    loopbitmap()

                                    NColor = bitmap.GetPixel(a, b)

                                    g.DrawLine(c, 120, 90, a, b)
                                    PictureBox1.Image = bitmap.Clone

                                    If NColor.R >= 175 And NColor.R <= 190 And NColor.G >= 175 And NColor.G <= 190 And NColor.B >= 175 And NColor.B <= 190 Then 'aiguille

                                        Label1.ForeColor = Color.Green

                                        InputSim.Mouse.LeftButtonClick()
                                        Bval = "OK"

                                        'Await Task.Delay(500)

                                        Exit While

                                    End If

                                End While

                            End If

                        Next
                    Next

                    PictureBox1.Image = bitmap.Clone
                    bitmap.Dispose()
                    g.Dispose()

                    Exit While

                End While

            End While

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        getval()
    End Sub

    Private Async Function getval() As Task

        Do

            Await Task.Delay(100)

            If (Label1.ForeColor = Color.Green) Then

                Label1.Text = "Done"

                Await Task.Delay(1000)

                Label1.ForeColor = Color.Black

            End If

            Label1.Text = Rval + " " + Gval + " " + Bval

        Loop

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Button1.Text = "Start" Then

            BackgroundWorker1.RunWorkerAsync()
            Button1.Text = "Stop"

        Else

            Button1.Text = "Start"

        End If

        Rval = "0"
        Gval = "0"
        Bval = "0"

    End Sub
End Class
