Public Class Form1
    Private buttonSize As Integer = 100
    Private empty As Point
    Private fieldOffset As Point
    Private moveCount As Integer
    Private done As Boolean

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.Resource1._15
        fieldOffset.Y = 25
        Shuffle()
    End Sub

    Private Sub Shuffle()
        Randomize()
        Dim i As Integer
        For i = 0 To 15
            Swap(New Point(i Mod 4, Math.Floor(i / 4)), New Point(Math.Floor(3 * Rnd()), Math.Floor(3 * Rnd())))
        Next
        moveCount = 0
        done = False
        SetWindowTitle("New")
    End Sub

    Private Sub Swap(a As Point, b As Point)
        If a = b Then
            Return
        End If
        Dim btn1 As Button = GetButtonByPosition(a)
        Dim btn2 As Button = GetButtonByPosition(b)
        If btn1 Is Nothing Then
            MoveButton(btn2, a)
            empty = b
        ElseIf btn2 Is Nothing Then
            MoveButton(btn1, b)
            empty = a
        Else
            MoveButton(btn1, b)
            MoveButton(btn2, a)
        End If
    End Sub

    Private Function GetDirectionToEmpty(position As Point)
        Dim direction = New Point(empty.X - position.X, empty.Y - position.Y)
        If Math.Abs(direction.X) > 1 Or Math.Abs(direction.Y) > 1 Or (direction.X <> 0 And direction.Y <> 0) Then
            Return Nothing
        Else
            Return direction
        End If
    End Function

    Private Function GetButtonPosition(btn As Button) As Point
        Return New Point(Math.Floor(btn.Left / 100), Math.Floor(btn.Top / 100))
    End Function

    Private Function GetButtonByPosition(position As Point) As Button
        For Each item As Control In Controls
            If Not TypeOf item Is Button Then
                Continue For
            End If
            Dim btn As Button = DirectCast(item, Button)
            If (btn.Width <> buttonSize Or btn.Height <> buttonSize) Then
                Continue For
            End If
            If GetButtonPosition(btn) = position Then
                Return btn
            End If
        Next
        Return Nothing
    End Function

    Private Sub MoveButton(btn As Button, position As Point)
        btn.Left = buttonSize * position.X + fieldOffset.X
        btn.Top = buttonSize * position.Y + fieldOffset.Y
    End Sub

    Private Function IsSolved() As Boolean
        For i = 0 To 14
            Dim btn As Button = GetButtonByPosition(New Point(i Mod 4, Math.Floor(i / 4)))
            If btn Is Nothing OrElse Not btn.Text.Equals(CStr(i + 1)) Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub SetWindowTitle(message As String)
        Text = String.Format("Fifteen puzzle - {0}", message)
    End Sub
    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button4.Click, Button3.Click, Button2.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click, Button1.Click
        If done Then
            Select Case MessageBox.Show(String.Format("This puzzle was solved in {0} moves. Do you want to try new one?", moveCount), "Start new", MessageBoxButtons.YesNo)
                Case Windows.Forms.DialogResult.No
                    Shuffle()
            End Select
            Return
        End If
        Dim btn As Button = DirectCast(sender, Button)
        Dim position = GetButtonPosition(btn)
        Dim direction = GetDirectionToEmpty(position)
        My.Computer.Audio.Stop()
        If direction <> Nothing Then
            MoveButton(btn, New Point(position.X + direction.X, position.Y + direction.Y))
            empty = position
            moveCount += 1
            My.Computer.Audio.Play(My.Resources.Resource1.tap, AudioPlayMode.Background)

        Else
            My.Computer.Audio.Play(My.Resources.Resource1.blocked, AudioPlayMode.Background)

        End If

        SetWindowTitle(String.Format("Moves: {0}", moveCount))
        If IsSolved() Then
            Dim message = String.Format("Victory in {0} click(s)!", moveCount)
            SetWindowTitle(message)
            MessageBox.Show(message, "Tadaa!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            done = True
        End If
    End Sub
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        If moveCount > 0 Then
            Dim result As Integer = MessageBox.Show("Do you want to give up and start new game?", "Start new", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If result = DialogResult.No Then
                Return
            End If
        End If
        Shuffle()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim message As String
        If Not done And moveCount > 0 Then
            message = "There are still puzzle to solve, are you sure want to quit?"
        ElseIf Not done And moveCount = 0 Then
            message = "Giving up before even trying to solve?"
        Else
            message = "Are you sure that this was enough puzzle solving for today?"
        End If
        Select Case MessageBox.Show(message, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            Case Windows.Forms.DialogResult.No
                e.Cancel = True
        End Select
    End Sub

End Class
