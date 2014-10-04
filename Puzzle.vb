Public Class Puzzle
    Private buttonSize As Integer = 100
    Private empty As Point
    Private fieldOffset As Point
    Private moveCount As Integer
    Private done As Boolean

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MenuStrip1.Items(0).Text = My.Resources.All.Game
        NewToolStripMenuItem.Text = My.Resources.All.StartNew
        ExitToolStripMenuItem.Text = My.Resources.All.Quit
        Me.Icon = My.Resources.All.Icon15
        fieldOffset.Y = 25
        Shuffle()
    End Sub

    Private Iterator Function AllPoints() As IEnumerable(Of Point)
        For i As Integer = 1 To 15
            Yield New Point(i Mod 4, Math.Floor(i / 4))
        Next
    End Function

    Private Sub Shuffle()
        For Each position As Point In AllPoints()
            Swap(position, New Point(Math.Floor(3 * Rnd()), Math.Floor(3 * Rnd())))
        Next
        For Each position As Point In AllPoints()
            SetButtonBackground(GetButtonByPosition(position), position)
        Next
        moveCount = 0
        done = False
        SetWindowTitle(My.Resources.All.StatusNew)
    End Sub

    Private Sub Swap(a As Point, b As Point)
        If a = b Then
            Return
        End If
        Dim btn1 As Button = GetButtonByPosition(a)
        Dim btn2 As Button = GetButtonByPosition(b)
        If btn1 Is Nothing Then
            TeleportButton(btn2, a)
            empty = b
        ElseIf btn2 Is Nothing Then
            TeleportButton(btn1, b)
            empty = a
        Else
            TeleportButton(btn1, b)
            TeleportButton(btn2, a)
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
        Return New Point(Math.Floor(btn.Left / buttonSize), Math.Floor(btn.Top / buttonSize))
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

    Private Sub TeleportButton(btn As Button, position As Point)
        btn.Left = buttonSize * position.X + fieldOffset.X
        btn.Top = buttonSize * position.Y + fieldOffset.Y
    End Sub

    Private Sub MoveButton(btn As Button, position As Point)
        TeleportButton(btn, position)
        SetButtonBackground(btn, position)
    End Sub

    Private Sub SetButtonBackground(btn As Button, position As Point)
        If IsButtonInHome(position) Then
            btn.BackgroundImage = My.Resources.All.button0
        ElseIf btn IsNot Nothing Then
            btn.BackgroundImage = My.Resources.All.button3
        End If
    End Sub

    Private Function IsButtonInHome(position As Point)
        Return IsButtonInHome(position, GetButtonByPosition(position))
    End Function

    Private Function IsButtonInHome(position As Point, button As Button)
        If button Is Nothing OrElse Not button.Text.Equals(CStr(position.X + (position.Y * 4) + 1)) Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function IsSolved() As Boolean
        Dim i As Integer
        For i = 0 To 14
            If Not IsButtonInHome(New Point(i Mod 4, Math.Floor(i / 4))) Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Sub SetWindowTitle(message As String)
        Text = String.Format(My.Resources.All.Window_Title, message)
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button4.Click, Button3.Click, Button2.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click, Button1.Click
        If done Then
            Select Case MessageBox.Show(String.Format(My.Resources.All.Solved, moveCount), My.Resources.All.StartNew, MessageBoxButtons.YesNo)
                Case Windows.Forms.DialogResult.No
                    Shuffle()
            End Select
            Return
        End If
        Dim btn As Button = DirectCast(sender, Button)
        Dim position = GetButtonPosition(btn)
        Dim direction = GetDirectionToEmpty(position)
        Dim sound As System.IO.UnmanagedMemoryStream
        My.Computer.Audio.Stop()
        If direction IsNot Nothing Then
            MoveButton(btn, New Point(position.X + direction.X, position.Y + direction.Y))
            empty = position
            moveCount += 1
            sound = My.Resources.All.tap
        Else
            sound = My.Resources.All.blocked
        End If
        My.Computer.Audio.Play(sound, AudioPlayMode.Background)
        SetWindowTitle(String.Format(My.Resources.All.Moves, moveCount))
        If IsSolved() Then
            Dim message = String.Format(My.Resources.All.Victory, moveCount)
            SetWindowTitle(message)
            MessageBox.Show(message, My.Resources.All.Tadaa, MessageBoxButtons.OK, MessageBoxIcon.Information)
            done = True
        End If
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        If moveCount > 0 Then
            Dim result As Integer = MessageBox.Show(My.Resources.All.GiveUp_StartNew, My.Resources.All.StartNew, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If result = DialogResult.No Then
                Return
            End If
        End If
        Shuffle()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim message As String
        If Not done And moveCount > 0 Then
            message = My.Resources.All.Quit_Unsolved
        ElseIf Not done And moveCount = 0 Then
            message = My.Resources.All.Quit_NotStarted
        Else
            message = My.Resources.All.Quit_EnoughToday
        End If
        Select Case MessageBox.Show(message, My.Resources.All.Confirm, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            Case Windows.Forms.DialogResult.No
                e.Cancel = True
        End Select
    End Sub
End Class
