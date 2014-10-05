Imports System.IO

Public Class Puzzle
    Private tileSize As Integer = 100
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

    Private Function HomePosition(number As Integer)
        Return New Point(number Mod 4, Math.Floor(number / 4))
    End Function

    Private Positions As List(Of Point) =
        Enumerable.Range(0, 15) _
        .Select(Function(i) New Point(i Mod 4, Math.Floor(i / 4))) _
        .ToList()

    ReadOnly Iterator Property AllTiles As IEnumerable(Of Button)
        Get
            For Each item As Control In Controls
                If Not TypeOf item Is Button Then
                    Continue For
                End If
                Dim tile = DirectCast(item, Button)
                If (tile.Width <> tileSize Or tile.Height <> tileSize) Then
                    Continue For
                End If
                Yield tile
            Next
        End Get
    End Property

    Private Sub Shuffle()
        TeleportToSolvedState()
        Positions.ForEach(Sub(position) Swap(position, New Point(Math.Floor(3 * Rnd()), Math.Floor(3 * Rnd()))))
        Positions.ForEach(Sub(position) SetButtonBackground(GetButtonByPosition(position), position))
        moveCount = 0
        done = False
        SetWindowTitle(My.Resources.All.StatusNew)
    End Sub

    Private Sub TeleportToSolvedState()
        AllTiles.ToList().ForEach(Sub(tile) TeleportButton(tile, HomePosition(Int(tile.Text))))
    End Sub

    Private Sub Swap(a As Point, b As Point)
        If a = b Then
            Return
        End If
        Dim btn1 = GetButtonByPosition(a)
        Dim btn2 = GetButtonByPosition(b)
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

    Private Function GetButtonPosition(btn As Button)
        Return New Point(Math.Floor(btn.Left / tileSize), Math.Floor(btn.Top / tileSize))
    End Function

    Private Function GetButtonByPosition(position As Point) As Button
        Return AllTiles.Where(Function(tile) GetButtonPosition(tile) = position).FirstOrDefault
    End Function

    Private Sub TeleportButton(btn As Button, position As Point)
        btn.Left = tileSize * position.X + fieldOffset.X
        btn.Top = tileSize * position.Y + fieldOffset.Y
    End Sub

    Private Sub MoveButton(btn As Button, position As Point)
        TeleportButton(btn, position)
        SetButtonBackground(btn, position)
    End Sub

    Private Sub SetButtonBackground(tile As Button, position As Point)
        If IsInHome(position) Then
            tile.BackgroundImage = My.Resources.All.button0
        ElseIf tile IsNot Nothing Then
            tile.BackgroundImage = My.Resources.All.button3
        End If
    End Sub

    Private Function IsInHome(position As Point)
        Return IsButtonInHome(position, GetButtonByPosition(position))
    End Function

    Private Function IsButtonInHome(position As Point, button As Button)
        Return button IsNot Nothing AndAlso button.Text.Equals(CStr(position.X + (position.Y * 4) + 1))
    End Function

    Private Function IsSolved()
        Return Enumerable.Range(0, 14) _
            .Select(Function(i) HomePosition(i)) _
            .All(Function(position) IsInHome(position))
    End Function

    Private Sub SetWindowTitle(message As String)
        Text = String.Format(My.Resources.All.Window_Title, message)
    End Sub

    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Tile1.Click, Tile2.Click, Tile3.Click, Tile4.Click, Tile5.Click, Tile6.Click, Tile7.Click, Tile8.Click, Tile9.Click, Tile10.Click, Tile11.Click, Tile12.Click, Tile13.Click, Tile14.Click, Tile15.Click
        If done Then
            Select Case MessageBox.Show(String.Format(My.Resources.All.Solved, moveCount), My.Resources.All.StartNew, MessageBoxButtons.YesNo)
                Case Windows.Forms.DialogResult.No
                    Shuffle()
            End Select
            Return
        End If
        Dim btn = DirectCast(sender, Button)
        Dim position = GetButtonPosition(btn)
        Dim direction = GetDirectionToEmpty(position)
        Dim sound As UnmanagedMemoryStream
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
            Dim result = MessageBox.Show(My.Resources.All.GiveUp_StartNew, My.Resources.All.StartNew, MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If result = DialogResult.No Then
                Return
            End If
        End If
        Shuffle()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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
