Attribute VB_Name = "mdlFunctions"
Option Compare Database
Option Explicit
Public clickCount As Integer
Public lastSymbol As String
Public TicTacToe(0, 8)
Public player1 As String
Public player2 As String

Public Sub changeSymbol()
    If lastSymbol = "X" Then lastSymbol = "O" Else lastSymbol = "X"
    Call refreshMessage
End Sub

Public Sub refreshMessage()
    Select Case lastSymbol
        Case "X"
            Forms!frmTicTacToe!x.Visible = True
            Forms!frmTicTacToe!o.Visible = False
            Forms!frmTicTacToe!p2.FontBold = False
            Forms!frmTicTacToe!p1.FontBold = True
        Case "O"
            Forms!frmTicTacToe!x.Visible = False
            Forms!frmTicTacToe!o.Visible = True
            Forms!frmTicTacToe!p1.FontBold = False
            Forms!frmTicTacToe!p2.FontBold = True
    End Select
End Sub
Public Sub changePlayers()
    player1 = InputBox("Player 1 name")
    player2 = InputBox("Player 2 name")
    Call showCurrentPlayers
    Call refreshMessage
    Forms!frmTicTacToe!s1 = 0: Forms!frmTicTacToe!s2 = 0
    Call RestartGame
End Sub
Public Sub RestartGame()
    Dim i As Integer, ctl As Control
    
    clickCount = 0
    
    For i = 0 To 8
        TicTacToe(0, i) = ""
    Next i
    For Each ctl In Forms!frmTicTacToe.Form.Controls
        If Left(ctl.Name, 1) = "r" Then ctl.Caption = ""
        If Left(ctl.Name, 1) = "l" Then ctl.Visible = False
    Next
End Sub

Public Function updateMatrix(Pos As Integer)
    TicTacToe(0, Pos) = lastSymbol
End Function

Public Function gameOver() As Integer
Dim i As Integer
    For i = 0 To 8
        gameOver = gameOver + Len(TicTacToe(0, i))
    Next
End Function

Public Sub verifyWinner(Pos As Integer)
    Dim x As String
    x = lastSymbol
    
    TicTacToe(0, Pos) = x
    
    If gameOver >= 9 Then MsgBox "Game error", vbCritical, "": RestartGame: Exit Sub
    
    If TicTacToe(0, 0) = x And TicTacToe(0, 1) = x And TicTacToe(0, 2) = x Then Call playerWinner(0)
    If TicTacToe(0, 3) = x And TicTacToe(0, 4) = x And TicTacToe(0, 5) = x Then Call playerWinner(1)
    If TicTacToe(0, 6) = x And TicTacToe(0, 7) = x And TicTacToe(0, 8) = x Then Call playerWinner(2)
    If TicTacToe(0, 0) = x And TicTacToe(0, 3) = x And TicTacToe(0, 6) = x Then Call playerWinner(3)
    If TicTacToe(0, 1) = x And TicTacToe(0, 4) = x And TicTacToe(0, 7) = x Then Call playerWinner(4)
    If TicTacToe(0, 2) = x And TicTacToe(0, 5) = x And TicTacToe(0, 8) = x Then Call playerWinner(5)
    If TicTacToe(0, 0) = x And TicTacToe(0, 4) = x And TicTacToe(0, 8) = x Then Call playerWinner(6)
    If TicTacToe(0, 2) = x And TicTacToe(0, 4) = x And TicTacToe(0, 6) = x Then Call playerWinner(7)
    Call changeSymbol
End Sub

Public Sub playerWinner(probabilityRow As Integer)
    Dim ctl As Control
    
    For Each ctl In Forms!frmTicTacToe
        If ctl.Name = "l" & probabilityRow Then ctl.Visible = True
    Next
    
    Call AddScore
    MsgBox playerWinnerName & " you WON!", vbInformation, ""
    Call RestartGame
End Sub


Public Sub AddScore()
    Select Case lastSymbol
        Case "X"
            Forms!frmTicTacToe!s1 = Forms!frmTicTacToe!s1 + 1
        Case "O"
            Forms!frmTicTacToe!s2 = Forms!frmTicTacToe!s2 + 1
    End Select
End Sub

Public Function playerWinnerName() As String
    Select Case lastSymbol
        Case "X"
            playerWinnerName = player1
        Case "O"
            playerWinnerName = player2
    End Select
End Function

Public Sub showCurrentPlayers()
    Forms!frmTicTacToe!p1.Caption = player1 & "(X)"
    Forms!frmTicTacToe!p2.Caption = player2 & "(O)"
End Sub
