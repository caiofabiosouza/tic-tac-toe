VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btChangePlayers_Click()
    Call changePlayers
End Sub

Private Sub Form_Load()
    lastSymbol = "X"
    Call changePlayers
    Call refreshMessage
End Sub

Private Sub r0_Click()
    If verifyClickSymbol(r0) = True Then Exit Sub
    Me.r0.Caption = lastSymbol
    Call verifyWinner(0)
End Sub

Private Sub r1_Click()
    If verifyClickSymbol(r1) = True Then Exit Sub
    Me.r1.Caption = lastSymbol
    Call verifyWinner(1)
End Sub

Private Sub r2_Click()
    If verifyClickSymbol(r2) = True Then Exit Sub
    Me.r2.Caption = lastSymbol
    Call verifyWinner(2)
    
End Sub

Private Sub r3_Click()
    If verifyClickSymbol(r3) = True Then Exit Sub
    Me.r3.Caption = lastSymbol
    Call verifyWinner(3)
    
End Sub

Private Sub r4_Click()
    If verifyClickSymbol(r4) = True Then Exit Sub
    Me.r4.Caption = lastSymbol
    Call verifyWinner(4)
    
End Sub

Private Sub r5_Click()
    If verifyClickSymbol(r5) = True Then Exit Sub
    Me.r5.Caption = lastSymbol
    Call verifyWinner(5)

End Sub

Private Sub r6_Click()
    If verifyClickSymbol(r6) = True Then Exit Sub
    Me.r6.Caption = lastSymbol
    Call verifyWinner(6)
    
End Sub

Private Sub r7_Click()
    If verifyClickSymbol(r7) = True Then Exit Sub
    Me.r7.Caption = lastSymbol
    Call verifyWinner(7)
    
End Sub

Private Sub r8_Click()
    If verifyClickSymbol(r8) = True Then Exit Sub
    Me.r8.Caption = lastSymbol
    Call verifyWinner(8)
    
End Sub

Function verifyClickSymbol(ctl As Control) As Boolean
    If Len(ctl.Caption) > 0 Then
        verifyClickSymbol = True
        MsgBox "Click on another area!", vbExclamation, ""
    Else
        Select Case lastSymbol
            Case "X"
                ctl.ForeColor = Me.x.BorderColor
            Case "O"
                ctl.ForeColor = Me.o.BorderColor
        End Select
    End If
End Function
