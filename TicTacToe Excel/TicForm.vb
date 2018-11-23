VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TicForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3285
   OleObjectBlob   =   "TicForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private xTurn As Boolean
Private endGame As Boolean

Private Function getTurn() As String
    getTurn = IIf(xTurn, "X", "O") ' In normal VB change 'getTurn =' by 'return'
End Function

Private Function checkVictory() As Boolean
    Dim check(0 To 7) As Boolean
    Dim turn As String
    turn = getTurn()
    
    ' Victory positions
    check(0) = (b0.Caption = turn And b1.Caption = turn And b2.Caption = turn)
    check(1) = (b3.Caption = turn And b4.Caption = turn And b5.Caption = turn)
    check(2) = (b6.Caption = turn And b7.Caption = turn And b8.Caption = turn)
    
    check(3) = (b0.Caption = turn And b3.Caption = turn And b6.Caption = turn)
    check(4) = (b1.Caption = turn And b4.Caption = turn And b7.Caption = turn)
    check(5) = (b2.Caption = turn And b5.Caption = turn And b8.Caption = turn)
    
    check(6) = (b0.Caption = turn And b4.Caption = turn And b8.Caption = turn)
    check(7) = (b2.Caption = turn And b4.Caption = turn And b6.Caption = turn)
    
   ' MsgBox check(0)
    For Each Item In check
        If Item Then
           infodump.Caption = turn & " won"
           endGame = True
           checkVictory = True
           Exit Function
        End If
    Next
    
    checkVictory = False
End Function

Private Function checkDraw() As Boolean
    Dim draw As Boolean
    Dim turn As String
    turn = getTurn()
    
    draw = b0.Caption <> "" And b1.Caption <> "" And b2.Caption <> "" And b3.Caption <> "" And b4.Caption <> "" And b5.Caption <> "" And b6.Caption <> "" And b7.Caption <> "" And b8.Caption <> ""
    
    If draw Then
       infodump.Caption = "draw"
       endGame = True
       checkDraw = True
       Exit Function
    End If
  
    
    checkDraw = False
End Function

Private Sub selectBtn(ByRef btn As CommandButton)
    If endGame Or btn.Caption <> "" Then
        Exit Sub
    End If
    
    btn.Caption = getTurn()
    
    If checkVictory() Then
        Exit Sub
    End If
     
    If checkDraw() Then
        Exit Sub
    End If
    
    xTurn = Not xTurn
    infodump.Caption = getTurn() & " turn"
End Sub

' evnts

Private Sub UserForm_Initialize()
    xTurn = True
    endGame = False
End Sub

Private Sub b0_Click()
    Call selectBtn(b0)
End Sub

Private Sub b1_Click()
    Call selectBtn(b1)
End Sub

Private Sub b2_Click()
    Call selectBtn(b2)
End Sub

Private Sub b3_Click()
    Call selectBtn(b3)
End Sub

Private Sub b4_Click()
    Call selectBtn(b4)
End Sub

Private Sub b5_Click()
    Call selectBtn(b5)
End Sub

Private Sub b6_Click()
    Call selectBtn(b6)
End Sub

Private Sub b7_Click()
    Call selectBtn(b7)
End Sub

Private Sub b8_Click()
    Call selectBtn(b8)
End Sub
