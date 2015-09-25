VERSION 5.00
Begin VB.Form frmM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   14400
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   1320
      ScaleHeight     =   5055
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   480
      Width           =   9255
   End
End
Attribute VB_Name = "frmM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.ScaleMode = 3
    picM.ScaleMode = 3
    picM.AutoRedraw = True
End Sub

Private Sub Form_Resize()
    picM.Move 0, 0, 960, 720
    Call Draw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMExisted = False
End Sub

Private Sub picM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDP.Xi = X
    MouseDP.Yi = Y
    If Button = 1 Then
        MouseDL = True
    ElseIf Button = 2 Then
        MouseDL = True
    End If
End Sub

Private Sub picM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseP.Xi = X
    MouseP.Yi = Y
End Sub

Private Sub picM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MouseDL = False
    ElseIf Button = 2 Then
        MouseDL = False
    End If
End Sub
