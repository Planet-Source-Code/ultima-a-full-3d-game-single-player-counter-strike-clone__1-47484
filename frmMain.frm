VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "SturmNacht"
   ClientHeight    =   3195
   ClientLeft      =   4995
   ClientTop       =   5250
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Running As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Running = False
    End If
End Sub

Private Sub Form_Load()
    DoEvents
    Me.Show
    
    Init Me.HWND
    
    Running = True
    
    While Running = True
        MainLoop
    Wend
    KillApp
End Sub
