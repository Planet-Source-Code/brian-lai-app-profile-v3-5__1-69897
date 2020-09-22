VERSION 5.00
Begin VB.Form frmDumbAss 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '¨S¦³®Ø½u
   ClientHeight    =   495
   ClientLeft      =   600
   ClientTop       =   600
   ClientWidth     =   615
   ControlBox      =   0   'False
   Icon            =   "frmDumbAss.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '¤â°Ê
   ScaleHeight     =   495
   ScaleWidth      =   615
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      OLEDropMode     =   1  '¤â°Ê
      Picture         =   "frmDumbAss.frx":000C
      Stretch         =   -1  'True
      ToolTipText     =   "Drag a file here to open"
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape Border 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmDumbAss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myX As Single, myY As Single

Private Sub Form_Activate()
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Border.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Image1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    MakeTransparent Me.hWnd, 80
    
    SkinForm Me
    SkinFormEx Me
    OnTop Me.hWnd, True
'    EventSound "WinOpen" 'o please... a drop box doesnt need sounds.
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        frmMain.MDIForm_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    myX = X
    myY = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim kX As Long, kY As Long
    If Button = 1 Then
        kX = Me.Left + X - myX
        If kX < 0 Then kX = 0
        If kX > Screen.Width - Me.ScaleWidth Then kX = Screen.Width - Me.ScaleWidth
        kY = Me.Top + Y - myY
        If kY < 0 Then kY = 0
        If kY > Screen.Height - Me.ScaleHeight Then kY = Screen.Height - Me.ScaleHeight
        Me.Move kX, kY
    End If
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
        frmMain.MDIForm_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
