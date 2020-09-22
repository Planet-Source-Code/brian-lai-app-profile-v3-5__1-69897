VERSION 5.00
Begin VB.UserControl HProgressBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "HProgressBar.ctx":0000
   Begin VB.Image imgB 
      Height          =   195
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image Prog 
      Height          =   195
      Left            =   720
      Picture         =   "HProgressBar.ctx":0312
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2640
   End
   Begin VB.Image BackArray 
      Height          =   225
      Index           =   0
      Left            =   0
      Picture         =   "HProgressBar.ctx":2BF4
      Top             =   0
      Width           =   135
   End
   Begin VB.Image BackArray 
      Height          =   225
      Index           =   1
      Left            =   3000
      Picture         =   "HProgressBar.ctx":2DDA
      Top             =   0
      Width           =   135
   End
   Begin VB.Image BackArray 
      Height          =   225
      Index           =   2
      Left            =   720
      Picture         =   "HProgressBar.ctx":2FC0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "HProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Long, myValue As Long, myMax As Long
Dim myEnabled As Boolean

Public Event Change(MyVal As Long, myMaxVal As Long)

Private Sub UserControl_InitProperties()
    On Error Resume Next
    myEnabled = True
    myValue = 0
    myMax = 100
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Size UserControl.Width, 225 'limiting the height of the bar
    BackArray(0).Move 0, 0
    BackArray(1).Move UserControl.Width - BackArray(1).Width, 0
    BackArray(2).Move BackArray(0).Width, 0, UserControl.Width - BackArray(0).Width - BackArray(1).Width
    Prog.Move 15, 15, (UserControl.Width - 30) * myValue / (myMax + 0.0000001), 195
End Sub

Public Property Let Enabled(ByVal nwEnabled As Boolean)
    myEnabled = nwEnabled
    Prog.Visible = myEnabled
    UserControl_Resize
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = myEnabled
End Property

'Public Property Get Picture() As StdPicture
'    On Error Resume Next
'    Set Picture = Prog.Picture
'End Property
'
'Public Property Set Picture(ByVal newPic As StdPicture)
'    On Error Resume Next
'    If newPic Is Nothing Then
'        Prog.Picture = newPic
'    Else
'        Prog.Picture = imgB.Picture
'    End If
'    PropertyChanged "picture"
'End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        myEnabled = .ReadProperty("Enabled", True)
        myValue = .ReadProperty("Value", 0)
        myMax = .ReadProperty("Max", 100)
        'Set Prog.Picture = .ReadProperty("picture", imgB.Picture)
    End With 'PROPBAG
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "Enabled", myEnabled, True
        .WriteProperty "Value", myValue, 0
        .WriteProperty "Max", myMax, 100
        '.WriteProperty "picture", Prog.Picture, imgB.Picture
    End With
End Sub

Public Property Get Value() As Long
    On Error Resume Next
    Value = myValue
End Property

Public Property Let Value(ByVal nwVal As Long)
    On Error Resume Next
    myValue = nwVal
    UserControl_Resize
    PropertyChanged "Value"
End Property

Public Property Get Max() As Long
    On Error Resume Next
    Max = myMax
End Property

Public Property Let Max(ByVal nwMax As Long)
    On Error Resume Next
    myMax = nwMax
    UserControl_Resize
    PropertyChanged "Max"
End Property



