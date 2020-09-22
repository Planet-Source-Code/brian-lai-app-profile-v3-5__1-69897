VERSION 5.00
Begin VB.Form frmMood 
   Caption         =   "Moodometer"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMood.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   7335
   WindowState     =   2  '³Ì¤j¤Æ
   Begin ProFile.CB btnMood 
      Height          =   1095
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1931
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMood.frx":000C
      PICPOS          =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ProFile.CB btnMood 
      Height          =   1095
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1931
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMood.frx":0028
      PICPOS          =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ProFile.CB btnMood 
      Height          =   1095
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1931
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMood.frx":0044
      PICPOS          =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ProFile.CB btnMood 
      Height          =   1095
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1931
      BTYPE           =   9
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmMood.frx":0060
      PICPOS          =   0
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblAxis 
      Alignment       =   2  '¸m¤¤¹ï»ô
      AutoSize        =   -1  'True
      Caption         =   "Slow"
      Height          =   225
      Index           =   3
      Left            =   930
      TabIndex        =   7
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label lblAxis 
      Alignment       =   2  '¸m¤¤¹ï»ô
      AutoSize        =   -1  'True
      Caption         =   "Fast"
      Height          =   225
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label lblAxis 
      Alignment       =   2  '¸m¤¤¹ï»ô
      AutoSize        =   -1  'True
      Caption         =   "Depressing"
      Height          =   225
      Index           =   1
      Left            =   1800
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblAxis 
      Alignment       =   2  '¸m¤¤¹ï»ô
      AutoSize        =   -1  'True
      Caption         =   "Happy"
      Height          =   225
      Index           =   0
      Left            =   2010
      TabIndex        =   4
      Top             =   1680
      Width           =   555
   End
   Begin VB.Shape Axis 
      Height          =   2415
      Index           =   1
      Left            =   2280
      Top             =   1920
      Width           =   15
   End
   Begin VB.Shape Axis 
      Height          =   15
      Index           =   0
      Left            =   960
      Top             =   3120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Pd As Integer = 15

Private Sub btnMood_Click(Index As Integer)
    On Error Resume Next
    MsgBox "Not implemented", vbCritical
End Sub

Private Sub btnMood_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim I As Long
    Dim K As String
    For I = 1 To Data.Files.Count Step 1
        K = Data.Files(I)
        SStatus "Adding mood for " & Data.Files(I), vbExclamation
        WriteINI "ProFile", FileNameOnly(Data.Files(I), True), CStr(Index), FindPath(PathOnly(Data.Files(I)), "mood.ini")
    Next
    SStatus
    For I = 0 To btnMood.UBound Step 1
        If I <> Index Then
            btnMood(I).Caption = ""
        Else
            btnMood(I).Caption = Data.Files.Count & " file" & IIf(Data.Files.Count > 1, "s", "") & " added"
        End If
    Next
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    'SyncTabs
    SyncCaptionEx Int(Me.Tag), Me.Caption
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    'browser index
    Dim I As Integer
    BRWIndex = BRWIndex + 1
    Me.Tag = BRWIndex
    I = frmMain.Tb.AddTab("Loading...")
    frmMain.Tb.TabTag(I) = Me.Tag
    'browser index

    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me
    
    SyncCaptionEx Int(Me.Tag), Me.Caption
    
    EventSound "WinOpen"

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With lblAxis(0)
        .Move (Me.ScaleWidth - .Width) \ 2, 0
    End With
    With lblAxis(1)
        .Move (Me.ScaleWidth - .Width) \ 2, Me.ScaleHeight - .Height
    End With
    With lblAxis(2)
        .Move Me.ScaleWidth - .Width, (Me.ScaleHeight - .Height) \ 2
    End With
    With lblAxis(3)
        .Move 0, (Me.ScaleHeight - .Height) \ 2
    End With
    Axis(0).Move 0, (Me.ScaleHeight) \ 2, Me.ScaleWidth, 15
    Axis(1).Move Me.ScaleWidth \ 2, lblAxis(0).Height, 15, Me.ScaleHeight
    btnMood(0).Move lblAxis(3).Width + Pd, lblAxis(0).Height + Pd, Axis(1).Left - lblAxis(3).Width - Pd * 2, Axis(0).Top - lblAxis(0).Height - Pd * 2
    btnMood(1).Move Axis(1).Left + Axis(1).Width + Pd, btnMood(0).Top, lblAxis(2).Left - Axis(1).Left - Axis(1).Width - Pd * 2, btnMood(0).Height
    btnMood(2).Move btnMood(0).Left, Axis(0).Top + Axis(0).Height + Pd, btnMood(0).Width, lblAxis(1).Top - Axis(0).Top - Axis(0).Height - Pd * 2
    btnMood(3).Move btnMood(1).Left, btnMood(2).Top, btnMood(1).Width, btnMood(2).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    EventSound "WinClose"
    Me.Tag = "CLOSED" 'suppose this is a tag for not marking the browser on the tabs
    SyncTabs
End Sub

