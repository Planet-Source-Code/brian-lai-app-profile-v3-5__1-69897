VERSION 5.00
Begin VB.Form OpenFileDlg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Open File"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OpenFileDlg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CommandButton btnExec 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton btnExec 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtFN 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.ComboBox cboFT 
      Height          =   330
      ItemData        =   "OpenFileDlg.frx":000C
      Left            =   720
      List            =   "OpenFileDlg.frx":0025
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "&Remember"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1380
      Width           =   2055
   End
   Begin VB.Image IMG 
      Height          =   495
      Left            =   120
      Picture         =   "OpenFileDlg.frx":0098
      Stretch         =   -1  'True
      ToolTipText     =   "Double-click me for full file name"
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgPadlock 
      Height          =   270
      Left            =   480
      Picture         =   "OpenFileDlg.frx":06F7
      ToolTipText     =   "Not locked by sandbox mode"
      Top             =   1365
      Width           =   195
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "FormatCode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A53928&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   180
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Open this like a"
      Height          =   210
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   1470
   End
   Begin VB.Image IMGbkg 
      Height          =   1815
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "OpenFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyChoice As Long
Dim MyHover As Integer

Public Function AsType(FileExtension As String, Optional FileNme As String, Optional IgnoreRMBFileExtFlag As Boolean) As Long
    On Error Resume Next
    Dim A As String
    A = Trim$(GetSet("PFT_" & UCase$(FileExtension), "-1"))
    If A <> "" And A <> "-1" Then 'say, -1 is the "undefined" number
        AsType = Val(Left$(A, 2))
        chkDontAsk.Value = 1
        cboFT.ListIndex = AsType
        If IgnoreRMBFileExtFlag Then GoTo ThereInstead 'ya ok...
        
        EventSound "MSGSkip"
        
        Exit Function
    Else
        chkDontAsk.Value = 0
    End If
ThereInstead:
    A = "" 'flush
    A = FileNameOnly(FileNme)
    If Len(A) > 25 Then A = "..." & Right$(A, 25)
    lblFormat.Caption = A 'UCase$(FileExtension)
    chkDontAsk.Tag = FileExtension
    Me.Tag = FileNme
    txtFN.Text = FileNme
    
    SkinForm Me 'the best place to put it
    SkinFormEx Me
    
    Me.Show 1
    AsType = MyChoice
    Unload Me
End Function

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    Dim A As String
    If Index = 0 Then
        Dim I As Integer
        MyChoice = cboFT.ListIndex
        If chkDontAsk.Value = 1 Then
            A = Str(MyChoice)
            If A <> "-1" Then 'only make a new record if it's not "undefined"
                'SaveSet "PFT_" & UCase$(lblFormat.Caption), A
                WriteINI UserName, "PFT_" & UCase$(chkDontAsk.Tag), A, SettingsFile, True  'not sandboxed
            End If
        Else
            A = "-1"
        End If
    Else
        MyChoice = 99 'Ridiculous
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    InitCommonControls
'    F1.FadeIn
End Sub

'Private Sub Form_Deactivate()
'    F1.FadeOut
'End Sub
'
'Private Sub Form_Load()
'    On Error Resume Next
'    F1.PrepareFade
'End Sub

Private Sub Form_Paint()
    On Error Resume Next
    EventSound "WinOpen"
End Sub

Private Sub Form_Unload(Cancel As Integer)

    EventSound "WinClose"
    
End Sub

'Private Sub Form_Load()
'    On Error Resume Next
'        'moved
''    SkinForm Me
''    SkinFormEx Me
'
'End Sub

Private Sub lstFT_DblClick()
    On Error Resume Next
    btnExec_Click 0
End Sub

Private Sub IMG_DblClick()
    On Error Resume Next
    txtFN.Visible = Not txtFN.Visible
End Sub

