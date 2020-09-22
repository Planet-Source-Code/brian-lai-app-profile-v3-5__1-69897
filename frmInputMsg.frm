VERSION 5.00
Begin VB.Form frmInputMsg 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   " "
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
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
   Icon            =   "frmInputMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.TextBox txtW 
      Height          =   1455
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton btnYN 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton btnYN 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton BTN 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox CHK 
      Caption         =   "Remember answer"
      Height          =   210
      Left            =   720
      TabIndex        =   3
      Top             =   1755
      Width           =   2175
   End
   Begin VB.Image imgPadlock 
      Height          =   270
      Left            =   480
      Picture         =   "frmInputMsg.frx":000C
      ToolTipText     =   "Not locked by sandbox mode"
      Top             =   1725
      Width           =   195
   End
   Begin VB.Image IMG 
      Height          =   480
      Left            =   120
      Picture         =   "frmInputMsg.frx":031E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label LBL 
      BackStyle       =   0  '³z©ú
      Height          =   1530
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image IMGbkg 
      Height          =   2175
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmInputMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ORLY As Integer

Public Function MyMsgBoxEx(Prompt As String, SaveNum As Integer, Optional MSGStyle As VbMsgBoxStyle = vbOKOnly, Optional titLE As String, Optional HideCheckBox As Boolean) As VbMsgBoxResult
    On Error Resume Next
    Dim I As Integer
    I = Val(GetSet("DSA" & SaveNum, "0"))
    If I <> 0 Then 'if this message is set not to show again
        MyMsgBoxEx = ValToResult(I)
        'SStatus "DSA: " & SaveNum
        
        EventSound "MSGSkip"
        
        Exit Function 'then exit
    End If
    ShowButtonType IIf(MSGStyle = vbOKOnly, 0, 1) 'changes buttons
    LBL.Caption = Prompt
    CHK.Visible = Not HideCheckBox 'shows and hides the checkbox
    imgPadlock.Visible = Not HideCheckBox 'added padlock image
    CCaption titLE, Me
    Me.Tag = SaveNum
    Me.Show 1
    MyMsgBoxEx = ValToResult(ORLY)
End Function

Private Sub BTN_Click()
    On Error Resume Next
    ORLY = 1 'say 2 is YES and 3 is NO
    'SaveSet "DSA" & Me.Tag, IIf(CHK.Value = 0, 0, ORLY)
    WriteINI UserName, "DSA" & Me.Tag, IIf(CHK.Value = 0, 0, ORLY), SettingsFile, True 'not sandboxed
    Unload Me
End Sub

Private Sub btnYN_Click(Index As Integer)
    On Error Resume Next
    ORLY = Index + 2 'say 2 is YES and 3 is NO
    'SaveSet "DSA" & Me.Tag, IIf(CHK.Value = 0, 0, ORLY)
    WriteINI UserName, "DSA" & Me.Tag, IIf(CHK.Value = 0, 0, ORLY), SettingsFile, True 'not sandboxed
    Unload Me
End Sub

Private Sub Form_Activate()
    InitCommonControls
'    F1.FadeIn
End Sub

'Private Sub Form_Deactivate()
'    F1.FadeOut
'End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    SkinForm Me
    SkinFormEx Me

    EventSound "WinOpen"
    
'    F1.PrepareFade

End Sub

Sub ShowButtonType(Optional Which As Integer = 0)
    On Error Resume Next
    BTN.Visible = (Which = 0)
    btnYN(0).Visible = (Which <> 0)
    btnYN(1).Visible = (Which <> 0)
End Sub

Function ValToResult(Valu As Integer) As Integer
    On Error Resume Next
        Select Case Valu
        Case 1
            ValToResult = Val(vbOK)
        Case 2
            ValToResult = Val(vbYes)
        Case 3
            ValToResult = Val(vbNo)
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    EventSound "WinClose"
End Sub

Private Sub LBL_DblClick()
    On Error Resume Next 'copying apparatus
    txtW.Text = LBL.Caption
    txtW.Visible = Not txtW.Visible
End Sub

Private Sub txtW_DblClick()
LBL_DblClick
End Sub
