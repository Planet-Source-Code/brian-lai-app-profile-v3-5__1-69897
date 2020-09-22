VERSION 5.00
Begin VB.Form frmTXT 
   AutoRedraw      =   -1  'True
   Caption         =   "Document"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00AB7013&
   Icon            =   "frmTXT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8010
   ScaleWidth      =   9720
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox picROT 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9720
      TabIndex        =   1
      Top             =   0
      Width           =   9720
      Begin VB.CommandButton btnExec 
         Caption         =   "Encrypt"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   3
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton btnExec 
         Caption         =   "Decrypt"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   2
         Top             =   60
         Width           =   1095
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Copy"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":000C
         PICN            =   "frmTXT.frx":0028
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   1
         Left            =   1575
         TabIndex        =   5
         ToolTipText     =   "Cut"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":015A
         PICN            =   "frmTXT.frx":0176
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   2
         Left            =   1950
         TabIndex        =   6
         ToolTipText     =   "Paste"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":02A1
         PICN            =   "frmTXT.frx":02BD
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "New"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":03EB
         PICN            =   "frmTXT.frx":0407
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   4
         Left            =   375
         TabIndex        =   8
         ToolTipText     =   "Open"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":0535
         PICN            =   "frmTXT.frx":0551
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   5
         Left            =   750
         TabIndex        =   9
         ToolTipText     =   "Save"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":0666
         PICN            =   "frmTXT.frx":0682
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   6
         Left            =   2400
         TabIndex        =   10
         ToolTipText     =   "Page info"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":07B5
         PICN            =   "frmTXT.frx":07D1
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   7
         Left            =   2775
         TabIndex        =   11
         ToolTipText     =   "Font"
         Top             =   60
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   0
         MICON           =   "frmTXT.frx":0B23
         PICN            =   "frmTXT.frx":0B3F
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.TextBox txtBox 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  '¤â°Ê
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentlyOpenFile As String

Public myTag As String 'the tag used by the tabs

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    Dim I As Integer
    Dim K As String
    K = InputBox("Enter password:", , "4")
    If Len(K) = 0 Then Exit Sub
    I = Int(K)
    If I <= 0 Or I >= 10 Then
        MsgBox "You must enter an integer value higher than 0 and lower than 10."
        Exit Sub
    Else
        txtBox.Text = Encrypt(txtBox.Text, IIf(Index = 0, True, False), Int(I))
    End If
End Sub


Private Sub btnTool_Click(Index As Integer)
    On Error Resume Next
    With frmMain
        Select Case Index
            Case 0
                .titTextEditCopy_Click
            Case 1
                .titTextEditCut_Click
            Case 2
                .titTextEditPaste_Click
            Case 3
                .titLENewTextViewer_Click
            Case 4
                .titTextFileOpen_Click
            Case 5
                .titTextFileSave_Click
            Case 6
                FileProperties CurrentlyOpenFile
            Case 7
                .titTextViewFont_Click
        End Select
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    frmMain.titText.Visible = True
    frmMain.COF = Me.CurrentlyOpenFile
    
    'SyncTabs
    SyncCaptionEx Int(Me.Tag), Me.Caption
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titText.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    Mod32BitIcon.SetIcon Me.hwnd, "AAA"

    'browser index
    Dim I As Integer
    BRWIndex = BRWIndex + 1
    Me.Tag = BRWIndex
    I = frmMain.Tb.AddTab("Document")
    frmMain.Tb.TabTag(I) = Me.Tag
    'browser index


    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    EventSound "WinOpen"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Right$(Me.Caption, 1) = "*" Then 'unsaved eh
        Dim A As VbMsgBoxResult
        A = MsgBox("Do you want to save this file first?", vbYesNoCancel + vbQuestion)
        Select Case A
            Case vbYes
                frmMain.titTextFileSave_Click
            Case vbNo
                'do nothing?
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    Dim I As Long
    I = IIf(picROT.Visible, picROT.Height, 0)
    txtBox.Move 0, I, Me.ScaleWidth, Me.ScaleHeight - I
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    Dim G As Integer
    If FileLen(TheFN) > 64000 Then 'use for big files
        txtBox.Text = FileText(AddRecentItem(TheFN))
    Else
        Dim F As Integer
        Dim tmp As String, K As String
        F = FreeFile
        Open TheFN For Input As #F
            Do
                Line Input #F, tmp
                K = K & tmp & vbCrLf
            Loop Until EOF(F)
        Close #F
        txtBox.Text = K
    End If
    
        G = Int(GetSet("OnlyOneTxtViewer", "0"))

        If G > 0 Then
            Dim FRM As Form
            For Each FRM In Forms
                If TypeOf FRM Is frmTXT Then Unload FRM
            Next
        End If
    
    CurrentlyOpenFile = TheFN
    
    Me.CurrentlyOpenFile = TheFN
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    SyncCaptionEx Int(Me.Tag), Me.Caption
    
    Me.Show
    SStatus Me.Name & " opened " & TheFN, vbInformation
End Function

Public Sub ChangeFont()
    On Error Resume Next
    'Dim Response As VbMsgBoxResult
    With txtBox
        SelectFont.mFontName = txtBox.FontName
        SelectFont.mFontSize = txtBox.FontSize
        SelectFont.mBold = txtBox.FontBold
        SelectFont.mFontColor = txtBox.ForeColor
        SelectFont.mItalic = txtBox.FontItalic
        SelectFont.mStrikethru = txtBox.FontStrikethru
        SelectFont.mUnderline = txtBox.FontUnderline
        
        ShowFont
        .FontName = SelectFont.mFontName
        .FontSize = SelectFont.mFontSize
        .FontBold = SelectFont.mBold
        .FontItalic = SelectFont.mItalic
        .FontStrikethru = SelectFont.mStrikethru
        .FontUnderline = SelectFont.mUnderline
        .ForeColor = SelectFont.mFontColor
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
    EventSound "WinClose"
    Me.Tag = "CLOSED" 'suppose this is a tag for not marking the browser on the tabs
    SyncTabs
End Sub

Private Sub txtBox_Change()
    On Error Resume Next
    If Right$(Me.Caption, 1) <> "*" Then CCaption Me.Caption & "*", Me 'state of change
    
'    EventSound "Type"
    
    SStatus Len(txtBox.Text) & " characters [" & CurrentlyOpenFile & " ]", vbInformation
    
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next 'shortcuts
    Select Case Shift
        Case 2 'Ctrl
            Select Case KeyCode
                Case vbKeyA
                    frmMain.titTextEditSelectAll_Click
                Case vbKeyC
                    frmMain.titTextEditCopy_Click
                Case vbKeyD
                    frmMain.titTextViewFont_Click
                Case vbKeyO
                    frmMain.titTextFileOpen_Click
                Case vbKeyP
                    frmMain.titTextEditPaste_Click
                Case vbKeyS
                    frmMain.titTextFileSave_Click
                Case vbKeyU
                    frmMain.titTextFileOpenURL_Click
                Case vbKeyX
                    frmMain.titTextEditCut_Click
            End Select
    End Select
End Sub

Private Sub txtBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    LoadFile Data.Files.Item(1) 'just load it
End Sub

Private Sub txtVal_Change()

End Sub
