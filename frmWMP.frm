VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmWMP 
   AutoRedraw      =   -1  'True
   Caption         =   "Media"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   165
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWMP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   9405
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.PictureBox Pic 
      Align           =   4  '¹ï»ôªí³æ¥k¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   6435
      Left            =   7230
      ScaleHeight     =   6435
      ScaleWidth      =   2175
      TabIndex        =   7
      Top             =   495
      Visible         =   0   'False
      Width           =   2175
      Begin VB.ListBox WmpPL 
         Height          =   2940
         IntegralHeight  =   0   'False
         Left            =   0
         OLEDropMode     =   1  '¤â°Ê
         TabIndex        =   8
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox picSet 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9405
      TabIndex        =   6
      Top             =   0
      Width           =   9405
      Begin VB.TextBox txtSpd 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Text            =   "1"
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton btnExec 
         Caption         =   "&OK"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   4
         Top             =   60
         Width           =   975
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   0
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
         MICON           =   "frmWMP.frx":000C
         PICN            =   "frmWMP.frx":0028
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   6
         Left            =   375
         TabIndex        =   1
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
         MICON           =   "frmWMP.frx":013D
         PICN            =   "frmWMP.frx":0159
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   2
         ToolTipText     =   "Song Info (from the web)"
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
         MICON           =   "frmWMP.frx":04AB
         PICN            =   "frmWMP.frx":04C7
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnTool 
         Height          =   375
         Index           =   1
         Left            =   1215
         TabIndex        =   3
         ToolTipText     =   "Find lyrics (from the web)"
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
         MICON           =   "frmWMP.frx":05A0
         PICN            =   "frmWMP.frx":05BC
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Speed: (0.5 ~ 16)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   10
         Top             =   150
         Width           =   1305
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4815
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8493
      _cy             =   7223
   End
End
Attribute VB_Name = "frmWMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CurrentlyOpenFile As String
Public MEDOTTAG As Integer

Public myTag As String 'the tag used by the tabs

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    If Index = 0 Then WMP.settings.Rate = Val(txtSpd.Text)
End Sub

Private Sub btnTool_Click(Index As Integer)
    On Error Resume Next
    With frmMain
        Select Case Index
            Case 0
                .titMediaViewSearchSong_Click (0)
            Case 1
                .titMediaViewSearchSong_Click (1)
            Case 4
                .titTextFileOpen_Click
            Case 6
                FileProperties CurrentlyOpenFile
        End Select
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls

    frmMain.titMedia.Visible = True
    frmMain.COF = Me.CurrentlyOpenFile
    SStatus WMP.URL, vbInformation
    
    'SyncTabs
    SyncCaptionEx Int(Me.Tag), Me.Caption
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titMedia.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next

    'browser index
    Dim I As Integer
    BRWIndex = BRWIndex + 1
    Me.Tag = BRWIndex
    I = frmMain.Tb.AddTab(Me.Caption)
    frmMain.Tb.TabTag(I) = Me.Tag
    'browser index


    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    WMP.stretchToFit = GetSet("Media_Stretch", "1")
    WMP.uiMode = IIf(GetSet("Media_Controls", "1") = "1", "full", "none")
    
    EventSound "WinOpen"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim I As Long
    I = IIf(picSet.Visible, picSet.Height, 0)
    WMP.Move 0, I, Me.ScaleWidth - IIf(pic.Visible, pic.Width, 0), Me.ScaleHeight - I
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
    EventSound "WinClose"
    Me.Tag = "CLOSED" 'suppose this is a tag for not marking the browser on the tabs
    SyncTabs
    If GetSet("Sync_PSM", "0") = "1" Then Call CMD6("psm ")
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    Dim K As String, L As String
    Dim G As Integer
    
    G = Int(GetSet("OnlyOnePlayer", "1"))

    If G > 0 Then
        Dim FRM As Form
        For Each FRM In Forms
            If TypeOf FRM Is frmWMP Then
                If G = 1 Then
                    Dim CTL As Control
                    For Each CTL In FRM
                        CTL.Controls.Pause
                    Next
                ElseIf G = 2 Then
                    Unload FRM
                End If
            End If
        Next
    End If
    
    WMP.URL = AddRecentItem(TheFN)
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    SyncCaptionEx Int(Me.Tag), Me.Caption
    
    PSMThis
    
    SaveSet "Media_Last", TheFN
    
    CurrentlyOpenFile = TheFN
    Me.Show
    SStatus Me.Name & " opened " & TheFN, vbInformation
        
End Function

Private Sub Pic_Resize()
    On Error Resume Next
    WmpPL.Move 0, 0, pic.Width, pic.Height
End Sub

Private Sub WMP_CurrentItemChange(ByVal pdispMedia As Object)
    On Error Resume Next
    Dim G As Long
    Dim TheFN As String
    'play count
    G = 0
    TheFN = WMP.currentMedia.sourceURL
    G = CLng(ReadINI(PathOnly(TheFN), TrimFileExt(FileNameOnly(TheFN)), FindPath(App.Path, App.EXEName & " playcounts.ini")))
    WriteINI PathOnly(TheFN), TrimFileExt(FileNameOnly(TheFN)), CStr(G + 1), FindPath(App.Path, App.EXEName & " playcounts.ini"), False
End Sub

Private Sub WMP_MediaChange(ByVal Item As Object)
    On Error Resume Next
    PSMThis
    CCaption Item.Name, Me
End Sub

Private Sub WMP_MediaError(ByVal pMediaObject As Object)
    On Error Resume Next
    SStatus "Error when trying to play " & WMP.URL, vbCritical
End Sub

Public Sub PSMThis()
    On Error Resume Next
    Dim K As String
    If GetSet("Sync_PSM", "0") = "1" Then 'sync MSN PSM
        K = GetSet("PSM", "%c %p - %n")
        K = Replace(K, "%n", WMP.currentMedia.Name)
        K = Replace(K, "%t", Now())
        K = Replace(K, "%a", WMP.URL)
        K = Replace(K, "%f", FileNameOnly(WMP.URL))
        K = Replace(K, "%s", Round(Val(FileLen(WMP.URL)) / 1024 / 1024, 2) & " MB")
        K = Replace(K, "%l", WMP.currentMedia.durationString)
        K = Replace(K, "%c", App.CompanyName)
        K = Replace(K, "%p", App.ProductName)
        Call CMD6("psm " & K)
    End If
End Sub

Private Sub WMP_PlayStateChange(ByVal NewState As Long)
    On Error Resume Next
    Dim G As Integer
    
    Select Case NewState
        Case 1 'stopped
        Case 2 'paused
        Case 3 'playing
            G = Int(GetSet("OnlyOnePlayer", "1"))
            If G = 1 Then
                Dim FRM As Form
                For Each FRM In Forms
                    MEDOTTAG = 1
                    If TypeOf FRM Is frmWMP Then
                        If FRM.MEDOTTAG <> 1 Then
                            FRM.WMP.Controls.Pause
                        End If
                    End If
                Next
            End If
    End Select
    MEDOTTAG = 0
End Sub

Private Sub WmpPL_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    WmpPL.Clear
    Dim I As Long
    For I = 1 To Data.Files.Count Step 1
        WMP.currentPlaylist.appendItem WMP.newMedia(Data.Files(I))
    Next
    For I = 1 To WMP.currentPlaylist.Count Step 1
        WmpPL.AddItem WMP.currentPlaylist.Item(I - 1).sourceURL
    Next
End Sub
