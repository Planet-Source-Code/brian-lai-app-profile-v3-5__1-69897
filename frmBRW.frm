VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBRW 
   AutoRedraw      =   -1  'True
   Caption         =   "Browser"
   ClientHeight    =   6390
   ClientLeft      =   3060
   ClientTop       =   3420
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBRW.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '³Ì¤j¤Æ
   Begin SHDocVwCtl.WebBrowser BRW 
      CausesValidation=   0   'False
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8760
      ExtentX         =   15452
      ExtentY         =   9975
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8940
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8940
      Begin VB.TextBox cboAddress 
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   83
         Width           =   3735
      End
      Begin ProFile.CB btnBrw 
         Height          =   480
         Index           =   0
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Back"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":000C
         PICN            =   "frmBRW.frx":0028
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   480
         Index           =   1
         Left            =   450
         TabIndex        =   5
         ToolTipText     =   "Forward"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":073A
         PICN            =   "frmBRW.frx":0756
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   480
         Index           =   2
         Left            =   900
         TabIndex        =   6
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":0E68
         PICN            =   "frmBRW.frx":0E84
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   480
         Index           =   3
         Left            =   1350
         TabIndex        =   7
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":1596
         PICN            =   "frmBRW.frx":15B2
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   480
         Index           =   4
         Left            =   1800
         TabIndex        =   8
         ToolTipText     =   "Home"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":1CC4
         PICN            =   "frmBRW.frx":1CE0
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         CausesValidation=   0   'False
         Height          =   480
         Index           =   6
         Left            =   2700
         TabIndex        =   9
         ToolTipText     =   "Favorites"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":23F2
         PICN            =   "frmBRW.frx":240E
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnGo 
         CausesValidation=   0   'False
         Height          =   480
         Left            =   7800
         TabIndex        =   10
         ToolTipText     =   "Go"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":2B20
         PICN            =   "frmBRW.frx":2B3C
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.CB btnBrw 
         Height          =   480
         Index           =   5
         Left            =   2250
         TabIndex        =   11
         ToolTipText     =   "Zoom"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   847
         BTYPE           =   9
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":324E
         PICN            =   "frmBRW.frx":326A
         PICPOS          =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAddress 
         BackStyle       =   0  '³z©ú
         Caption         =   "A&ddress:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3165
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   113
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmBRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EventRunning As Boolean
Dim LastLogMsg As String
Public CurrentlyOpenFile As String
Dim YouCantChangeMyAddyBarTextNow As Boolean 'tag
Dim ShowTags As Boolean

Public myTag As String 'the tag used by the tabs

'Dim BrwGen As HTMLGenericElement
'Dim BrwHref As HTMLAnchorElement
Dim BrwEvent As IHTMLEventObj
Dim WithEvents BrwDoc As HTMLDocument
Attribute BrwDoc.VB_VarHelpID = -1
'Dim old_element As HTMLGenericElement

Private Declare Sub SHAutoComplete Lib "shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As Long) 'hinted by Juanito Dado Jr

Private Sub BRW_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    On Error Resume Next
    Form_Resize
    SetAddy BRW.LocationURL
    
    SyncCaption
        
    If GetSet("BRW_Log", "1") = "1" Then
        If BRW.LocationURL <> LastLogMsg Then
            cboAddress.BackColor = IIf(Left$(BRW.LocationURL, 5) = "https", RGB(222, 255, 140), RGB(255, 255, 255))
            Log BRW.LocationURL, True
            LastLogMsg = BRW.LocationURL
        End If
    End If
End Sub

Private Sub BRW_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Dim V As String, Z As String
    Dim I As Long
    Dim K() As String, L() As String
        
    EventRunning = True
    SetAddy BRW.LocationURL
    
    SyncCaption
    
    CurrentlyOpenFile = CStr(URL)
    
    If GetSet("BRW_Log", "1") = "1" Then
        If BRW.LocationURL <> LastLogMsg Then
            cboAddress.BackColor = IIf(Left$(BRW.LocationURL, 5) = "https", RGB(204, 255, 204), RGB(255, 255, 255))
            Log BRW.LocationURL, True
            LastLogMsg = BRW.LocationURL
        End If
    End If
    
    V = ReadINI(CStr(URL), "V", FindPath(App.Path, AutoFillINI)) 'field names
    K = Split(V, ",")
    Z = ReadINI(CStr(URL), "Z", FindPath(App.Path, AutoFillINI)) 'field values
    L = Split(Z, ",")
    
    For I = 0 To UBound(K) Step 1
        BRW.Document.All.Item(K(I)).Value = L(I)
    Next
    
    frmMain.ActiveForm.BRW.SetFocus
    Set BrwDoc = BRW.Document
    EventRunning = False
    
End Sub

Function SyncCaption()
    On Error Resume Next
    Dim I As Integer
    
    If Len(BRW.Document.titLE) > 0 Then
        CCaption BRW.Document.titLE, Me
    ElseIf Len(BRW.LocationURL) > 0 Then
        CCaption BRW.LocationURL, Me
    Else
        CCaption "Loading...", Me
    End If
    
    If Me.Caption = "Browser" Then Exit Function
    
    SyncCaptionEx Int(Me.Tag), Me.Caption
    
End Function

Private Sub BRW_FileDownload(Cancel As Boolean)
    On Error Resume Next
    If GetSet("BRW_Download", "1") = "0" Then Cancel = True
End Sub

Private Sub BRW_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    If BRW.busy = True Then Exit Sub 'Blocking onload popups
    
    If GetSet("Browser_AllowNewWindow", "1") = "1" Then
        If GetSet("NewBrowserInTab", "1") = "1" Then
            Dim F As New frmBRW
            Set ppDisp = F.BRW.object
            F.Show
        Else
            'default handler
        End If
    Else
        Cancel = True
    End If
End Sub

Private Sub BRW_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    SProgress Progress, 0, ProgressMax
End Sub

Private Sub BRW_StatusTextChange(ByVal Text As String)
    SStatus Text, vbInformation
End Sub

Private Sub BRW_TitleChange(ByVal Text As String)
    On Error Resume Next
    CCaption Text, Me
    SyncCaption
End Sub

Private Sub BRW_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    On Error Resume Next
    Unload Me
End Sub

Private Sub BrwDoc_onmousemove()
    On Error Resume Next
    Dim K As String
    If ShowTags Then
        Set BrwEvent = BrwDoc.parentWindow.event
        'K = BrwDoc.elementFromPoint(BrwEvent.clientX, BrwEvent.clientY).outerHTML
        K = BrwDoc.elementFromPoint(BrwEvent.clientX, BrwEvent.clientY).getAttribute("Movie")
'        If Len(K) = 0 Then
'            K = BrwDoc.elementFromPoint(BrwEvent.clientX, BrwEvent.clientY).ID
'        End If
'        If Len(K) = 0 Then
'            K = BrwDoc.elementFromPoint(BrwEvent.clientX, BrwEvent.clientY).tagName
'        End If
        SStatus K, vbInformation
    End If
End Sub

Public Sub btnBrw_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            BRW.GoBack
        Case 1
            BRW.GoForward
        Case 2
            BRW.Refresh
        Case 3
            BRW.stop
            CCaption BRW.LocationName, Me
        Case 4
            BRW.GoHome
        Case 5
            'BRW.GoSearch
            PopupMenu frmMain.titBrowserZoom, , btnBrw(5).Left, btnBrw(5).Top + btnBrw(5).Height
        Case 6
            PopupMenu frmMain.titBrowserP, , btnBrw(6).Left, btnBrw(6).Top + btnBrw(6).Height, frmMain.titBrowserPBMThis
            'frmMain.titBrowserFavorites_Click
    End Select
End Sub

Private Sub btnGo_Click()
    EventRunning = False
    cboAddress_KeyDown vbKeyReturn, 0
    EventRunning = True
End Sub

Private Sub cboAddress_Change()
    On Error Resume Next
    cboAddress.BackColor = IIf(Left$(BRW.LocationURL, 5) = "https", RGB(204, 255, 204), RGB(255, 255, 255))
End Sub

Private Sub cboAddress_DblClick()
    On Error Resume Next
    With cboAddress
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub cboAddress_GotFocus()
    On Error Resume Next
    YouCantChangeMyAddyBarTextNow = True
    cboAddress_DblClick
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If Shift = 4 Then cboAddress_DblClick
    If KeyCode = vbKeyReturn Then
        Select Case Shift
            Case 0 'none
                If EventRunning Then Exit Sub
                URLCorrect
                LoadFile ParsedAddy(cboAddress.Text)
                BRW.SetFocus
            Case 2 'ctrl
                LoadFile "http://www." & cboAddress.Text & ".com"
                BRW.SetFocus
        End Select
    End If
End Sub

Public Function URLCorrect()
    On Error Resume Next
    ''intelligently" corrects what program thinks is wrong
    Dim K As String
    Dim I As Integer
    I = Int(GetSet("URL_AutoCorrect", "1"))
    If I = 0 Then Exit Function 'if "do nothing" then really do nothing
    
    K = cboAddress.Text 'this is the URL
    
    K = StrSwaps(K, ".cmo", ".com", True)
    K = StrSwaps(K, ".cm", ".com", True)
    K = StrSwaps(K, ".co", ".com", True)
    K = StrSwaps(K, ".og", ".org", True)
    K = StrSwaps(K, "ww.", "www.", False)
    K = StrSwaps(K, "wwww.", "www.", False)
    K = StrSwaps(K, "qqq.", "www.", False)
    K = StrSwaps(K, "http:..", "http://", False)
    
    If K <> cboAddress.Text Then 'if we have a value changed
        If I = 2 Then
            cboAddress.Text = K
        Else
            If MsgBox("Do you mean " & vbCrLf & K & "?", vbQuestion + vbYesNo) = vbYes Then
                cboAddress.Text = K
            End If
        End If
    End If
End Function

Public Function StrSwaps(KHere As String, FromW As String, ToW As String, Optional OnTheRight As Boolean = True) As String
    On Error Resume Next
    StrSwaps = KHere
    If OnTheRight Then
        If LCase$(Right$(KHere, Len(FromW))) = FromW Then
            StrSwaps = Left$(KHere, Len(KHere) - Len(FromW)) & ToW
        End If
    ElseIf OnTheRight = False Then
        If LCase$(Left$(KHere, Len(FromW))) = FromW Then
            StrSwaps = ToW & Right$(KHere, Len(KHere) - Len(FromW))
        End If
    End If
End Function


Private Sub cboAddress_LostFocus()
    On Error Resume Next
    YouCantChangeMyAddyBarTextNow = False 'reset the tag
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    frmMain.titBrowser.Visible = True
    If GetSet("BRW_AutoFavsBarSwitch", "1") = "1" Then frmMain.GoToPath FavsPath, False
    frmMain.COF = Me.CurrentlyOpenFile
    
    'SyncTabs
    SyncCaptionEx Int(Me.Tag), Me.Caption
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    With frmMain
        .titBrowser.Visible = False
        If GetSet("BRW_AutoFavsBarSwitch", "1") = "1" Then .GoToPath GetSet("Recent_Path"), False
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim I As Integer
    EventSound "WinOpen"
    SkinForm Me
    SkinFormEx Me
    SHAutoComplete cboAddress.hWnd, &H0
    
    BRW.Silent = True
    Me.Icon = frmMain.Icon
    
    'browser index
    BRWIndex = BRWIndex + 1
    Me.Tag = BRWIndex
    I = frmMain.Tb.AddTab("Loading...")
    frmMain.Tb.TabTag(I) = Me.Tag
    'browser index
    
    ShowTags = (GetSet("ShowTags", "0") = "1")
    
    Select Case Val(GetSet("Browser_Init", "2"))
        Case 0
            BRW.Navigate2 "about:blank"
        Case 1
            BRW.Navigate2 GetSet("Browser_LastURL")
        Case 2
            BRW.GoHome
    End Select
    
    Form_Resize
    'DSA 17
End Sub

Private Sub BRW_DownloadComplete()
    On Error Resume Next
    
    EventRunning = True
    SetAddy BRW.LocationURL
    EventRunning = False
    SyncCaption

End Sub

Public Function FavAddy(WhichFile As String, Optional SignalFromFavForm As Boolean) As String
    On Error Resume Next

        Dim A As String
        A = ReadINI("DEFAULT", "BASEURL", WhichFile)
        If Len(A) > 0 Then
            FavAddy = A
        Else
            A = ReadINI("InternetShortcut", "URL", WhichFile)
            If Len(A) > 0 Then
                FavAddy = A
            End If
        End If
End Function

Private Sub BRW_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Dim I As Integer
    Dim bFound As Boolean
    CCaption BRW.Document.titLE, Me
    EventRunning = True
    
    Form_Resize
    SetAddy BRW.LocationURL
    If GetSet("BRW_Log", "1") = "1" Then
        If BRW.LocationURL <> LastLogMsg Then
            Log BRW.LocationURL, True
            LastLogMsg = BRW.LocationURL
        End If
    End If
    
    EventRunning = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    SaveSet "Browser_LastURL", BRW.LocationURL
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With BRW
        .Move 0, 480, Me.ScaleWidth, Me.ScaleHeight - 480 '480 being the top
    End With
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    Dim G As Integer
    BRW.Navigate2 AddRecentItem(TheFN)
    
    G = Int(GetSet("OnlyOneBrowser", "0"))

    If G > 0 Then
        Dim FRM As Form
        For Each FRM In Forms
            If TypeOf FRM Is frmBRW Then Unload FRM
        Next
    End If
    
    CurrentlyOpenFile = TheFN
    CCaption FileNameOnly(TheFN), Me 'TrimFileNameLOL(TheFN), Me
    Me.CurrentlyOpenFile = TheFN
    Me.Show
    Form_Resize
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmMain.LastClosedURL = BRW.LocationURL
    Form_Deactivate
    EventSound "WinClose"
    Me.Tag = "CLOSED" 'suppose this is a tag for not marking the browser on the tabs
    
    SyncTabs
End Sub

Function SetAddy(WhatText As String)
    On Error Resume Next
    If YouCantChangeMyAddyBarTextNow = False Then
        cboAddress.Text = WhatText
    End If
End Function

Private Sub picAddress_Resize()
    On Error Resume Next
    btnGo.Move picAddress.Width - btnGo.Width
    cboAddress.Width = picAddress.Width - cboAddress.Left - btnGo.Width
End Sub
