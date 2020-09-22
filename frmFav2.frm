VERSION 5.00
Begin VB.Form frmFav2 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Favorites"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFav2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnExec 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton btnPrefs 
      Caption         =   "&Tools"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   4020
      Width           =   4455
   End
   Begin VB.FileListBox filFav 
      Height          =   2820
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.CommandButton btnExec 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Open Favorites"
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
      Left            =   600
      TabIndex        =   6
      Top             =   158
      Width           =   1590
   End
   Begin VB.Image IMG 
      Height          =   360
      Left            =   120
      Picture         =   "frmFav2.frx":000C
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblBMInfo 
      BackStyle       =   0  '³z©ú
      Caption         =   "Start typing at the box below"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   5415
   End
   Begin VB.Image IMGbkg 
      Height          =   4455
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
   Begin VB.Menu titCat 
      Caption         =   "Category Ppup"
      Visible         =   0   'False
      Begin VB.Menu titPrefs 
         Caption         =   "Preferences..."
      End
      Begin VB.Menu titCatEditGotoFolder 
         Caption         =   "Go to folder"
      End
      Begin VB.Menu titCatChangeLoc 
         Caption         =   "Show from folder..."
      End
      Begin VB.Menu titS02 
         Caption         =   "-"
      End
      Begin VB.Menu titBackupTo 
         Caption         =   "Back up..."
      End
      Begin VB.Menu titLoadFrom 
         Caption         =   "Restore..."
      End
   End
   Begin VB.Menu FileListPopup 
      Caption         =   "FileListPopup"
      Visible         =   0   'False
      Begin VB.Menu titFLOpenBM 
         Caption         =   "Open"
      End
      Begin VB.Menu titFLRename 
         Caption         =   "Give a new name..."
      End
      Begin VB.Menu titFLEditURL 
         Caption         =   "Edit URL..."
      End
      Begin VB.Menu titFLDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmFav2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WhatWasIt As String

Dim FavURL As String
Dim FavOpenWith As Integer

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    Dim A As New frmTXT
    Dim B As New frmWMP
    Dim C As New frmIMG
    Dim D As New frmBRW
    
    If filFav.ListIndex < 0 Then filFav.ListIndex = 0
    
    FavAddyEx FindPath(filFav.Path, filFav.List(filFav.ListIndex)) ', lstCat.ListIndex
    
    Unload Me 'seems like it works only when focused form is frmMain, so yeah whatever
    
    If InStr(1, FavURL, "[criteria]") > 0 Then
            FavURL = Replace(FavURL, "[criteria]", InputBox("This favorite has a field to be filled [criteria]. Enter parameter here:"))
    End If
    
    If Len(FavURL) = 0 Then Exit Sub 'there's no point of going on.
    
    If Index > 0 Then
        Select Case FavOpenWith
            Case 0 'bookmarks
                D.LoadFile FavURL
            Case 1 'media
                B.LoadFile FavURL
            Case 2 'image
                C.LoadFile FavURL
            Case 3 'programs
                Shell FavURL, vbNormalFocus
            Case 4 'text
                A.LoadFile FavURL
        End Select
        SStatus App.ProductName & " opened " & FavURL, vbInformation
    End If
    
End Sub

Private Function FavAddyEx(WhichFile As String) As String ', Optional ForceOpenWith As Integer = 0
        Dim A As String
        A = ReadINI("DEFAULT", "BASEURL", WhichFile)
        If Len(A) > 0 Then
            FavAddyEx = A
        Else
            A = ReadINI("InternetShortcut", "URL", WhichFile)
            If Len(A) > 0 Then
                FavAddyEx = A
            End If
        End If
'        FavOpenWith = Val(ReadINI(App.ProductName, "OpenWith", WhichFile))
'        If ForceOpenWith <> 0 Then FavOpenWith = ForceOpenWith
        FavURL = FavAddyEx
End Function

Private Sub btnPrefs_Click()
    PopupMenu titCat, , btnPrefs.Left, btnPrefs.Top + btnPrefs.Height
End Sub

Private Sub filFav_Click()
    lblBMInfo.Caption = filFav.List(filFav.ListIndex) & _
        vbCrLf & FavAddyEx(FindPath(filFav.Path, filFav.List(filFav.ListIndex)))
End Sub

Private Sub filFav_DblClick()
    btnExec_Click 1
End Sub

Private Sub filfav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        filfav_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub filfav_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim Ix As Long
    Dim Mx As Long, My As Long
    If Button = 2 Then
        Mx = CLng(X / Screen.TwipsPerPixelX)
        My = CLng(Y / Screen.TwipsPerPixelY)
        Ix = SendMessage(filFav.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((My * 65536) + Mx))
        If Ix < filFav.ListCount Then
            filFav.Selected(Ix) = True
            PopupMenu FileListPopup, , Mx * Screen.TwipsPerPixelX + filFav.Left, My * Screen.TwipsPerPixelY + filFav.Top
        End If
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    InitCommonControls
    
    'F1.FadeIn
    
    If GetSet("FAV_Enable", "1") = "0" Then
        MsgBox "The favorites function has been disabled.", vbCritical
        Unload Me
        Exit Sub
    End If
    
End Sub

'Private Sub Form_Deactivate()
'    F1.FadeOut
'End Sub

Private Sub Form_Load()
    On Error Resume Next
    
'    F1.PrepareFade
    
    SkinForm Me
    SkinFormEx Me
    Me.Move frmMain.Left + (frmMain.Width - Me.Width) / 2, frmMain.Top + (frmMain.Height - Me.Height) / 2
    
'    lstCat.ListIndex = 0
'    lstCat_Click
    
    filFav.Path = GetSet("FAV_Bookmarks", FavsPath)
    
    EventSound "WinOpen"
        
    WhatWasIt = "Search Bookmarks..."
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    EventSound "WinClose"
End Sub

Private Sub lblBMInfo_DblClick()
    On Error Resume Next
    Clipboard.SetText lblBMInfo.Caption
    Beep
End Sub

'Private Sub lstCat_Click()
'    On Error Resume Next
'    Dim K(5) As String, J As String, L As String
'    Dim I As Integer
'
'    For I = 0 To lstCat.ListCount - 1 Step 1
'        J = lstCat.List(I)
'        If J = "Bookmarks" Then
'            K(I) = GetSet("FAV_" & J, FavsPath)
'        Else
'            L = FindPath(App.Path, "Favs")
'            MkDir L
'            L = FindPath(App.Path, "Favs\" & J)
'            MkDir L
'            K(I) = GetSet("FAV_" & J, L)
'        End If
'    Next
'    filFav.Path = K(lstCat.ListIndex)
'
'    txtSearch.Text = DefaultSearchStr
'
'    txtSearch.SetFocus
'End Sub

'Private Sub lstCat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'    If Button = 2 Then
'        lstCat_MouseMove Button, Shift, X, Y
'    End If
'End Sub
'
'Private Sub lstCat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'    Dim Ix As Long
'    Dim Mx As Long, My As Long
'    If Button = 2 Then
'        Mx = CLng(X / Screen.TwipsPerPixelX)
'        My = CLng(Y / Screen.TwipsPerPixelY)
'        Ix = SendMessage(lstCat.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((My * 65536) + Mx))
'        If Ix < lstCat.ListCount Then
'            lstCat.Selected(Ix) = True
'            PopupMenu titCat, , Mx * Screen.TwipsPerPixelX + lstCat.Left, My * Screen.TwipsPerPixelY + lstCat.Top, titCatNew
'        End If
'    End If
'End Sub

Private Sub titBackupTo_Click()
    On Error Resume Next
    Dim K As String
    Dim I As Integer
    K = BrowseForFolder(Me.hWnd, "Select a folder you want to back up your bookmarks to...")
    If Len(K) > 0 Then
        For I = 0 To filFav.ListCount - 1 Step 1
            FileCopy FindPath(filFav.Path, filFav.List(I)), FindPath(K, filFav.List(I))
            SStatus I + 1 & " out of " & filFav.ListCount & " completed", vbInformation
            SProgress CLng(I), , filFav.ListCount - 1
            DoEvents
        Next
        SStatus
        SProgress 0
        MsgBox "All favourites shown now are backed up to " & K & ".", vbInformation
    End If
End Sub

Private Sub titCatChangeLoc_Click()
    On Error Resume Next
    Dim K As String
    K = BrowseForFolder(Me.hWnd)
    If Len(K) > 0 Then SaveSet "FAV_" & filFav.List(filFav.ListIndex), K
    filFav.Refresh
End Sub

Private Sub titCatEditGotoFolder_Click()
    On Error Resume Next
    Dim J As String
    J = filFav.Path
    Shell "explorer " & J, vbNormalFocus
End Sub
'
'Private Sub titCatNew_Click()
'    On Error Resume Next
'    Dim J As String, K As String
'    J = InputBox("Please enter a name for this favorite." & vbCrLf & "This will also be your file name.", , frmMain.COF)
'    K = InputBox("Please enter a location for this favorite.", , frmMain.COF)
'
'End Sub

Private Sub titFLDelete_Click()
    On Error Resume Next
    Dim typOperation As SHFILEOPSTRUCT
    With typOperation
            .wFunc = &H3
            .pFrom = FindPath(filFav.Path, filFav.FileName)
            .fFlags = &H40
        End With
        SHFileOperation typOperation
    filFav.Refresh
End Sub

Private Sub titFLEditURL_Click()
    On Error Resume Next
    Dim K As String, TF As String
    If filFav.ListIndex < 0 Then Exit Sub 'if user still hasnt selected anything then there will be an error, so nah
    TF = FindPath(filFav.Path, filFav.FileName)
    K = InputBox("Edit URL in this favorite:", , FavAddyEx(TF))
    If Len(K) = 0 Then Exit Sub
    
    WriteINI "DEFAULT", "BASEURL", K, TF
    WriteINI "InternetShortcut", "URL", K, TF
    
End Sub

Private Sub titFLOpenBM_Click()
    On Error Resume Next
    btnExec_Click 1
End Sub

Private Sub titFLRename_Click()
    On Error Resume Next
    Dim K As String
    K = InputBox("Rename to: (please include extension)", , filFav.FileName)
    If Len(K) = 0 Then Exit Sub
    
    Name FindPath(filFav.Path, filFav.FileName) As FindPath(filFav.Path, K)
    filFav.Refresh
    
End Sub

Private Sub titLoadFrom_Click()
    On Error Resume Next
    Dim K As String, J As String
    Dim I As Integer
    K = BrowseForFolder(Me.hWnd, "Select a folder you want to restore your bookmarks from...")
    If Len(K) > 0 Then
        J = filFav.Path
        filFav.Path = K
        For I = 0 To filFav.ListCount - 1 Step 1
            FileCopy FindPath(K, filFav.List(I)), FindPath(J, filFav.List(I))
            SProgress CLng(I), , filFav.ListCount - 1
            SStatus I + 1 & " out of " & filFav.ListCount & " completed", vbInformation
            DoEvents
        Next
        SProgress 0
        SStatus
        filFav.Path = J
    End If
End Sub

Private Sub titPrefs_Click()
    On Error Resume Next
    frmPrefs.GoToTab 2
    frmPrefs.Show 1
    'MsgBox "Please open this dialog again to see any changes you made.", vbInformation
    Form_Load
'    F1.FadeIn 'Bug here.
    
End Sub

Private Sub txtSearch_Change()
    On Error Resume Next
    Dim K As String
    With txtSearch
        If .Text = WhatWasIt Then Exit Sub
        If .Text = "" Or .Text = DefaultSearchStr Then
            filFav.Pattern = "*"
        Else
            K = Replace(.Text, " ", "*")
            filFav.Pattern = "*" & K & "*"
        End If
        WhatWasIt = .Text
    End With
End Sub

Private Sub txtSearch_GotFocus()
    TBFocus txtSearch, True, DefaultSearchStr
End Sub

Private Sub txtSearch_LostFocus()
    TBFocus txtSearch, False, DefaultSearchStr
End Sub

Private Function DefaultSearchStr() As String
    On Error Resume Next
    DefaultSearchStr = DefaultFilterString & "..." '"Search " & lstCat.List(lstCat.ListIndex) & "..."
End Function
