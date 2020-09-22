Attribute VB_Name = "ModCTl"
Option Explicit

Public FavStr As String

Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Public Declare Function IsThemeActive Lib "uxtheme.dll" () As Boolean
Public Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SHBrowseForFolder Lib "SHELL32" (lpBI As BrowseInfo) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SHGetPathFromIDList Lib "SHELL32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHShutDownDialog Lib "SHELL32" Alias "#60" (ByVal YourGuess As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Public Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long

Public Const DefaultSearchURL As String = "http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
Public Const SoftwareHomePage As String = "http://thinc.no-ip.info"
Public Const DefaultSearchAgent As String = "Google"
Public Const DefaultFilterString As String = "Search"
Public Const DefaultTmpFileName As String = "tempfile.tmp"
Public Const PathAbbrev As String = "..\"
Public Const LB_ITEMFROMPOINT = &H1A9
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const MAX_PATH = 260
Public Const DefaultSkinFile As String = "{app}\skin.ini"
Public Const defaultUpdateURL As String = "http://thinc.myvnc.com/vb/texp/update.ini"


'Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000


Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

Public BRWIndex As Integer 'stores the index of browsers (because the .tag is used)

Public Function UserName() As String
On Error Resume Next
    Dim lpBuffer As String
    Dim J
    lpBuffer = Space$(255)
    GetUserName lpBuffer, Len(lpBuffer)
        J = InStr(lpBuffer, Chr$(0))
    If J > 0 Then UserName = Left$(lpBuffer, J - 1)
End Function

Public Sub SkinFormEx(Which As Form)
    On Error Resume Next
    Dim A As Control
    Dim B As Integer, E As Integer, F As Integer, I As Integer
    Dim C As String
    Dim D As Boolean

    'B = Val(GetSet("CTL_Flatten", "0")) 'moved here to increase speed
    B = IIf(IsAppThemed And IsThemeActive, 0, 1)
'    C = GetSet("Lang")
'    E = Val(GetSet("CTL_FontSize"))
    F = Val(GetSet("CTL_ShowName"))
    
'    If Len(C) > 0 Then SkinForm Which, C 'language
    Which.IMGbkg.Move 0, 0, Which.ScaleWidth, Which.ScaleHeight 'default background handler
            
    If B <> 0 Or F <> 0 Then 'if user has something E <> 0 Or
        For Each A In Which
            If Len(A.Name) = 0 Then Exit For

            If B = 1 Then
                If TypeOf A Is CommandButton Then BTFlat A
                CtlFlat A
            End If

            If F = 1 Then
                I = A.Index
                If Err Then
                    A.ToolTipText = A.Name
                Else
                    A.ToolTipText = A.Name & " " & A.Index
                End If
                Err.Clear
            End If

            If A.Font <> "Marlett" And A.Font <> "Webdings" And Left$(A.Font, 9) <> "Wingdings" Then
                A.Font = GetSet("Font", "MS Shell Dlg") 'Fonts, replace all except pictures
            End If
            If E > 7 And E <= 72 Then
                A.FontSize = E 'size controls only if a value is present
            End If
        Next
    End If
    If B = 1 Then FormFlat Which

    DropShadow Which.hWnd
End Sub

Public Sub BTFlat(bt As CommandButton)
    On Error Resume Next
        If GetWindowLong&(bt.hWnd, -16) And &H8000& Then Exit Sub
        SetWindowLong bt.hWnd, -16, GetWindowLong&(bt.hWnd, -16) Or &H8000&
        bt.Refresh
End Sub

Public Sub CtlFlat(CL As Control)
    On Error Resume Next
        CL.Appearance = 0   'flat
        CL.BackColor = frmMain.picBrw.BackColor  'for cham buttons, and they change backcolor to the same as the container
        CL.ColorScheme = 2 'for cham buttons only
        CL.BackOver = frmMain.picBrw.BackColor 'for cham buttons only
End Sub

Public Sub FormFlat(Which As Form)
    On Error Resume Next
    Which.Appearance = 0
    Which.BackColor = &H8000000F 'looks more natural
End Sub

Public Sub LoadSearchProvider()
    On Error Resume Next
    frmMain.txtSearch.Text = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
End Sub

Function FileText(ByVal FN As String) As String
    Dim handle As Integer
    If Len(Dir$(FN)) = 0 Then FileText = ""
    handle = FreeFile
    Open FN For Binary As #handle
    ' read the string and close the file
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

Public Function BrowseForFolder(Owner As Long, Optional szTitle As String = "Select Folder...") As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo

    With tBrowseInfo
        .hwndOwner = Owner
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    End If
End Function

Public Function AF() As Form
    On Error Resume Next
    Dim K As Form
    'set
    
    Set K = frmMain.ActiveForm 'returns active form... in a shorter syntax in the main program.
    
    Set AF = K
End Function

Public Function DisplayFilter() As String
    On Error Resume Next
    DisplayFilter = DefaultFilterString & "..."
End Function

Public Function DownloadFile(URL As String, Optional SaveAsFile As String) As String
    On Error Resume Next
    If Len(SaveAsFile) = 0 Then SaveAsFile = FindPath(GetTempDir, DefaultTmpFileName)
    URLDownloadToFile 0, URL, SaveAsFile, 0, 0
    DownloadFile = SaveAsFile
End Function

Public Sub LoadRecents()
    On Error Resume Next
    Dim I As Integer
    Dim J As String, K As String
    With frmMain
        For I = 0 To 9 Step 1
            J = GetSet("Recent" & I)
            K = FileNameOnly(J) 'TrimFileNameLOL(J)
            If Len(K) > 50 Then K = Left$(K, 50) & "..."
            .titLERecentFilesArray(I).Caption = K 'TrimFileNameLOL(J, , , IIf(InStr(1, J, "/") > 0, "/", "\"))
            .titLERecentFilesArray(I).Tag = J
            .titLERecentFilesArray(I).Visible = (Trim$(.titLERecentFilesArray(I).Tag) <> "")
        Next
    End With
End Sub

Public Function AddRecentItem(WhatFileName As String) As String
    On Error Resume Next
    Dim I As Integer
    Dim A As String
        For I = 9 To 0 Step -1
            If GetSet("Recent" & I) = WhatFileName Then
                AddRecentItem = WhatFileName
                Exit Function
            End If
        Next
        For I = 9 To 1 Step -1 'end with 1! the first one is discarded ah mah...
            A = GetSet("Recent" & I - 1)
            If Len(A) > 0 Then
                SaveSet "Recent" & I, A
            End If
        Next I
        SaveSet "Recent0", WhatFileName
        LoadRecents
        AddRecentItem = WhatFileName
End Function

Public Sub LoadRecentFolders()
    On Error Resume Next
    Dim I As Integer, L As Integer
    Dim J As String, K As String
    With frmMain
        For I = 0 To 9 Step 1
            J = GetSet("RecentF" & I)
            For L = 9 To 0 Step -1
                    If .titLERecentFoldersArray(L).Tag = J Then
                        GoTo OhImHere
                    End If
            Next
            K = FileNameOnly(J) 'TrimFileNameLOL(J) ', , True)
            If Len(K) > 50 Then K = Left$(K, 50) & "..."
            .titLERecentFoldersArray(I).Caption = K 'TrimFileNameLOL(J, , True)
            .titLERecentFoldersArray(I).Tag = J
OhImHere:
            .titLERecentFoldersArray(I).Visible = (Trim$(.titLERecentFoldersArray(I).Tag) <> "")
        Next
    End With
End Sub

Public Function AddRecentFolder(WhatFolderName As String) As String
    On Error Resume Next
    Dim A As String
    Dim I As Integer ', J As Integer
        For I = 9 To 0 Step -1
            If GetSet("RecentF" & I) = WhatFolderName Then
                AddRecentFolder = WhatFolderName
                Exit Function
            End If
        Next
        For I = 9 To 1 Step -1 'end with 1! the first one is discarded ah mah...
            A = GetSet("RecentF" & I - 1)
            If Len(A) > 0 Then
                SaveSet "RecentF" & I, A
            End If
        Next I
        SaveSet "RecentF0", WhatFolderName
        LoadRecentFolders
        AddRecentFolder = WhatFolderName
End Function

Public Function MyMsgBox(Prompt As String, SaveNum As Integer, Optional MSGStyle As VbMsgBoxStyle = vbOKOnly, Optional titLE As String, Optional HideCheckBox As Boolean) As VbMsgBoxResult
    On Error Resume Next
    MyMsgBox = frmInputMsg.MyMsgBoxEx(Prompt, SaveNum, MSGStyle, titLE, HideCheckBox) 'wrapper
End Function

Public Function ParsedAddy(WhatNow As String) As String
    On Error Resume Next
    Dim Cmd As String, SRC As String
    If InStr(1, WhatNow, " ") <= 0 Then
        ParsedAddy = WhatNow
    Else
        Cmd = LCase$(Trim$(Left$(WhatNow, InStr(1, WhatNow, " "))))
        SRC = Mid$(WhatNow, InStr(1, WhatNow, " ") + 1)
        Select Case Cmd
            Case "g", "google"
                ParsedAddy = "http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
            Case "vb", "vb6"
                ParsedAddy = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&txtCriteria=%s&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&lngWId=1"
            Case "i", "image"
                ParsedAddy = "http://images.google.com/images?q=%s"
            Case "u", "youtube"
                ParsedAddy = "http://www.youtube.com/results?search_query=%s"
            Case "isohunt"
                ParsedAddy = "http://www.isohunt.com/torrents/?ihq=%s"
            Case "lucky", "gl", "feelinglucky", "imfeelinglucky", "imlucky"
                ParsedAddy = "http://www.google.com/search?hl=en&q=%s&btnI"
            Case "wiki"
                ParsedAddy = "http://en.wikipedia.org/wiki/%s"
            Case "stock"
                Dim M As String
                M = FillString(SRC, 4, "0", 0)
                ParsedAddy = "http://hk.finance.yahoo.com/q?s=" & M & ".HK"
        End Select
        ParsedAddy = Replace(ParsedAddy, "%s", SRC)
    End If
End Function

Public Function FavsPath() As String
    On Error Resume Next
    FavsPath = GetSet("FavsPath", "C:\Documents and Settings\" & UserName & "\Favorites")
End Function

Public Function OnTop(TheHwnd As Long, TrueOrFalse As Boolean)
    On Error Resume Next
    SetWindowPos TheHwnd, IIf(TrueOrFalse, -1, -2), 0, 0, 0, 0, 3 '&H1 Or &H10 Or &H2 Or &H40
End Function

Public Function SyncTabs()
    On Error Resume Next
    Dim K As String
    frmMain.Tb.RemoveAllTabs
    Dim FRM As Form
    For Each FRM In Forms
        If FRM.MDIChild = True Then
            If Not TypeOf FRM Is frmMain Then
                With frmMain.Tb
                    If FRM.Tag <> "CLOSED" Then
                        .AddTab First(FRM.Caption)
                        .TabTooltip(.TabUBound) = .TabCaption(.TabLBound)
                        .TabTag(.TabUBound) = FRM.Tag
                        .ActiveTab = Val(AF.Tag)
                    End If
                End With
            End If
        End If
    Next
End Function

Public Function First(What As String, Optional NumberOfChars As Integer = 15) As String
    On Error Resume Next
    If Len(What) >= NumberOfChars Then
        First = Left$(What, NumberOfChars) & "..."
    Else
        First = What
    End If
End Function

Public Function SpecialFolder(ByVal folder_number As Long) As String
    On Error Resume Next
    Dim Path As String
    Path = Space$(MAX_PATH)
    If SHGetSpecialFolderPath(0, Path, folder_number, False) Then
        SpecialFolder = Left$(Path, InStr(Path, Chr$(0)))
    End If
    SpecialFolder = TrimNull(SpecialFolder)
End Function

Public Function FileProperties(FN As String)
    On Error Resume Next
    Dim A As String, B As String
    A = FN
    
    If Len(A) = 0 Then
        MyMsgBox "You must save this file to see the file properties.", 8, , "File Info", True
        Exit Function
    End If
    
    If GetAttrib(A, vbArchive) = True Then B = B & "Archived"
    If GetAttrib(A, vbCompressed) = True Then B = B & ", Compressed"
    If GetAttrib(A, vbDirectory) = True Then B = B & ", Directory"
    If GetAttrib(A, vbHidden) = True Then B = B & ", Hidden"
    If GetAttrib(A, vbNormal) = True Then B = B & ", Normal"
    If GetAttrib(A, vbReadOnly) = True Then B = B & ", Read Only"
    If GetAttrib(A, vbTemporary) = True Then B = B & ", Temporary"
    If GetAttrib(A, vbVolume) = True Then B = B & ", Volume"
    If Left$(B, 2) = ", " Then B = Mid$(B, 3)
    MyMsgBox A & vbCrLf & vbCrLf & _
                    B & vbCrLf & vbCrLf & _
                    Round(Val(FileLen(A)) / 1024 / 1024, 2) & " MB", 8, , "File Info", True
End Function

Public Function doUD() As String
    On Error Resume Next
    Dim K As String
    
    K = GetSet("UpdaterURL", defaultUpdateURL)
    DeleteUrlCacheEntry K 'clear it
    K = DownloadFile(K) 'INI file
    
    If Val(ReadINI("Updater", "Major", K)) > App.Major Then
        doUD = ReadINI("Updater", "URL", K)
    ElseIf Val(ReadINI("Updater", "Major", K)) = App.Major Then 'version is the same, look in minor
        If Val(ReadINI("Updater", "Minor", K)) > App.Minor Then
            doUD = ReadINI("Updater", "URL", K)
        End If
    End If
End Function

Function UpdateMe(Where As String)
    On Error Resume Next
    Dim Response As VbMsgBoxResult
    
    If MsgBox("Do you want to download a new version of " & App.ProductName & " from:" & vbCrLf & vbCrLf & Where, vbYesNo + vbQuestion) = vbNo Then Exit Function
    
    With cmndlg
        .filefilter = "Application(*.exe)|*.exe"
        .dialogtitle = "Where do you want to save the file to?"
        SaveFile
        If Len(.FileName) = 0 Then Exit Function
        
        If LCase$(Right$(.FileName, 4)) <> ".exe" Then .FileName = .FileName & ".exe"
        
        Dim K As String
        K = DownloadFile(Where, .FileName)
        
        MsgBox "Download complete!" & vbCrLf & vbCrLf & K & vbCrLf & vbCrLf & "You can start using the new version by opening the file you just downloaded.", vbInformation
    End With
End Function

Function SyncCaptionEx(MyIndex As Integer, Cpn As String)
    On Error Resume Next
    Dim I As Integer
    With frmMain.Tb
        For I = 1 To .TabUBound Step 1
            If .TabTag(I) = MyIndex Then
                        .TabCaption(I) = First(Cpn)
                        .TabTooltip(I) = Cpn
                        .ActiveTab = I
                Exit For
            End If
        Next
    End With
End Function















Public Function DSA(Index As Integer) 'dont show again collection


'Do Not Show again UBound: Use 19

    On Error Resume Next
    Dim K As String
    Select Case Index
        Case 1
            K = "You are in sandbox mode - this setting will not be changed. You will need to Edit the INI manually in the preferences window."
        Case 2
            K = "To start using " & App.ProductName & ", please navigate to a file with the browser bar."
        Case 3
            If Dir(FindPath(App.Path, "ese.exe")) <> "" Then Exit Function
            K = App.ProductName & " will run this code only if the component ""ESE.exe"" is available." & vbCrLf & vbCrLf & "You can get a copy of the component at " & SoftwareHomePage & "."
        Case 4
            K = App.ProductName & " will download this file and may stop responding while downloading a large file."
        Case 5
            K = "The sandbox mode stops your settings file from being changed or written into."
        Case 6
            K = "This will also change your start up form tag to Media."
        Case 7
            If Dir(GetSet("PSMLoc")) <> "" Then Exit Function
            K = App.ProductName & " will run this code only if the messenger plugin is available." & vbCrLf & vbCrLf & "You can get a copy of the component at " & SoftwareHomePage & "."
        Case 8 'file info diag
            DoEvents
        Case 9
            K = App.ProductName & " may not able to start after this." & vbCrLf & "To reset this:" & vbCrLf & "- delete " & App.EXEName & ".exe.manifest from the directory." & vbCrLf & "- come back to this window to uncheck this checkbox."
        Case 10
            If Dir(FindPath(App.Path, "teext.exe")) <> "" Then Exit Function
            K = App.ProductName & " will run this code only if the component ""TEExt.exe"" is available." & vbCrLf & vbCrLf & "You can get a copy of the component at " & SoftwareHomePage & "."
        Case 11
            K = "This will lag a lot if you open a lot of files." & vbCrLf & vbCrLf & "Just to let you know." & vbCrLf & vbCrLf & "If you think that's the case, save some RAM and click OK."
        Case 12
            K = App.ProductName & " is set not to allow termination of the program for this session." & vbCrLf & vbCrLf & "Please find alternative ways of doing this yourself."
        Case 13
            K = "Closing this application will also close all the windows inside it." & vbCrLf & vbCrLf & "If you have made the mistake, please remember it and do not do it again."
        '14-15 used
        Case 16
            K = "Job completed."
        Case 17
            'this is old and not used
            K = "Why not make " & App.CompanyName & " nicer by using a skin?" & vbCrLf & vbCrLf & "You can use a SKIN.INI to enhance the appeal."
    End Select
    MyMsgBox K, Index
End Function
