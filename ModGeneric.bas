Attribute VB_Name = "ModGeneric"
Option Explicit

Public Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long
Public Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = (-26)

Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function PathStripPath Lib "shlwapi" Alias "PathStripPathA" (ByVal pPath As String) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public hFile As Long

Public Enum TimeId
    Created = 0
    Modified = 1
    Accessed = 2
End Enum

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private LocalDate                   As FILETIME

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private FileProps As WIN32_FIND_DATA

Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public SysDate                     As SYSTEMTIME

'Constants
Public Const SB_HORZ = 0 'Horizontal Scrollbar
Public Const SB_VERT = 1 'Vertical Scrollbbar
Public Const SB_BOTH = 3 'Both ScrollBars

Public Enum vbFileAttributes
  vbNormal = 0         ' Normal
  vbReadOnly = 1       ' Read-only
  vbHidden = 2         ' Hidden
  vbSystem = 4         ' System file
  vbVolume = 8         ' Volume label
  vbDirectory = 16     ' Directory or folder
  vbArchive = 32       ' File has changed since last backup
  vbTemporary = &H100  ' 256
  vbCompressed = &H800 ' 2048
End Enum
   
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Sub Pause(PauseTime As Long)
    On Error Resume Next
    Dim StartTime As Long
    
    StartTime = Timer
    Do While Timer - StartTime < PauseTime
        DoEvents
    Loop
End Sub
'its as simple as that


Public Function GetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes) As Boolean
  If (LenB(sFileSpec) <> 0) Then
    GetAttrib = (GetAttributes(sFileSpec) And Attrib) = Attrib
  End If
End Function

Public Sub SetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes, Optional fTurnOff As Boolean)
  If (LenB(sFileSpec) <> 0) Then
    If (Attrib = vbNormal) Then
      SetAttributes sFileSpec, vbNormal
    ElseIf fTurnOff Then
      SetAttributes sFileSpec, GetAttributes(sFileSpec) And (Not Attrib)
    Else
      SetAttributes sFileSpec, GetAttributes(sFileSpec) Or Attrib
    End If
  End If
End Sub

Function ControlNum(YourVal As Long, Optional MinVal As Long = 0, Optional MaxVal As Long = 100) As Long
    On Error Resume Next
    If YourVal < MinVal Then YourVal = MinVal
    If YourVal > MaxVal Then YourVal = MaxVal
    ControlNum = YourVal
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function

Public Function GetString(Which As String, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetString = Arr(SectionNo)
End Function

Public Sub SStatus(Optional What As String = "Ready", Optional MyType As VbMsgBoxStyle)
    On Error Resume Next
    Dim K As Long
    
    Select Case MyType
        Case vbCritical
            K = RGB(255, 0, 0)
        Case vbInformation
            K = RGB(0, 0, 64)
        Case vbExclamation
            K = RGB(0, 64, 64)
        Case Else
            K = RGB(0, 0, 0)
    End Select
        
    With frmMain.lblStatus
        If .Caption = What Then Exit Sub
        .Caption = What
        .ForeColor = K
    End With
End Sub

Public Sub SProgress(Value As Long, Optional ValueMin As Long = 0, Optional ValueMax As Long = 100)
    On Error GoTo errrr
    Dim B As Double
    With frmMain
        If Value <= ValueMin Or Value > ValueMax Then
            .picProgress.Visible = False
        Else
            .picProgress.Visible = True
            .Prg1.Height = .picProgress.Height
            .Prg1.Value = Value
            .Prg1.Max = ValueMax
        End If
        'this part is for SStatus
        If .picProgress.Visible Then
            .lblStatus.Left = .picProgress.Left + .picProgress.Width + 60
        Else
            .lblStatus.Left = 30
        End If
        '/this part is for SStatus
    End With
    Exit Sub
errrr:
End Sub

'Returns only file name
Public Function TrimFileNameLOL(FromWhat As String, Optional ForceLong As Boolean = False, _
                                                                               Optional AddDotDotDot As Boolean = False, _
                                                                               Optional Separator As String = "\") As String
    On Error Resume Next
    'obsolete function
    TrimFileNameLOL = FileNameOnly(FromWhat)
End Function

Public Function FileNameOnly(sPath As String, Optional RemoveExt As Boolean) As String
    On Error Resume Next
    Dim K As String
    K = sPath
    If GetSet("ShowFullPaths") = "1" Then
        FileNameOnly = K
    Else
        Call PathStripPath(K)
        FileNameOnly = TrimNull(K)
    End If
    If Len(FileNameOnly) > 0 Then
            If RemoveExt Then
                FileNameOnly = Left$(FileNameOnly, InStrRev(FileNameOnly, ".") - 1)
            End If
    End If
End Function

'Returns only path
Public Function PathOnly(FromWhat As String) As String
    On Error Resume Next
    PathOnly = Left$(FromWhat, InStrRev(FromWhat, "\") - 1)
End Function

Public Sub TXTFileSave(Text As String, filepath As String)
    On Error Resume Next
    Dim F As Integer
    F = FreeFile
    Open filepath For Output As #F
        Print #F, Text
    Close #F
    Exit Sub
End Sub

Public Function MyVer() As String
    On Error Resume Next
    Dim AppRevision As Integer, AppMinor As Integer, AppMajor As Integer
    
    AppRevision = App.Revision
    AppMinor = App.Minor
    AppMajor = App.Major
    
    If AppRevision >= 10 Then
        AppMinor = AppMinor + 1
        AppRevision = AppRevision - 10
    End If
    
    If AppRevision > 0 Then AppMinor = AppMinor + 1
    
    If AppMinor >= 10 Then
        AppMajor = AppMajor + 1
        AppMinor = AppMinor - 10
    End If
    
'    MyVer = "V." & AppMajor & "." & AppMinor & IIf(AppRevision > 0, " Beta", "")
    MyVer = AppMajor & "." & AppMinor & IIf(AppRevision > 0, " beta", "")
    
End Function

Public Function FileExists(TheFN As String) As Boolean
    'does not work for hidden or whtever objects.
    'edited from copied code. his remarks:
    '#          I wish that it will help sumebody              #
    '#     I think this is one of the easiest way to do it     #
    Dim Var1 As String       'Variable for this module.
    On Error GoTo NotThere       'Simulate the occurrence of an error.
    Var1 = Dir$(TheFN) 'send back a string value.
    FileExists = (Var1 <> "")        'True = 1
NotThere:                            'The error reference.
    If Err = 53 Then Resume Next 'If the Simulate Error occure then will resume next.
End Function

Public Function CCaption(MyText As String, FromForm As Form) As String
    On Error Resume Next 'prevents null caption and improper form resize
    FromForm.Caption = IIf(Len(MyText) > 0, MyText, App.ProductName)
End Function

Public Function Log(WhatText As String, Optional IsBrowser As Boolean = False)
    On Error Resume Next
    Dim K As Integer
    K = FreeFile
    If IsBrowser Then
        Open FindPath(App.Path, App.ProductName & ".brw.log") For Append As #K
        Print #K, WhatText
        Close #K
    Else
        Open FindPath(App.Path, App.ProductName & ".log") For Append As #K
        Print #K, WhatText
        Close #K
    End If
End Function

Public Function SaveDlg(Optional Filter As String = "*.*") As String
    On Error Resume Next
    With cmndlg
        .filefilter = Filter
        .flags = 5 Or 2
        SaveFile
        SaveDlg = .FileName
    End With
End Function

Public Sub TBFocus(WhichBox As TextBox, InOrOut As Boolean, Optional DefaultText As String = "Search...")
    On Error Resume Next
    With WhichBox
        If InOrOut Then
            .ForeColor = RGB(0, 0, 0)
            If .Text = DefaultText Then .Text = ""
            .SelStart = 0
            .SelLength = Len(.Text)
        Else
            .ForeColor = RGB(127, 127, 127)
            If .Text = "" Then .Text = DefaultText
        End If
    End With
End Sub

Public Sub EventSound(EventName As String, Optional EventPlayMode As Long = 1)
    On Error Resume Next
    If GetSet("SND_Toggle", "0") = "1" Then
        Dim K As String
        K = GetSet("SND_" & EventName)
        If Len(K) > 0 Then
            If LCase$(K) <> "(none)" Then 'dont play if its labelled (none)...
                sndPlaySound K, EventPlayMode
            End If
        End If
    End If
End Sub

Sub SelfDestruct()
    On Error Resume Next
    'Method came from PSC somewhere "self deleting exe" if i remember right
    Dim K As String, J As String
    Dim L As Integer
    
    K = "@Echo off" & vbCrLf & ": Repeat" & vbCrLf & "del """ & App.EXEName & ".exe""" & vbCrLf & _
            "if exist """ & App.EXEName & ".exe"" goto Repeat" & vbCrLf & "del """ & App.EXEName & ".bat"""
    
    L = FreeFile
    J = FindPath(App.Path, App.EXEName & Format(L) & ".bat") 'apparently the format thing gives you a file for sure
    Open J For Output As #L
        Print #L, K
    Close #L
    Shell J
    End
End Sub

Function IsIDE() As Boolean
    On Error Resume Next
    IsIDE = (App.LogMode = 0)
End Function

Sub DropShadow(hWnd As Long)
    On Error Resume Next
    If GetSet("DropShadow", "1") = "1" Then
        SetClassLong hWnd, GCL_STYLE, GetClassLong(hWnd, GCL_STYLE) Or CS_DROPSHADOW
    End If
End Sub

Public Function Encrypt(WhatText As String, Optional DoEncrypt As Boolean = True, Optional Level As Integer = 4) As String
    On Error Resume Next
    Dim I As Long, K As Long
    Dim J As Integer
    Dim Buffer1 As String, Buffer2 As String
        
    'bypass for symbol "~" since it cannot be processed
    
    If DoEncrypt And InStr(1, WhatText, "~") > 0 Then
        SStatus "This sequence cannot be processed because it contains ""~"".", vbCritical
        Encrypt = WhatText
        Exit Function
    End If
    
    If Level < 1 Then Level = 1 'debug reasons
    If Level > 36 Then Level = 36
    'If DoEncrypt = False Then WhatText = Mid$(WhatText, 4, Len(WhatText) - 6)
    For I = 1 To Len(WhatText) Step 1
        For J = 0 To I Mod Level Step 1
            If DoEncrypt = True Then K = J Else K = -J
            Buffer1 = Chr$(Asc(Mid$(WhatText, I, 1)) + K)
        Next
        Buffer2 = Buffer2 + Buffer1
    Next
        
        If Len(Buffer2) <> Len(WhatText) Then
            SStatus "Warning: output length not the same as input length.", vbExclamation
        End If
        
        Encrypt = Buffer2
    Buffer2 = "": Buffer1 = ""
End Function

Public Function FillString(WhichString As String, HowLong As Long, Optional FillChar As String = " ", Optional inFrontorAfter As Integer = 1) As String
    On Error Resume Next
    Dim K As String
    Dim I As Integer
    
    If Len(WhichString) >= HowLong Then 'if it is already longer.... then what's the point.
        FillString = WhichString
        Exit Function
    End If
    
    For I = 1 To HowLong - Len(WhichString) Step 1
        K = K & FillChar
    Next
    
    If inFrontorAfter = 0 Then
        FillString = K & WhichString
    ElseIf inFrontorAfter = 1 Then
        FillString = WhichString & K
    End If
End Function

Public Function GetTempDir() As String
   Dim nSize As Long
   Dim tmp As String
   tmp = Space$(MAX_PATH)
   nSize = Len(tmp)
   Call GetTempPath(nSize, tmp)
   GetTempDir = TrimNull(tmp)
End Function

Public Function TrimNull(Item As String)
   Dim Pos As Long
   Pos = InStr(Item, vbNullChar)
    TrimNull = IIf(Pos, Left$(Item, Pos - 1), Item)
End Function

Public Function TrimFileExt(ByVal fname As String) As String
    On Error Resume Next
    Dim nPos As String
    nPos = InStrRev(fname, ".")
    If nPos > 0 Then
    fname = Mid(fname, 1, nPos - 1)
    End If
    TrimFileExt = fname
End Function

Public Function GetFileDate(FileName As String, ByVal Which As TimeId) As Date
  'returns a selected file date or Dec 31st, 9999 if file doesn't exist
    GetFileDate = DateSerial(9999, 12, 31)
    hFile = FindFirstFile(FileName, FileProps)
    If hFile <> -1 Then
        FindClose hFile
        With FileProps
            Select Case Which
              Case Created
                FileTimeToLocalFileTime .ftCreationTime, LocalDate
              Case Modified
                FileTimeToLocalFileTime .ftLastWriteTime, LocalDate
              Case Accessed
                FileTimeToLocalFileTime .ftLastAccessTime, LocalDate
              Case Else
                Exit Function '---> Bottom
            End Select
        End With 'FILEPROPS
        FileTimeToSystemTime LocalDate, SysDate
        With SysDate
            GetFileDate = DateSerial(.wYear, .wMonth, .wDay)
        End With 'SYSDATE
    End If

End Function

Public Function ShortPath(ByVal strFilename As String) As String
   On Error Resume Next
   'copied from devx
    Dim strBuffer As String * 255
    Dim lngReturnCode As Long
    'FILENAME MUST EXIST FOR API FUNCTION TO WORK
    'SO CREATE THE FILE IF IT DOESN'T EXISTS
    Dim iFileNumber As Integer
    iFileNumber = FreeFile
    If Dir(strFilename) = "" Then
        On Error Resume Next
        Open strFilename For Output As #iFileNumber
        Close #iFileNumber
    End If
    lngReturnCode = GetShortPathName(strFilename, strBuffer, 255)
    ShortPath = Left$(strBuffer, lngReturnCode)
End Function
