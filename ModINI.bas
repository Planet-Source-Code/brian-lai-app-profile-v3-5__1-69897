Attribute VB_Name = "ModINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Const PublicUserName = "Public"
Public Const TempINI = "pf-cache.ini"
Public Const AutoFillINI = "AutoFill.ini"

Public CacheINI As Boolean

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
On Error Resume Next
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName, Optional ToOriginalToo As Boolean) As Integer
On Error Resume Next
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
    If ToOriginalToo Then
        writeprivateprofilestring sSection, sKeyName, sNewString, SettingsFile(True)
    End If
End Function

Function GetSet(Key As String, Optional Default As String, Optional ForUser As String, Optional ReplaceAppWithPath As Boolean = True, Optional OnlySettingsFromSetFile As Boolean) As String
    On Error Resume Next
    Dim Buffer As String, Buffer2 As String
    Dim ReadFrom As String
    Dim G As Integer
    
    If Len(ForUser) = 0 Then ForUser = UserName 'if there is no username then fill in user name for me
    
    ReadFrom = SettingsFile(Not CacheINI)
    
    If OnlySettingsFromSetFile = False Then
        Buffer2 = ReadINI(ForUser, "SkinSet", ReadFrom) 'get if user wants interaction
        If Buffer2 = "1" Then 'if the user didn't disable this
            Buffer2 = ""
            Buffer2 = ReadINI(ForUser, "SkinFile", ReadFrom) 'get where the skin file is
            If Len(Buffer2) > 0 Then 'if theres a skin
                Buffer = ReadINI("Settings", Key, Buffer2) 'try again with the public section
                If Len(Buffer) > 0 Then
                    G = 1
                    GetSet = Buffer 'if theres something then so be it
                    GoTo ExitFunction
                End If
            End If
        End If
    End If
    
    Buffer = ReadINI(ForUser, Key, ReadFrom) 'read the entry from my username section
    If Len(Buffer) > 0 Then 'if there's an entry in my user name
        G = 2
        GetSet = Buffer 'then so be it
        GoTo ExitFunction
    End If
    
    Buffer = ReadINI(PublicUserName, Key, ReadFrom) 'try again with the public section
    If Len(Buffer) > 0 Then 'if there's an entry in the public user name
        G = 3
        GetSet = Buffer 'then so be it
        GoTo ExitFunction
    End If
    
    If Len(Default) > 0 Then 'if nothing else is present but I have a default...
        G = 4
        GetSet = Default 'then so be it
        GoTo ExitFunction
    End If

Exit Function

ExitFunction:
    If ReplaceAppWithPath Then GetSet = ReplaceDynamicPaths(GetSet)
End Function

Function SaveSet(Key As String, Value As String, Optional ForUser As String) As String
On Error Resume Next
    If Len(ForUser) = 0 Then ForUser = UserName
    If ReadINI(UserName, "Sandbox", SettingsFile) <> "1" Then 'this stops all ini writes via SaveSet if sandbox is on
        WriteINI ForUser, Key, Value, SettingsFile, True
    End If
SaveSet = Key
End Function

Public Function SettingsFile(Optional RealFile As Boolean) As String 'REALFILE is optional false by default because, well, Caching is on by default
    On Error Resume Next
    If RealFile Then
         'you asked for real
         SettingsFile = FindPath(App.Path, App.ProductName & ".ini")
    Else
        SettingsFile = FindPath(GetTempDir, TempINI)
    End If
End Function
