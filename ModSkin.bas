Attribute VB_Name = "ModSkin"
Option Explicit
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
'No need for other INI readers

Function SkinForm(WhichForm As Form, Optional FromINIFile As String)
    On Error Resume Next
    Dim MyKeys As String * 50000 'This number is the limit of the length of the read string
    Dim EachElement As Variant, EachKey As Variant
    Dim CtlName As String, CtlProp As String, CtlPropVal As String
    Dim ItemIdx As Integer
    Dim CTx ' As Control
        
    If Len(FromINIFile) = 0 Then FromINIFile = GetSet("SkinFile", DefaultSkinFile) 'FindPath(App.Path, "skin.ini"))
    If Dir(FromINIFile) = "" Then Exit Function 'If there's no such file, who cares?
    
    'SkinForm uses GetPrivateProfileSection
    GetPrivateProfileSection WhichForm.Name, MyKeys, 60000, FromINIFile
    EachKey = Split(MyKeys, Chr(0))
    For Each EachElement In EachKey
    
        If LenB(EachElement) = 0 Then Exit For 'not bothered

        'EachElement is in Label1 BackColor=255 form right now, so split with GetPrivString
        CtlName = GetPrivString(GetPrivString(EachElement, 0, "="), 0, " ") 'This gets the control name
        ItemIdx = Val(Mid$(CtlName, InStr(1, CtlName, "(") + 1, (InStrRev(CtlName, ")") - (InStr(1, CtlName, "(") + 1))))
        'ItemIdx is calculated by a crazy length of code
        CtlProp = GetPrivString(GetPrivString(EachElement, 0, "="), 1, " ") 'This gets the property, e.g. BackColor
        CtlPropVal = GetPrivString(EachElement, 1, "=") 'This gets the value of that property
        CtlPropVal = ReplaceDynamicPaths(CtlPropVal)
        
        If Len(CtlProp) = 0 Then Exit For
        If Len(CtlName) = 0 Then Exit For
        
        If CtlName = "form" Then
            Set CTx = WhichForm
        Else
            If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                Set CTx = WhichForm.Controls(CtlName).Item(ItemIdx) 'set item as something in an array
            Else
                Set CTx = WhichForm.Controls(CtlName) 'set item
            End If
        End If
                
        Select Case Trim$(UCase$(CtlProp))
            Case "BACKCOLOR", "BC"
                CTx.BackColor = Val(CtlPropVal)
            Case "BACKOVER", "BO"
                CTx.BackOver = Val(CtlPropVal)
            Case "BORDERSTYLE"
                CTx.BorderStyle = Val(CtlPropVal)
            Case "BUTTONTYPE"
                CTx.ButtonType = Val(CtlPropVal)
            Case "CAPTION", "CPN"
                CTx.Caption = CtlPropVal
            Case "COLORSCHEME"
                CTx.ColorScheme = Val(CtlPropVal)
            Case "FONT"
                CTx.FontName = GetString(CtlPropVal, 0)
                CTx.FontSize = GetString(CtlPropVal, 1)
                CTx.FontBold = (Val(GetString(CtlPropVal, 2)) = 1)
                CTx.FontItalic = (Val(GetString(CtlPropVal, 3)) = 1)
            Case "FORECOLOR", "FC"
                    CTx.ForeColor = Val(CtlPropVal)
            Case "HEIGHT", "H"
                CTx.Height = Val(CtlPropVal)
            Case "LEFT", "L"
                CTx.Left = Val(CtlPropVal)
            Case "PICTURE", "PIC"
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                CTx.Picture = LoadPicture(CtlPropVal)
            Case "PICTURENORMAL", "PN"
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                CTx.PictureNormal = LoadPicture(CtlPropVal)
            Case "PICTUREOVER", "PO"
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                CTx.PictureOver = LoadPicture(CtlPropVal)
            Case "STRETCH"
                CTx.Stretch = (Val(CtlPropVal) = 1)
            Case "TEXT", "TXT"
                CTx.Text = CtlPropVal
            Case "TOP", "T"
                CTx.Top = Val(CtlPropVal)
            Case "VISIBLE", "VS"
                CTx.Visible = (Val(CtlPropVal) = 1)
            Case "WIDTH", "W"
                CTx.Width = Val(CtlPropVal)
        End Select
        
        Set CTx = Nothing 'theres a bug here which I don't know why, but it should never hurt to unload memory
        DoEvents
    Next
End Function

Private Function GetPrivString(Which As Variant, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    'This GetPrivString is only for use in SkinForm because the source is a Variant there...
    'The general function is GetString.
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetPrivString = Arr(SectionNo)
End Function

Public Function ReplaceDynamicPaths(FromWhat As String) As String
    On Error Resume Next
    'Converts to local paths
    FromWhat = Replace(FromWhat, "{a}", App.Path)
    FromWhat = Replace(FromWhat, "{app}", App.Path)
    FromWhat = Replace(FromWhat, "{s}", SkinPath)
    FromWhat = Replace(FromWhat, "{skin}", SkinPath)
    ReplaceDynamicPaths = FromWhat
End Function

Public Function SkinPath() As String
    On Error Resume Next
    Dim K As String
    K = ReadINI(UserName, "SkinFile", SettingsFile)
    If Len(K) = 0 Then K = FindPath(App.Path, "skin.ini") 'low level programming - because too many things depend on this
    SkinPath = Left$(K, InStrRev(K, "\") - 1)
End Function
