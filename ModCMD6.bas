Attribute VB_Name = "ModCMD6"
Option Explicit

Function CMD6(Cmd As String, Optional DebugError As Boolean = False)
    On Error Resume Next
    Dim Params As String
    Dim D As New frmIMG
    Dim S As New frmBRW
    Dim I As Long 'NOTE I as a LONG here
    
    Params = Mid$(Cmd, InStr(1, Cmd, " ")) 'the rest of the command
    
    Select Case LCase$(GetString(Cmd, 0, " "))
        Case "."
            S.LoadFile "http://www.slashdot.org"
        Case "about"
            frmPrefs.GoToTab 7
            frmPrefs.Show 1
        Case "end"
            End
        Case "forum"
            D.LoadFile DownloadFile("http://www.kgv.net/blai/Images/Fi0" & FillString(GetString(Cmd, 1, " "), 2, "0", 0) & ".jpg")
        Case "gotopath"
            frmMain.GoToPath Params
        Case "iside"
            MsgBox IsIDE
        Case "kill"
            Select Case LCase$(GetString(Cmd, 1, " "))
                Case "me"
                    If MsgBox("Are you sure you want to remove this program from the hard drive?", vbYesNo + vbExclamation) = vbYes Then
                        SelfDestruct
                    End If
                Case "comp"
                    SHShutDownDialog 0
            End Select
        Case "minigame"
            Shell FindPath(App.Path, "TEExt.exe tetramg" & GetString(Cmd, 1, " ")), vbNormalFocus
        Case "mood"
            frmMood.Show
        Case "psm"
            Shell GetSet("PSMLoc", "{app}\PSMChanger.exe", , True) & Params
        Case "porn" 'EMERSON LI
            Select Case Val(GetString(Cmd, 1, " "))
                Case 1
                    D.LoadFile DownloadFile("http://www.kgv.net/eli/nice.jpg")
                Case 2
                    D.LoadFile DownloadFile("http://z.about.com/d/healing/1/0/u/L/affirm_random.jpg")
                Case 3
                    D.LoadFile DownloadFile("http://www.kgv.net/blai/Images/lee_hyori_wallpaper_22.jpg")
                Case 4
                    D.LoadFile DownloadFile("http://facultystaff.vwc.edu/~mhall/advertisements/underwear/images/woman%20looking%20at%20camera%20with%20hand%20between%20legs--unknown%20women's%2001.jpg")
            End Select
        Case "shell" 'just do it
            Shell Trim$(Params), vbNormalFocus
        Case "sign"
            D.LoadFile DownloadFile("http://www.kgv.net/blai/Images/S(" & Format(GetString(Cmd, 1, " ")) & ").jpg")
'        Case "splash"
'            frmSplash.Show 1
        Case "status"
            SStatus Params
        Case "trans", "transparent", "t"
            If Val(GetString(Cmd, 1, " ")) <> 10 Then
                MakeTransparent frmMain.hWnd, Val(GetString(Cmd, 1, " ")) * 10 '25
            Else
                MakeOpaque frmMain.hWnd
            End If
        Case "tray"
            NoSysIcon False
        Case Else
            DoEvents 'no actions defined for ProFile
    End Select
End Function
