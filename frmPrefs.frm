VERSION 5.00
Begin VB.Form frmPrefs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Options"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   FillColor       =   &H80000016&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   4
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   18
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sync song name with messenger"
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   102
         ToolTipText     =   "Sync_PSM,0"
         Top             =   120
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel message when ProFile closes"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   54
         ToolTipText     =   "PSMCancelOnExit,1"
         Top             =   840
         Width           =   5055
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   7
         Left            =   6000
         TabIndex        =   23
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "PSMLoc,{app}\PSMChanger.exe"
         Top             =   1590
         Width           =   5655
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   8
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "PSM,%c %p - %n"
         Top             =   2520
         Width           =   5655
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Replace other YE company names with pow!!"
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "YE_Elim,1"
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%a - file full path"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   2400
         TabIndex        =   103
         Top             =   3840
         Width           =   1245
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%p - program name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   21
         Left            =   240
         TabIndex        =   60
         Top             =   3600
         Width           =   1440
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%c  - company name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   59
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%l  - length of the song"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   33
         Top             =   3360
         Width           =   1710
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%s  - the size of the file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   46
         Left            =   2400
         TabIndex        =   32
         Top             =   3360
         Width           =   1725
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%t  - the time the message is changed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   45
         Left            =   2400
         TabIndex        =   31
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%f  - name of the file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   44
         Left            =   240
         TabIndex        =   30
         Top             =   3120
         Width           =   1545
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "%n  - name of the song"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   38
         Left            =   2400
         TabIndex        =   29
         Top             =   3120
         Width           =   1710
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Symbols you can use:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   25
         Left            =   240
         TabIndex        =   28
         Top             =   2880
         Width           =   1560
      End
      Begin VB.Label lblURL 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "What's this?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   28
         Left            =   240
         TabIndex        =   27
         ToolTipText     =   "http://thinc.no-ip.info/projs/profile"
         Top             =   1920
         Width           =   870
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Your plugin:"
         Height          =   210
         Index           =   23
         Left            =   270
         TabIndex        =   25
         Top             =   1275
         Width           =   990
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Show your personal message like this:"
         Height          =   210
         Index           =   24
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton btnUnloadMe 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   5160
      Width           =   975
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4455
      Index           =   0
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   34
      Top             =   480
      Width           =   6495
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   2
         ItemData        =   "frmPrefs.frx":000C
         Left            =   0
         List            =   "frmPrefs.frx":0022
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   100
         ToolTipText     =   "RandomMusic,1"
         Top             =   1740
         Width           =   6495
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "Change >>"
         Height          =   375
         Index           =   5
         Left            =   5400
         TabIndex        =   97
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   5
         Left            =   0
         TabIndex        =   96
         ToolTipText     =   "SkinFile,"
         Top             =   360
         Width           =   5295
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Protect settings file"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   75
         ToolTipText     =   "Sandbox,0"
         Top             =   2520
         Width           =   6615
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Store settings locally (INI caching)"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   56
         ToolTipText     =   "Caching,1"
         Top             =   2880
         Width           =   6615
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   8
         Left            =   0
         Top             =   2280
         Width           =   6495
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "If I use random playback (Ctrl+M):"
         Height          =   210
         Index           =   12
         Left            =   0
         TabIndex        =   101
         Top             =   1440
         Width           =   4260
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   7
         Left            =   0
         Top             =   1320
         Width           =   6495
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Choose a skin."
         Height          =   210
         Index           =   8
         Left            =   0
         TabIndex        =   99
         Top             =   120
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSkin 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "N / A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   98
         Top             =   720
         Width           =   6450
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   6
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   81
      Top             =   480
      Width           =   6495
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   3
         ItemData        =   "frmPrefs.frx":00A6
         Left            =   1440
         List            =   "frmPrefs.frx":00B6
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   87
         ToolTipText     =   "OnlyOnePlayer,1"
         Top             =   600
         Width           =   5055
      End
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   4
         ItemData        =   "frmPrefs.frx":0116
         Left            =   1440
         List            =   "frmPrefs.frx":0123
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   86
         ToolTipText     =   "OnlyOneImgViewer,0"
         Top             =   1080
         Width           =   5055
      End
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   5
         ItemData        =   "frmPrefs.frx":0170
         Left            =   1440
         List            =   "frmPrefs.frx":017D
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   85
         ToolTipText     =   "OnlyOneTxtViewer,0"
         Top             =   1560
         Width           =   5055
      End
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   6
         ItemData        =   "frmPrefs.frx":01CA
         Left            =   1440
         List            =   "frmPrefs.frx":01D7
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   84
         ToolTipText     =   "OnlyOneBrowser,0"
         Top             =   2040
         Width           =   5055
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   6
         Left            =   6000
         TabIndex        =   83
         Top             =   4020
         Width           =   495
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   82
         ToolTipText     =   "WebEditor,notepad"
         Top             =   4050
         Width           =   5895
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Media Player:"
         Height          =   210
         Index           =   14
         Left            =   240
         TabIndex        =   94
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Image Viewer:"
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   93
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Text Viewer:"
         Height          =   210
         Index           =   50
         Left            =   240
         TabIndex        =   92
         Top             =   1620
         Width           =   1080
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Browser:"
         Height          =   210
         Index           =   51
         Left            =   600
         TabIndex        =   91
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Args: [exe] %f"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   18
         Left            =   5325
         TabIndex        =   90
         Top             =   3735
         Width           =   1095
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Web page editor:"
         Height          =   210
         Index           =   16
         Left            =   0
         TabIndex        =   89
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "You can open more than one file at a time. Change your settings here."
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   10
         Left            =   0
         TabIndex        =   88
         Top             =   120
         Width           =   5880
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   3
         Left            =   0
         Top             =   3600
         Width           =   6495
      End
   End
   Begin ProFile.Tab Tbx1 
      Height          =   4935
      Left            =   120
      TabIndex        =   74
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8705
      BackColor       =   -2147483633
      CloseButton     =   0   'False
      BlurForeColor   =   0
      ActiveForeColor =   0
      picture         =   "frmPrefs.frx":0226
      AllTabsForeColor=   -2147483630
      FontName        =   "Tahoma"
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   1
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   4
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open files after parsing"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   77
         ToolTipText     =   "OpenOnParse,1"
         Top             =   840
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show full paths in menus"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   76
         ToolTipText     =   "ShowFullPaths,"
         Top             =   1200
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Correct when ProFile flies out of screen"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   69
         ToolTipText     =   "MDIForm_AutoCenter,1"
         Top             =   480
         Width           =   4935
      End
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   1
         ItemData        =   "frmPrefs.frx":0242
         Left            =   2640
         List            =   "frmPrefs.frx":025B
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   65
         ToolTipText     =   "OpenOnStart,4"
         Top             =   3120
         Width           =   3855
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   0
         Left            =   2640
         TabIndex        =   58
         ToolTipText     =   "UpdaterURL,"
         Top             =   4080
         Width           =   3855
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check for updates on startup"
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   57
         ToolTipText     =   "UpdateOnStart,0"
         Top             =   3720
         Width           =   5055
      End
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   0
         ItemData        =   "frmPrefs.frx":02C8
         Left            =   2640
         List            =   "frmPrefs.frx":02D8
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   52
         ToolTipText     =   "Multiple_Instance,0"
         Top             =   2640
         Width           =   3855
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warn me when closing many windows"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   51
         ToolTipText     =   "MDIForm_MDIWarning,1"
         Top             =   120
         Width           =   5055
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   1
         Left            =   0
         Top             =   2400
         Width           =   6495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   0
         Left            =   0
         Top             =   3600
         Width           =   6495
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "From:"
         Height          =   210
         Index           =   1
         Left            =   2115
         TabIndex        =   80
         Top             =   4125
         Width           =   465
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Show on startup:"
         Height          =   210
         Index           =   7
         Left            =   1140
         TabIndex        =   66
         Top             =   3180
         Width           =   1440
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Multiple instances of program:"
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   53
         Top             =   2700
         Width           =   2445
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   7
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      Begin VB.CommandButton btnCachedINI 
         Caption         =   "Go to INI cache"
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   4080
         Width           =   1935
      End
      Begin VB.ListBox LstCredits 
         Height          =   2175
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrefs.frx":0313
         Left            =   2040
         List            =   "frmPrefs.frx":0341
         TabIndex        =   5
         Top             =   2280
         Width           =   4455
      End
      Begin VB.CommandButton btnShellINI 
         Caption         =   "&Go to INI"
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   2
         Left            =   0
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "List of beta testers: (thank you!)"
         Height          =   210
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
         Width           =   2715
      End
      Begin VB.Image imgLogo 
         Height          =   645
         Left            =   0
         Picture         =   "frmPrefs.frx":03FB
         Top             =   480
         Width           =   2190
      End
      Begin VB.Label lblProdVer 
         BackStyle       =   0  '³z©ú
         Caption         =   "Version "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label lblProdDes 
         BackStyle       =   0  '³z©ú
         Caption         =   "Description"
         Height          =   855
         Left            =   0
         TabIndex        =   3
         Top             =   1200
         Width           =   6855
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   5
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   19
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Play sounds when events take place"
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   95
         ToolTipText     =   "SND_Toggle,0"
         Top             =   120
         Width           =   4095
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   10
         Left            =   1560
         TabIndex        =   44
         ToolTipText     =   "SND_Start,(none)"
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   11
         Left            =   1560
         TabIndex        =   43
         ToolTipText     =   "SND_Close,(none)"
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   14
         Left            =   1560
         TabIndex        =   42
         ToolTipText     =   "SND_WinOpen,(none)"
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   15
         Left            =   1560
         TabIndex        =   41
         ToolTipText     =   "SND_WinClose,(none)"
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   16
         Left            =   1560
         TabIndex        =   40
         ToolTipText     =   "SND_MSGSkip,(none)"
         Top             =   2400
         Width           =   4335
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   6000
         TabIndex        =   39
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   6000
         TabIndex        =   38
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   4
         Left            =   6000
         TabIndex        =   37
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   5
         Left            =   6000
         TabIndex        =   36
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton btnSndPlaySound 
         Caption         =   "..."
         Height          =   315
         Index           =   6
         Left            =   6000
         TabIndex        =   35
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Program starts:"
         Height          =   210
         Index           =   27
         Left            =   225
         TabIndex        =   49
         Top             =   525
         UseMnemonic     =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Program closes:"
         Height          =   210
         Index           =   29
         Left            =   195
         TabIndex        =   48
         Top             =   1005
         UseMnemonic     =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Window opens:"
         Height          =   210
         Index           =   32
         Left            =   180
         TabIndex        =   47
         Top             =   1485
         UseMnemonic     =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Window closes:"
         Height          =   210
         Index           =   33
         Left            =   195
         TabIndex        =   46
         Top             =   1965
         UseMnemonic     =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Decision made:"
         Height          =   210
         Index           =   34
         Left            =   240
         TabIndex        =   45
         Top             =   2445
         UseMnemonic     =   0   'False
         Width           =   1230
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   3
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   9
      Top             =   480
      Width           =   6495
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show the search bar"
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   26
         ToolTipText     =   "SearchBar,1"
         Top             =   120
         Width           =   5055
      End
      Begin VB.ListBox lstSearchURL 
         Height          =   1740
         ItemData        =   "frmPrefs.frx":0C2B
         Left            =   1920
         List            =   "frmPrefs.frx":0C4A
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ListBox lstSearchName 
         Height          =   2340
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrefs.frx":0E6A
         Left            =   240
         List            =   "frmPrefs.frx":0E89
         TabIndex        =   16
         Top             =   720
         Width           =   6255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Search_Provider_Name,Google"
         Top             =   3360
         Width           =   6255
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Search_Provider_URL,http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
         Top             =   4080
         Width           =   6255
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Select a search engine you would like to use."
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   4380
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Search Provider Name:"
         Height          =   210
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   4380
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Search Provider Search String:"
         Height          =   210
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   4380
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   4455
      Index           =   2
      Left            =   240
      ScaleHeight     =   4455
      ScaleWidth      =   6495
      TabIndex        =   17
      Top             =   480
      Width           =   6495
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   7
         ItemData        =   "frmPrefs.frx":0F03
         Left            =   2880
         List            =   "frmPrefs.frx":0F13
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   78
         ToolTipText     =   "URL_AutoCorrect,1"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Open new windows in new tabs"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   72
         ToolTipText     =   "NewBrowserInTab,1"
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close button on every tab (restart to see effect)"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   71
         ToolTipText     =   "TabCloseButton,1"
         Top             =   3480
         Width           =   5055
      End
      Begin VB.ComboBox cboOpt 
         Height          =   330
         Index           =   8
         ItemData        =   "frmPrefs.frx":0F5F
         Left            =   2880
         List            =   "frmPrefs.frx":0F6F
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   70
         ToolTipText     =   "Browser_Init,2"
         Top             =   2520
         Width           =   3615
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Keep web browsing history"
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   68
         ToolTipText     =   "BRW_Log,1"
         Top             =   120
         Width           =   5055
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show HTML tags when I mouse my mouse over them"
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   67
         ToolTipText     =   "ShowTags,0"
         Top             =   480
         Width           =   5055
      End
      Begin VB.CommandButton btnFolderBrowse 
         Caption         =   "Browse..."
         Height          =   375
         Index           =   19
         Left            =   5400
         TabIndex        =   63
         Top             =   1245
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   19
         Left            =   2880
         TabIndex        =   62
         ToolTipText     =   "FAV_Bookmarks,"
         Top             =   1275
         Width           =   2415
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Use favorites toolbar"
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   50
         ToolTipText     =   "BRW_AutoFavsBarSwitch,1"
         Top             =   960
         Width           =   5055
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   6
         Left            =   0
         Top             =   1800
         Width           =   6495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   5
         Left            =   0
         Top             =   840
         Width           =   6495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000016&
         FillColor       =   &H80000016&
         Height          =   15
         Index           =   4
         Left            =   0
         Top             =   3000
         Width           =   6855
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Wong URL e.g. ""ww.google.cm"":"
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   79
         Top             =   2100
         Width           =   2745
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "When browser opens:"
         Height          =   210
         Index           =   15
         Left            =   960
         TabIndex        =   73
         Top             =   2580
         Width           =   1830
      End
      Begin VB.Label lblNotification 
         Alignment       =   1  '¾a¥k¹ï»ô
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Favorites folder:"
         Height          =   210
         Index           =   20
         Left            =   1515
         TabIndex        =   64
         Top             =   1320
         Width           =   1305
      End
   End
   Begin VB.Label lblNotification 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Preferences cannot be sandboxed"
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
      Index           =   22
      Left            =   360
      TabIndex        =   61
      Top             =   5280
      Width           =   2475
   End
   Begin VB.Image imgPadlock 
      Height          =   270
      Left            =   120
      Picture         =   "frmPrefs.frx":0FBA
      ToolTipText     =   "Not locked by sandbox mode"
      Top             =   5205
      Width           =   195
   End
   Begin VB.Image IMGbkg 
      Height          =   5655
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
   Begin VB.Menu titSounds 
      Caption         =   "Sounds"
      Visible         =   0   'False
      Begin VB.Menu titSoundsPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu titS14 
         Caption         =   "-"
      End
      Begin VB.Menu titSoundsBrowseFile 
         Caption         =   "Browse for a file"
      End
      Begin VB.Menu titSoundsRemoveFile 
         Caption         =   "Do not play this file"
      End
   End
   Begin VB.Menu titFiles 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu titFilesBrowse 
         Caption         =   "Browse..."
      End
      Begin VB.Menu titFilesGoto 
         Caption         =   "Goto path"
      End
      Begin VB.Menu titFilesClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu titPages 
      Caption         =   "Tab1Popup"
      Visible         =   0   'False
      Begin VB.Menu titPR 
         Caption         =   "Filler"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu titPR 
         Caption         =   "Other general prefs"
         Index           =   1
      End
      Begin VB.Menu titPR 
         Caption         =   "Appearance"
         Index           =   2
      End
      Begin VB.Menu titPR 
         Caption         =   "Search options"
         Index           =   3
      End
      Begin VB.Menu titPR 
         Caption         =   "Messenger options"
         Index           =   4
      End
      Begin VB.Menu titPR 
         Caption         =   "Sound prefs"
         Index           =   5
      End
      Begin VB.Menu titPR 
         Caption         =   "Favorites options"
         Index           =   6
      End
      Begin VB.Menu titPR 
         Caption         =   "Tab options"
         Index           =   7
      End
      Begin VB.Menu titPR 
         Caption         =   "Media options"
         Index           =   8
      End
      Begin VB.Menu titPR 
         Caption         =   "File Open options"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InaSub As Boolean
Dim TheFilter As String
Dim InTheTab As Integer

Private Sub btnBrowse_Click(Index As Integer)
    On Error Resume Next

    titFiles.Tag = Index
    
    If Index = 5 Then
        TheFilter = "Configuration file (*.ini)|*.ini"
        PopupMenu titFiles, , picTabSwitch(InTheTab).Left + btnBrowse(Index).Left, picTabSwitch(5).Top + btnBrowse(Index).Top + btnBrowse(Index).Height, titFilesBrowse
    Else
        TheFilter = "All files (*.*)|*.*"
        PopupMenu titFiles, , picTabSwitch(InTheTab).Left + btnBrowse(Index).Left, picTabSwitch(5).Top + btnBrowse(Index).Top + btnBrowse(Index).Height, titFilesBrowse
    End If
    
End Sub

Private Sub btnCachedINI_Click()
    On Error Resume Next
    Shell "explorer " & GetTempDir, vbNormalFocus
End Sub

Private Sub btnFolderBrowse_Click(Index As Integer)
    On Error Resume Next
    K = BrowseForFolder(Me.hWnd)
    If Len(K) = 0 Then Exit Sub
    
    txtData(Index).Text = K
End Sub

Private Sub btnGoTo_Click(Index As Integer)
    On Error Resume Next
    Shell "explorer " & PathOnly(txtData(Index).Text), vbNormalFocus
End Sub

Private Sub btnGoTab_Click(Index As Integer)
    On Error Resume Next
    GoToTab Index
End Sub

Private Sub btnOK_Click()
    On Error Resume Next
        
    Dim I As Integer
    For I = 0 To chkOpt.UBound Step 1 'Save Settings
        If Len(chkOpt(I).Tag) > 0 Then
            WriteINI UserName, GetString(chkOpt(I).ToolTipText), Str$(chkOpt(I).Value), SettingsFile, True
            'WriteINI UserName, GetString(chkOpt(i).ToolTipText), Str$(chkOpt(i).Value), FindPath(GetTempDir, TempINI)
        End If
    Next
    For I = 0 To txtData.UBound Step 1 'Save Settings
        If Len(txtData(I).Tag) > 0 Then
            WriteINI UserName, GetString(txtData(I).ToolTipText), txtData(I).Text, SettingsFile, True
            'WriteINI UserName, GetString(txtData(i).ToolTipText), txtData(i).Text, FindPath(GetTempDir, TempINI)
        End If
    Next
    For I = 0 To cboOpt.UBound Step 1 'Save Settings
        If Len(cboOpt(I).Tag) > 0 Then
            WriteINI UserName, GetString(cboOpt(I).ToolTipText), cboOpt(I).ListIndex - 1, SettingsFile, True
            'WriteINI UserName, GetString(cboOpt(i).ToolTipText), cboOpt(i).ListIndex - 1, FindPath(GetTempDir, TempINI)
        End If
    Next
    If Len(txtData(24).Tag) > 0 Or Len(txtData(31).Tag) > 0 Then 'if display is changed
        SkinForm frmMain
        SkinFormEx frmMain 'for the heck of it
    End If
    
    Unload Me
End Sub


'Private Sub btnSelectFont_Click()
'    If Len(txtData(31).Text) > 0 Then SelectFont.mFontSize = Val(txtData(31).Text)
'    SelectFont.mFontName = txtData(24).Text
'    ShowFont
'    If Len(SelectFont.mFontName) > 0 Then txtData(24).Text = SelectFont.mFontName
'    If Len(SelectFont.mFontSize) > 0 Then txtData(31).Text = SelectFont.mFontSize
'End Sub

Private Sub btnShellINI_Click()
    On Error Resume Next
    Shell "explorer " & PathOnly(SettingsFile(True)), vbNormalFocus
End Sub

Private Sub btnSndPlaySound_Click(Index As Integer)
    On Error Resume Next
    titFiles.Tag = Index + 10
    PopupMenu titSounds, , picTabSwitch(5).Left + btnSndPlaySound(Index).Left, picTabSwitch(5).Top + btnSndPlaySound(Index).Top + btnSndPlaySound(Index).Height, titSoundsPlay
End Sub

Private Sub btnSplash_Click()
'    On Error Resume Next
'    With frmSplash
'        .ForceClick = True
'        .T1.Enabled = False
'        .Show 1
'    End With
End Sub

'Private Sub btnTab_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            GoToTab 0
'        Case 1
'            PopupMenu titPages, , btnTab(Index).Left, btnTab(Index).Height
'        Case 2
'            GoToTab 10
'    End Select
'End Sub

Private Sub btnUnloadMe_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub cboOpt_Change(Index As Integer)
    On Error Resume Next 'for real-time stuff, no undo
    If InaSub Then Exit Sub
    cboOpt(Index).Tag = "EDITED"
End Sub

Private Sub cboOpt_Click(Index As Integer)
    cboOpt_Change Index 'stub
End Sub

Private Sub chkOpt_Click(Index As Integer)
    On Error Resume Next 'for real-time stuff, no undo
    If InaSub Then Exit Sub
    With chkOpt(Index)
        Select Case Index
            Case 3
                If .Value = 1 Then DSA 9
        End Select
    End With
    chkOpt(Index).Tag = "EDITED"
End Sub

Private Sub Form_Activate()
    MakeTransparent frmMain.hWnd, 80
    InitCommonControls
    If CheckPW = False Then Unload Me 'protection
'    F1.FadeIn
End Sub

'Private Sub Form_Deactivate()
'    F1.FadeOut
'End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    
    Dim K As String
'    F1.PrepareFade
    InaSub = True
    
    Tbx1.AddTabs "Popular", "General", "Browser", "Search", "Messenger", "Sound", "Actions", "About"
    
    lblProdVer.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblProdDes.Caption = App.ProductName & " " & MyVer & ", some rights reserved by Thinc." & vbCrLf & _
    "Made by Brian Lai" & vbCrLf & SoftwareHomePage
        
    SkinForm Me
    SkinFormEx Me

    SProgress 33
    SStatus "Loading checkboxes", vbExclamation
    For I = 0 To chkOpt.UBound Step 1 'Load Settings
        K = GetString(chkOpt(I).ToolTipText, 1)
        K = ReplaceDynamicPaths(K)
        chkOpt(I).Value = GetSet(GetString(chkOpt(I).ToolTipText), K, , , True)
        DoEvents
    Next
    SProgress 66
    
    
    SStatus "Loading combo boxes", vbExclamation
    For I = 0 To cboOpt.UBound Step 1 'Load Settings
        K = GetString(cboOpt(I).ToolTipText, 1)
        K = ReplaceDynamicPaths(K)
        Dim L As String
        L = GetSet(GetString(cboOpt(I).ToolTipText), K, , , True)
        If Len(L) > 0 Then
            cboOpt(I).ListIndex = Val(L) + 1 '+1 is to make up for "nothing"=0
        Else
            cboOpt(I).ListIndex = 0
        End If
        DoEvents
    Next
    SProgress 100
    
    SStatus "Loading textboxes", vbExclamation
    For I = 0 To txtData.UBound Step 1 'Load Settings
        K = GetString(txtData(I).ToolTipText, 1)
        K = ReplaceDynamicPaths(K)
        txtData(I).Text = GetSet(GetString(txtData(I).ToolTipText), K, , , True)
        DoEvents
    Next
    SProgress 0
    
    InaSub = False
    
    lblSkin(0).Caption = "Skin Info: " & SkinInfo("Info")
    EventSound "WinOpen"
    SStatus
    
End Sub

Public Function SkinInfo(KeyName As String) As String
    On Error Resume Next
    SkinInfo = ReadINI("Skin", KeyName, GetSet("SkinFile"))
    If Len(SkinInfo) = 0 Then SkinInfo = "n/a"
End Function

'Private Sub Form_Resize()
'    On Error Resume Next
'    'This is a fix for a random WinPos bug
'    Me.Height = btnUnloadMe.Top + btnUnloadMe.Height + 120 + (Me.Height - Me.ScaleHeight)
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    EventSound "WinClose"
    MakeOpaque frmMain.hWnd
End Sub

Private Sub lblURL_Click(Index As Integer)
    On Error Resume Next
    Shell "explorer " & lblURL(Index).ToolTipText, vbNormalFocus
End Sub

Private Sub lstSearchName_Click()
    On Error Resume Next
    txtData(2).Text = lstSearchName.List(lstSearchName.ListIndex)
    txtData(3).Text = lstSearchURL.List(lstSearchName.ListIndex)
End Sub

Private Sub LstTab_Click()
    On Error Resume Next
    'picTabSwitch(LstTab.ListIndex).ZOrder 0
    GoToTab LstTab.ListIndex
End Sub

Public Function GoToTab(Index As Integer)
    'On Error Resume Next
    Dim I As Integer
    InTheTab = Index
    Tbx1.ActiveTab = Index + 1
    For I = 0 To picTabSwitch.UBound
        picTabSwitch(I).Visible = False 'for the sake of making things accessible
    Next
    picTabSwitch(Index).Visible = True
    picTabSwitch(Index).ZOrder 0
End Function

Private Sub Tbx1_TabClick(tIndex As Integer)
End Sub

Private Sub Tbx1_Click(tIndex As Integer)
    'On Error Resume Next
    GoToTab tIndex - 1
End Sub

Private Sub titFilesBrowse_Click()
    On Error Resume Next
    With cmndlg
        .filefilter = TheFilter 'load the one
        If Len(.filefilter) = 0 Then .filefilter = "any file (*.*)|*.*" 'if there isnt one
        OpenFile
        If Len(.FileName) = 0 Then Exit Sub
        txtData(Val(titFiles.Tag)).Text = .FileName
    End With
End Sub

Private Sub titFilesClear_Click()
    txtData(Val(titFiles.Tag)).Text = ""
End Sub

Private Sub titFilesGoto_Click()
    Shell "explorer " & PathOnly(txtData(Val(titFiles.Tag)).Text), vbNormalFocus
End Sub

Private Sub titPR_Click(Index As Integer)
    GoToTab Index
End Sub

Private Sub titSoundsBrowseFile_Click()
    TheFilter = "wave files (*.wav)|*.wav"
    titFilesBrowse_Click
End Sub

Private Sub titSoundsPlay_Click()
    On Error Resume Next
    sndPlaySound txtData(Val(titFiles.Tag)).Text, 1
End Sub

Private Sub titSoundsRemoveFile_Click()
    On Error Resume Next
    txtData(Val(titFiles.Tag)).Text = "(None)"
End Sub

Private Sub txtData_Change(Index As Integer)
    If InaSub Then Exit Sub
    If Index <= 18 And Index >= 10 Then 'this is for the sake of having no music file
        If txtData(Index).Text = "" Then txtData(Index).Text = "(None)"
    End If
    txtData(Index).Tag = "EDITED"
End Sub
