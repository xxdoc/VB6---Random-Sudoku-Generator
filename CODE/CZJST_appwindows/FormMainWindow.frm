VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Sudoku Generator　v0.11　by Sam Toki"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   13455
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7260
   ScaleWidth      =   13455
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TimerSudokuGenerator 
      Interval        =   20
      Left            =   315
      Top             =   420
   End
   Begin VB.Frame FrameControls 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   7035
      TabIndex        =   103
      Top             =   945
      Width           =   6210
      Begin VB.TextBox TextboxInput 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5355
         MaxLength       =   1
         MouseIcon       =   "FormMainWindow.frx":2524
         MousePointer    =   99  'Custom
         TabIndex        =   124
         Top             =   1995
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "C"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   4200
         MouseIcon       =   "FormMainWindow.frx":2676
         MousePointer    =   99  'Custom
         TabIndex        =   123
         Top             =   1995
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "9"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   9
         Left            =   3255
         MouseIcon       =   "FormMainWindow.frx":27C8
         MousePointer    =   99  'Custom
         TabIndex        =   121
         Top             =   1995
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "8"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   8
         Left            =   2310
         MouseIcon       =   "FormMainWindow.frx":291A
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   1995
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "7"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   7
         Left            =   1365
         MouseIcon       =   "FormMainWindow.frx":2A6C
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Top             =   1995
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   6
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":2BBE
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   1995
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   4200
         MouseIcon       =   "FormMainWindow.frx":2D10
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   1260
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   3255
         MouseIcon       =   "FormMainWindow.frx":2E62
         MousePointer    =   99  'Custom
         TabIndex        =   111
         Top             =   1260
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   2310
         MouseIcon       =   "FormMainWindow.frx":2FB4
         MousePointer    =   99  'Custom
         TabIndex        =   109
         Top             =   1260
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   1365
         MouseIcon       =   "FormMainWindow.frx":3106
         MousePointer    =   99  'Custom
         TabIndex        =   107
         Top             =   1260
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":3258
         MousePointer    =   99  'Custom
         TabIndex        =   105
         Top             =   1260
         Width           =   540
      End
      Begin VB.CommandButton CmdStartReset 
         Caption         =   "START / RESET"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":33AA
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   525
         Width           =   4320
      End
      Begin VB.Shape ShapeLightGameOngoingIndicator 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000C000&
         Height          =   330
         Left            =   5565
         Shape           =   3  'Circle
         Top             =   525
         Width           =   330
      End
      Begin VB.Shape ShapeLightGeneratingIndicator 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   330
         Left            =   5145
         Shape           =   3  'Circle
         Top             =   525
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   9
         Left            =   3780
         TabIndex        =   122
         Top             =   2205
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   8
         Left            =   2835
         TabIndex        =   120
         Top             =   2205
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   7
         Left            =   1890
         TabIndex        =   118
         Top             =   2205
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   6
         Left            =   945
         TabIndex        =   116
         Top             =   2205
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   4725
         TabIndex        =   114
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   3780
         TabIndex        =   112
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   2835
         TabIndex        =   110
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   1890
         TabIndex        =   108
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   945
         TabIndex        =   106
         Top             =   1470
         Width           =   330
      End
   End
   Begin VB.Frame FrameSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   7035
      TabIndex        =   125
      Top             =   3990
      Width           =   6210
      Begin VB.HScrollBar HScrollSettingsLargeBlockMaximumFixedAmount 
         Height          =   330
         Left            =   3675
         Max             =   8
         Min             =   1
         MouseIcon       =   "FormMainWindow.frx":34FC
         MousePointer    =   99  'Custom
         TabIndex        =   131
         Top             =   1365
         Value           =   6
         Width           =   2220
      End
      Begin VB.HScrollBar HScrollSettingsTotalFixedAmount 
         Height          =   330
         LargeChange     =   9
         Left            =   3675
         Max             =   54
         Min             =   9
         MouseIcon       =   "FormMainWindow.frx":364E
         MousePointer    =   99  'Custom
         TabIndex        =   128
         Top             =   525
         Value           =   36
         Width           =   2220
      End
      Begin VB.Label LabelSettingsLargeBlockMaximumFixedAmountIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3045
         TabIndex        =   130
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label LabelSettingsTotalFixedAmountIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3045
         TabIndex        =   127
         Top             =   525
         Width           =   540
      End
      Begin VB.Label LabelSettingsLargeBlockMaximumFixedAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum amount of fixed blocks in a large block:"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   315
         TabIndex        =   129
         Top             =   1050
         Width           =   5055
      End
      Begin VB.Label LabelSettingsTotalFixedAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount of all fixed blocks:"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   315
         TabIndex        =   126
         Top             =   525
         Width           =   2535
      End
   End
   Begin VB.Timer TimerProgressbarAnimation 
      Interval        =   1
      Left            =   7035
      Top             =   6720
   End
   Begin VB.CommandButton CmdEXIT 
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12180
      MouseIcon       =   "FormMainWindow.frx":37A0
      MousePointer    =   99  'Custom
      TabIndex        =   134
      Top             =   210
      Width           =   1065
   End
   Begin VB.Timer TimerSudokuBlockAnimation 
      Interval        =   20
      Left            =   6300
      Top             =   6405
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   12915
      Top             =   7035
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   435
      Left            =   315
      TabIndex        =   135
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
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
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Shape ShapeProgressbar 
      BackColor       =   &H00FF8800&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   7035
      Top             =   6615
      Width           =   6000
   End
   Begin VB.Label LabelRowIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   6270
      Width           =   330
   End
   Begin VB.Label LabelColumnIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6195
      TabIndex        =   11
      Top             =   105
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   6195
      TabIndex        =   20
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   8
      Left            =   5565
      TabIndex        =   19
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   4935
      TabIndex        =   18
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   4200
      TabIndex        =   17
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   3570
      TabIndex        =   16
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   2940
      TabIndex        =   15
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   2205
      TabIndex        =   14
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   1575
      TabIndex        =   13
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   945
      TabIndex        =   12
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   6720
      TabIndex        =   21
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   420
      TabIndex        =   9
      Top             =   6300
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   8
      Left            =   420
      TabIndex        =   8
      Top             =   5670
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   420
      TabIndex        =   7
      Top             =   5040
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   420
      TabIndex        =   6
      Top             =   4305
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   420
      TabIndex        =   5
      Top             =   3675
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   3045
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   420
      TabIndex        =   3
      Top             =   2310
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   6720
      Width           =   645
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1050
      Width           =   330
   End
   Begin VB.Label LabelClock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   11865
      TabIndex        =   133
      Top             =   6825
      Width           =   1380
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   4
      X1              =   4725
      X2              =   4725
      Y1              =   840
      Y2              =   6825
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   2730
      X2              =   2730
      Y1              =   840
      Y2              =   6825
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   2
      X1              =   735
      X2              =   6720
      Y1              =   4830
      Y2              =   4830
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      X1              =   735
      X2              =   6720
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   80
      Left            =   5460
      TabIndex        =   101
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   79
      Left            =   4830
      TabIndex        =   100
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   78
      Left            =   4095
      TabIndex        =   99
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   77
      Left            =   3465
      TabIndex        =   98
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   76
      Left            =   2835
      TabIndex        =   97
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   75
      Left            =   2100
      TabIndex        =   96
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   74
      Left            =   1470
      TabIndex        =   95
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   73
      Left            =   840
      TabIndex        =   94
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   72
      Left            =   6090
      TabIndex        =   93
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   71
      Left            =   5460
      TabIndex        =   92
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   70
      Left            =   4830
      TabIndex        =   91
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   69
      Left            =   4095
      TabIndex        =   90
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   68
      Left            =   3465
      TabIndex        =   89
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   67
      Left            =   2835
      TabIndex        =   88
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   66
      Left            =   2100
      TabIndex        =   87
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   65
      Left            =   1470
      TabIndex        =   86
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   64
      Left            =   840
      TabIndex        =   85
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   63
      Left            =   6090
      TabIndex        =   84
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   62
      Left            =   5460
      TabIndex        =   83
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   61
      Left            =   4830
      TabIndex        =   82
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   60
      Left            =   4095
      TabIndex        =   81
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   59
      Left            =   3465
      TabIndex        =   80
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   58
      Left            =   2835
      TabIndex        =   79
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   57
      Left            =   2100
      TabIndex        =   78
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   56
      Left            =   1470
      TabIndex        =   77
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   55
      Left            =   840
      TabIndex        =   76
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   54
      Left            =   6090
      TabIndex        =   75
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   53
      Left            =   5460
      TabIndex        =   74
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   52
      Left            =   4830
      TabIndex        =   73
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   51
      Left            =   4095
      TabIndex        =   72
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   50
      Left            =   3465
      TabIndex        =   71
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   49
      Left            =   2835
      TabIndex        =   70
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   48
      Left            =   2100
      TabIndex        =   69
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   47
      Left            =   1470
      TabIndex        =   68
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   46
      Left            =   840
      TabIndex        =   67
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   45
      Left            =   6090
      TabIndex        =   66
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   44
      Left            =   5460
      TabIndex        =   65
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   43
      Left            =   4830
      TabIndex        =   64
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   42
      Left            =   4095
      TabIndex        =   63
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   41
      Left            =   3465
      TabIndex        =   62
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   40
      Left            =   2835
      TabIndex        =   61
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   39
      Left            =   2100
      TabIndex        =   60
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   38
      Left            =   1470
      TabIndex        =   59
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   37
      Left            =   840
      TabIndex        =   58
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   36
      Left            =   6090
      TabIndex        =   57
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   35
      Left            =   5460
      TabIndex        =   56
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   34
      Left            =   4830
      TabIndex        =   55
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   33
      Left            =   4095
      TabIndex        =   54
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   32
      Left            =   3465
      TabIndex        =   53
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   31
      Left            =   2835
      TabIndex        =   52
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   30
      Left            =   2100
      TabIndex        =   51
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   29
      Left            =   1470
      TabIndex        =   50
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   28
      Left            =   840
      TabIndex        =   49
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   27
      Left            =   6090
      TabIndex        =   48
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   26
      Left            =   5460
      TabIndex        =   47
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   25
      Left            =   4830
      TabIndex        =   46
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   24
      Left            =   4095
      TabIndex        =   45
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   23
      Left            =   3465
      TabIndex        =   44
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   22
      Left            =   2835
      TabIndex        =   43
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   21
      Left            =   2100
      TabIndex        =   42
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   20
      Left            =   1470
      TabIndex        =   41
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   19
      Left            =   840
      TabIndex        =   40
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   18
      Left            =   6090
      TabIndex        =   39
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   17
      Left            =   5460
      TabIndex        =   38
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   16
      Left            =   4830
      TabIndex        =   37
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   15
      Left            =   4095
      TabIndex        =   36
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   14
      Left            =   3465
      TabIndex        =   35
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   13
      Left            =   2835
      TabIndex        =   34
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   12
      Left            =   2100
      TabIndex        =   33
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   11
      Left            =   1470
      TabIndex        =   32
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   10
      Left            =   840
      TabIndex        =   31
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   9
      Left            =   6090
      TabIndex        =   30
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   8
      Left            =   5460
      TabIndex        =   29
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   7
      Left            =   4830
      TabIndex        =   28
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   6
      Left            =   4095
      TabIndex        =   27
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   5
      Left            =   3465
      TabIndex        =   26
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   4
      Left            =   2835
      TabIndex        =   25
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   3
      Left            =   2100
      TabIndex        =   24
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   81
      Left            =   6090
      TabIndex        =   102
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   2
      Left            =   1470
      TabIndex        =   23
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   1
      Left            =   840
      TabIndex        =   22
      Top             =   945
      Width           =   510
   End
   Begin VB.Shape ShapeBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   120
      Left            =   7035
      Top             =   6615
      Width           =   6210
   End
   Begin VB.Label LabelStatusbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ready"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7035
      TabIndex        =   132
      Top             =   6195
      Width           =   6210
   End
   Begin VB.Menu MenuSoundSwitch 
      Caption         =   "Soun&d ON"
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "&About"
      Begin VB.Menu MenuAboutName 
         Caption         =   "Random Sudoku Generator"
      End
      Begin VB.Menu MenuAboutVersion 
         Caption         =   "v0.11 Beta Version　|　for Windows 7,8,10　|　English (US)"
      End
      Begin VB.Menu MenuAboutDate 
         Caption         =   "Last compiled on Thu, Sep 24, 2020"
      End
      Begin VB.Menu MenuAboutFirst 
         Caption         =   "First version built on Sat, Mar 28, 2020"
      End
      Begin VB.Menu MenuAbout1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutAuthor 
         Caption         =   "Author: Sam Toki"
      End
      Begin VB.Menu MenuAboutOrganization 
         Caption         =   "Organization: SAM TOKI STUDIO"
      End
      Begin VB.Menu MenuAboutFrom 
         Caption         =   "From: Xidian University, China"
      End
      Begin VB.Menu MenuAboutContact 
         Caption         =   "Contact: SamToki@outlook.com"
      End
      Begin VB.Menu MenuAbout2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutCopyright 
         Caption         =   "TM ＆ (C) 2015-2020 SAM TOKI STUDIO. All rights reserved."
      End
      Begin VB.Menu MenuAboutTrademark 
         Caption         =   "SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries."
      End
      Begin VB.Menu MenuAbout3_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAboutCommercial 
         Caption         =   "Commercial use of this software is strictly prohibited."
      End
   End
   Begin VB.Menu Menu1_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "Ａ字あ (&L)"
      Begin VB.Menu MenuLanguageENG 
         Caption         =   "English (United States)"
         Checked         =   -1  'True
         Shortcut        =   +{F1}
      End
      Begin VB.Menu MenuLanguageCHS 
         Caption         =   "中文（简体）"
         Enabled         =   0   'False
         Shortcut        =   +{F2}
      End
      Begin VB.Menu MenuLanguageCHT 
         Caption         =   "中文（繁體）"
         Enabled         =   0   'False
         Shortcut        =   +{F3}
      End
      Begin VB.Menu MenuLanguageJPN 
         Caption         =   "日本語"
         Enabled         =   0   'False
         Shortcut        =   +{F4}
      End
   End
   Begin VB.Menu Menu2_ 
      Caption         =   "　|　"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuEXIT 
      Caption         =   "E&XIT"
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === INFORMATION ===
'
'  SAM TOKI STUDIO
'  This is a .frm source code file.
'
'  Random Sudoku Generator
'
'  Powered by Sam Toki
'  Version: v0.11 Beta Version ENG
'  Date:    09/20/2020 (Sun.)
'  History: First version v0.10 Beta was built on 03/28/2020.
'
'  WARNING: Commercial use of this computer software is strictly prohibited.
'           Open source license:      GNU GPL v3
'           Creative Commons license: CC BY-NC 3.0
'
'  Copyright: TM & (C) 2015-2020 SAM TOKI STUDIO. All rights reserved.
'             SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries.
'
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === NOTES FOR REFERENCE ===
'
'  ...
'
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public setlanguage As String
Public soundswitch As Boolean

'Declare Game...
Public gamestatus As Integer  '0-Initial, 1-Generating, 2-Ongoing.
Public gameinputstep As Integer  '1-Row, 2-Column, 3-Number.

'Declare Sudoku...
Public sudokublockdata As Variant  '(1 To 9)(1 To 9)
Public sudokublockstatus As Variant  '(1 To 9)(1 To 9)  0-Filling, 1-Fixed, 2-FillingAgainstRule, 3-FixedAgainstRule.

Public sudokutotalfixed As Integer
Public sudokutotalfilling As Integer
Public sudokutotalgenerated As Integer
Public sudokutotalfilled As Integer
Public sudokulargeblockmaximumfixed As Integer
Public sudokulargeblockfixedcount As Variant  '(1 To 3)(1 To 3)
Public sudokufillednumbercount As Variant  '(1 To 9)

Public sudokucurrentrow As Integer
Public sudokucurrentcolumn As Integer

'Declare Lottery...
Public lotterytotal As Integer
Public lotterynumber As Integer

'Declare Animation...
Public progressbaranimationtarget As Long  'Range: 0~6210
Public labelrowindicatoranimationtarget As Long  'Range: 364~6270
Public labelcolumnindicatoranimationtarget As Long  'Range: 289~6195
Public sudokucurrentrowanimation As Integer
Public sudokucurrentcolumnanimation As Integer

'Declare Dialog...
Public answer
Public dontannoymeagain1 As Boolean

'Declare Others...
Public setanimationswitch As Boolean

'Declare Temp...
Public forloop1 As Integer
Public forloop2 As Integer
Public forloop3 As Integer
Public forloop4 As Integer
Public forloop5 As Integer
Public tempvariant As Integer
Public preventinfiniteloopcounter1 As Integer
Public preventinfiniteloopcounter2 As Integer

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Sub Form_Load()
        'Load and Initialization...

        'Initialize Menu...
        setlanguage = "ENG"
        soundswitch = True

        'Initialize Game...
        gamestatus = 0
        gameinputstep = 0

        'Initialize Sudoku...
        sudokublockdata = Array(Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                )
        sudokublockstatus = Array(Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                  )

        sudokutotalfixed = LabelSettingsTotalFixedAmountIndicator.Caption
        sudokutotalfilling = 81 - sudokutotalfixed
        sudokutotalgenerated = 0
        sudokutotalfilled = 0
        sudokulargeblockmaximumfixed = LabelSettingsLargeBlockMaximumFixedAmountIndicator.Caption
        sudokulargeblockfixedcount = Array(Array(0, 0, 0, 0), _
                                           Array(0, 0, 0, 0), _
                                           Array(0, 0, 0, 0), _
                                           Array(0, 0, 0, 0) _
                                           )
        sudokufillednumbercount = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        sudokucurrentrow = 0
        sudokucurrentcolumn = 0

        'Initialize Lottery...
        lotterytotal = 0
        lotterynumber = 0

        'Initialize Animation...
        progressbaranimationtarget = 0
        labelrowindicatoranimationtarget = 364
        labelcolumnindicatoranimationtarget = 289
        sudokucurrentrowanimation = 0
        sudokucurrentcolumnanimation = 0

        'Initialize Dialog...
        dontannoymeagain1 = False

        'Initialize Others...
        setanimationswitch = True
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'CMD Language...
    Public Sub MenuLanguageENG_Click()
        'Call ModuleLoadLanguage.LoadLanguageENG
    End Sub
    Public Sub MenuLanguageCHS_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHS
    End Sub
    Public Sub MenuLanguageCHT_Click()
        'Call ModuleLoadLanguage.LoadLanguageCHT
    End Sub
    Public Sub MenuLanguageJPN_Click()
        'Call ModuleLoadLanguage.LoadLanguageJPN
    End Sub

    'CMD Menu...
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub CmdEXIT_Click()
        Call MenuEXIT_Click
    End Sub
    Public Sub MenuSoundSwitch_Click()
        Select Case soundswitch
            Case True
                soundswitch = False
                MenuSoundSwitch.Caption = "Soun&d OFF"
            Case False
                soundswitch = True
                MenuSoundSwitch.Caption = "Soun&d ON"
        End Select
    End Sub

    'CMD Settings...
    Public Sub HScrollSettingsTotalFixedAmount_Change()
        sudokutotalfixed = HScrollSettingsTotalFixedAmount.Value
        LabelSettingsTotalFixedAmountIndicator.Caption = sudokutotalfixed
    End Sub
    Public Sub HScrollSettingsTotalFixedAmount_Scroll()
        Call HScrollSettingsTotalFixedAmount_Change
    End Sub
    Public Sub HScrollSettingsLargeBlockMaximumFixedAmount_Change()
        sudokulargeblockmaximumfixed = HScrollSettingsLargeBlockMaximumFixedAmount.Value
        LabelSettingsLargeBlockMaximumFixedAmountIndicator.Caption = sudokulargeblockmaximumfixed
        HScrollSettingsTotalFixedAmount.Max = sudokulargeblockmaximumfixed * 9
    End Sub
    Public Sub HScrollSettingsLargeBlockMaximumFixedAmount_Scroll()
        Call HScrollSettingsLargeBlockMaximumFixedAmount_Change
    End Sub

    'CMD Controls...
    Public Sub CmdStartReset_Click()
        If gamestatus = 0 Then
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Connection.wav"
            Call Form_Load: gamestatus = 1: Call Refresher
            LabelStatusbar.Caption = "Game started!"
        Else
            CmdStartReset.SetFocus
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Recycle.wav"
            Call Form_Load: gamestatus = 0: Call Refresher
            LabelStatusbar.Caption = "Game reset! Now you can adjust settings."
        End If
    End Sub
    Public Sub CmdNumber_Click(index As Integer)
        TextboxInput.Text = index
    End Sub
    Public Sub LabelSudokuBlock_Click(index As Integer)
        sudokucurrentrow = -Int(-index / 9)
        If (index Mod 9 = 0) Then
            sudokucurrentcolumn = 9
        Else
            sudokucurrentcolumn = index Mod 9
        End If

        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Menu Command.wav"
        Call Refresher: gameinputstep = 3: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1
        If (Not (gamestatus = 2)) Then Exit Sub
        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
    End Sub

    Public Sub TextboxInput_Change()
        If (Not (gamestatus = 2)) Then Exit Sub
        'If the change is to clear the textbox, then do nothing...
        If TextboxInput.Text = "" Then Exit Sub

        If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Startup.wav"

        Select Case gameinputstep
            Case 1
                Select Case TextboxInput.Text
                    Case 1
                        sudokucurrentrow = 1
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 2
                        sudokucurrentrow = 2
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 3
                        sudokucurrentrow = 3
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 4
                        sudokucurrentrow = 4
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 5
                        sudokucurrentrow = 5
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 6
                        sudokucurrentrow = 6
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 7
                        sudokucurrentrow = 7
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 8
                        sudokucurrentrow = 8
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case 9
                        sudokucurrentrow = 9
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 2/3: Now enter column number."
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case ""
                        TextboxInput.Text = "": gameinputstep = 1: Exit Sub
                    Case Else
                        MsgBox "CAUTION: Invalid input. You have pressed a wrong key. Please confirm that your fingers are on the right keys." & vbCrLf & vbCrLf & "NOTE: Acceptable keys are from 1 to 9.", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                        TextboxInput.Text = "": gameinputstep = 1: Exit Sub
                End Select
            Case 2
                Select Case TextboxInput.Text
                    Case 1
                        sudokucurrentcolumn = 1
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 2
                        sudokucurrentcolumn = 2
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 3
                        sudokucurrentcolumn = 3
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 4
                        sudokucurrentcolumn = 4
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 5
                        sudokucurrentcolumn = 5
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 6
                        sudokucurrentcolumn = 6
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 7
                        sudokucurrentcolumn = 7
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 8
                        sudokucurrentcolumn = 8
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case 9
                        sudokucurrentcolumn = 9
                        LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 3/3: Now fill the block."
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case ""
                        TextboxInput.Text = "": gameinputstep = 2: Exit Sub
                    Case Else
                        MsgBox "CAUTION: Invalid input. You have pressed a wrong key. Please confirm that your fingers are on the right keys." & vbCrLf & vbCrLf & "NOTE: Acceptable keys are from 1 to 9.", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                        TextboxInput.Text = "": gameinputstep = 2: Exit Sub
                End Select
            Case 3
                If (((sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 1) Or (sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 3)) And (dontannoymeagain1 = False)) Then
                    answer = MsgBox("CAUTION: You are trying to change the number in a fixed block." & vbCrLf & "Are you sure you want to proceed?" & vbCrLf & vbCrLf & "[Yes]  Sure, just proceed" & vbCrLf & "[No]  Do not proceed, leave that number unchanged" & vbCrLf & "[Cancel]  Just do it, and don't annoy me again", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Random Sudoku Generator")
                    If answer = vbNo Then
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    End If
                    If answer = vbCancel Then dontannoymeagain1 = True
                End If
                Select Case TextboxInput.Text
                    Case 1
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 1
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 2
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 2
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 3
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 3
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 4
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 4
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 5
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 5
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 6
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 6
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 7
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 7
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 8
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 8
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 9
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 9
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case 0
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuChecker: Call Refresher: Exit Sub
                    Case ""
                        TextboxInput.Text = "": gameinputstep = 3: Call SudokuChecker: Call Refresher: Exit Sub
                    Case Else
                        MsgBox "CAUTION: Invalid input. You have pressed a wrong key. Please confirm that your fingers are on the right keys." & vbCrLf & vbCrLf & "NOTE: Acceptable keys are from 1 to 9, and 0 for ""Clear"", Enter for ""Start New Game"".", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                        TextboxInput.Text = "": gameinputstep = 3: Call SudokuChecker: Call Refresher: Exit Sub
                End Select
        End Select
    End Sub

'[] TIMERS []

    Public Sub TimerClock_Timer()
        LabelClock.Caption = Format((Hour(Time)), "00") & ":" & Format((Minute(Time)), "00") & ":" & Format((Second(Time)), "00")
    End Sub

    Public Sub TimerSudokuGenerator_Timer()
        If Not (gamestatus = 1) Then Exit Sub

        Call SudokuGenerator

        'Finish generation...
        If sudokutotalgenerated = sudokutotalfixed Then
            If soundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Print Complete.wav"
            gamestatus = 2: gameinputstep = 1: sudokucurrentrow = 0: sudokucurrentcolumn = 0: Call Refresher: TextboxInput.SetFocus
        End If
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ENGINE []

    'Refresh statistics, display, and commands...
    Public Sub Refresher()
        'Statistics...
        Select Case gamestatus
            Case 0
                sudokutotalgenerated = 0
                sudokutotalfilled = 0
            Case 1
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If (((sudokublockstatus(forloop1)(forloop2) = 1) Or (sudokublockstatus(forloop1)(forloop2) = 3)) And (Not (sudokublockdata(forloop1)(forloop2) = 0))) Then tempvariant = tempvariant + 1
                    Next
                Next
                sudokutotalgenerated = tempvariant
                sudokutotalfilled = 0
            Case 2
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If ((Not ((sudokublockstatus(forloop1)(forloop2) = 1) Or (sudokublockstatus(forloop1)(forloop2) = 3))) And (Not (sudokublockdata(forloop1)(forloop2) = 0))) Then tempvariant = tempvariant + 1
                    Next
                Next
                sudokutotalgenerated = sudokutotalfixed
                sudokutotalfilled = tempvariant
                LabelStatusbar = "Game ongoing... --- Filled: " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Step 1/3: Please enter row number."

                'Judge Sudoku solved...
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If ((sudokublockstatus(forloop1)(forloop2) = 2) Or (sudokublockstatus(forloop1)(forloop2) = 3)) Then tempvariant = 444
                    Next
                Next
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If sudokublockdata(forloop1)(forloop2) = 0 Then tempvariant = 888
                    Next
                Next
                If ((sudokutotalfilled >= sudokutotalfilling) And (tempvariant = 0) And (gameinputstep = 1)) Then MsgBox "Congratulations!! You have solved the Sudoku.", vbInformation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select

        'Row and column number...
        For forloop1 = 0 To 9: LabelRow(forloop1).ForeColor = &HFFFFFF: Next
        If (Not (sudokucurrentrow = 0)) Then LabelRow(sudokucurrentrow).ForeColor = &H0&
        For forloop1 = 0 To 9: LabelColumn(forloop1).ForeColor = &HFFFFFF: Next
        If (Not (sudokucurrentcolumn = 0)) Then LabelColumn(sudokucurrentcolumn).ForeColor = &H0&

        'Sudoku blocks...
        For forloop1 = 1 To 81
            If (forloop1 Mod 9 = 0) Then
                tempvariant = 9
            Else
                tempvariant = forloop1 Mod 9
            End If
            Select Case sudokublockstatus(-Int(-forloop1 / 9))(tempvariant)
                Case 0
                    LabelSudokuBlock(forloop1).BackColor = &HFFFFFF: LabelSudokuBlock(forloop1).ForeColor = &H0&
                Case 1
                    LabelSudokuBlock(forloop1).BackColor = &HE0E0E0: LabelSudokuBlock(forloop1).ForeColor = &H0&
                Case 2
                    LabelSudokuBlock(forloop1).BackColor = &HFFFFFF: LabelSudokuBlock(forloop1).ForeColor = &HFF&
                Case 3
                    LabelSudokuBlock(forloop1).BackColor = &HE0E0E0: LabelSudokuBlock(forloop1).ForeColor = &HFF&
            End Select
            LabelSudokuBlock(forloop1).BorderStyle = 0
        Next

        For forloop1 = 1 To 81
            If (forloop1 Mod 9 = 0) Then
                tempvariant = 9
            Else
                tempvariant = forloop1 Mod 9
            End If
            LabelSudokuBlock(forloop1).Caption = sudokublockdata(-Int(-forloop1 / 9))(tempvariant)
            If LabelSudokuBlock(forloop1).Caption = "0" Then LabelSudokuBlock(forloop1).Caption = ""
        Next

        'Number count...
        For forloop1 = 1 To 9
            tempvariant = 0

            For forloop2 = 1 To 9
                For forloop3 = 1 To 9
                    If sudokublockdata(forloop2)(forloop3) = forloop1 Then tempvariant = tempvariant + 1
                Next
            Next

            sudokufillednumbercount(forloop1) = tempvariant
            LabelNumberCount(forloop1).Caption = tempvariant
            If tempvariant > 9 Then
                LabelNumberCount(forloop1).Caption = "X"
            End If
            If tempvariant >= 9 Then
                If tempvariant = 9 Then
                    LabelNumberCount(forloop1).ForeColor = &HC000&
                Else
                    LabelNumberCount(forloop1).ForeColor = &HFF&
                End If
            Else
                LabelNumberCount(forloop1).ForeColor = &H0&
            End If
        Next

        'Game status light bulb indicator...
        Select Case gamestatus
            Case 0
                ShapeLightGeneratingIndicator.FillStyle = 1
                ShapeLightGameOngoingIndicator.FillStyle = 1
            Case 1
                ShapeLightGeneratingIndicator.FillStyle = 0
                ShapeLightGameOngoingIndicator.FillStyle = 1
            Case 2
                ShapeLightGeneratingIndicator.FillStyle = 1
                ShapeLightGameOngoingIndicator.FillStyle = 0
        End Select

        'Progressbar...
        Select Case gamestatus
            Case 0
                ShapeProgressbar.BackColor = &HFF8800
                progressbaranimationtarget = 0
            Case 1
                ShapeProgressbar.BackColor = &HFF8800
                progressbaranimationtarget = (sudokutotalgenerated / sudokutotalfixed) * 6210
            Case 2
                ShapeProgressbar.BackColor = &HC000&
                progressbaranimationtarget = (sudokutotalfilled / sudokutotalfilling) * 6210
        End Select

        'Commands...
        If Not (gamestatus = 2) Then
            TextboxInput.Enabled = False: For forloop1 = 0 To 9: CmdNumber(forloop1).Enabled = False: Next
        Else
            TextboxInput.Enabled = True: For forloop1 = 0 To 9: CmdNumber(forloop1).Enabled = True: Next
        End If

        If gamestatus = 0 Then
            HScrollSettingsTotalFixedAmount.Enabled = True: HScrollSettingsLargeBlockMaximumFixedAmount.Enabled = True
        Else
            HScrollSettingsTotalFixedAmount.Enabled = False: HScrollSettingsLargeBlockMaximumFixedAmount.Enabled = False
        End If
    End Sub

    'Check Sudoku...
    Public Sub LargeBlockChecker()
        tempvariant = 0

        Select Case sudokucurrentrow
            Case 1 To 3
                Select Case sudokucurrentcolumn
                    Case 1 To 3
                        For forloop1 = 1 To 3
                            For forloop2 = 1 To 3
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 4 To 6
                        For forloop1 = 1 To 3
                            For forloop2 = 4 To 6
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 7 To 9
                        For forloop1 = 1 To 3
                            For forloop2 = 7 To 9
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                End Select
            Case 4 To 6
                Select Case sudokucurrentcolumn
                    Case 1 To 3
                        For forloop1 = 4 To 6
                            For forloop2 = 1 To 3
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 4 To 6
                        For forloop1 = 4 To 6
                            For forloop2 = 4 To 6
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 7 To 9
                        For forloop1 = 4 To 6
                            For forloop2 = 7 To 9
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                End Select
            Case 7 To 9
                Select Case sudokucurrentcolumn
                    Case 1 To 3
                        For forloop1 = 7 To 9
                            For forloop2 = 1 To 3
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 4 To 6
                        For forloop1 = 7 To 9
                            For forloop2 = 4 To 6
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 7 To 9
                        For forloop1 = 7 To 9
                            For forloop2 = 7 To 9
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= sudokulargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                End Select
        End Select
    End Sub

    Public Sub SudokuChecker()
        'Initialize sudokublockstatus...
        For forloop1 = 1 To 9
            For forloop2 = 1 To 9
                If sudokublockstatus(forloop1)(forloop2) = 2 Then sudokublockstatus(forloop1)(forloop2) = 0
                If sudokublockstatus(forloop1)(forloop2) = 3 Then sudokublockstatus(forloop1)(forloop2) = 1
            Next
        Next

        'Locate block...
        For forloop1 = 1 To 9
            For forloop2 = 1 To 9

                'Check horizontal numbers...
                For forloop3 = 1 To 9
                    If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop1)(forloop3)) And (Not (forloop2 = forloop3)) Then
                        Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                    End If
                Next

                'Check vertical numbers...
                For forloop3 = 1 To 9
                    If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop2)) And (Not (forloop1 = forloop3)) Then
                        Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                    End If
                Next

                'Check large block...
                Select Case forloop1
                    Case 1 To 3
                        Select Case forloop2
                            Case 1 To 3
                                For forloop3 = 1 To 3
                                    For forloop4 = 1 To 3
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 4 To 6
                                For forloop3 = 1 To 3
                                    For forloop4 = 4 To 6
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 7 To 9
                                For forloop3 = 1 To 3
                                    For forloop4 = 7 To 9
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                        End Select
                    Case 4 To 6
                        Select Case forloop2
                            Case 1 To 3
                                For forloop3 = 4 To 6
                                    For forloop4 = 1 To 3
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 4 To 6
                                For forloop3 = 4 To 6
                                    For forloop4 = 4 To 6
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 7 To 9
                                For forloop3 = 4 To 6
                                    For forloop4 = 7 To 9
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                        End Select
                    Case 7 To 9
                        Select Case forloop2
                            Case 1 To 3
                                For forloop3 = 7 To 9
                                    For forloop4 = 1 To 3
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 4 To 6
                                For forloop3 = 7 To 9
                                    For forloop4 = 4 To 6
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 7 To 9
                                For forloop3 = 7 To 9
                                    For forloop4 = 7 To 9
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                        End Select
                End Select

            Next
        Next
    End Sub

    'Generate Sudoku...
    Public Sub RandomNumberGenerator()
        If lotterytotal = 0 Then
            MsgBox "ERROR: Calling function ""RandomNumberGenerator"" when variant ""lotterytotal"" is 0." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End If

        lotterynumber = 0
        While lotterynumber = 0
            Randomize
            lotterynumber = Int((lotterytotal + 1) * Rnd)
        Wend
    End Sub

    Public Sub SudokuGenerator()
        If Not (gamestatus = 1) Then Exit Sub

        'Prevent infinite loop...
        preventinfiniteloopcounter1 = 0

SudokuGenerator_Relocate_:

        'Maximum toleration of infinite loop...
        preventinfiniteloopcounter1 = preventinfiniteloopcounter1 + 1
        If preventinfiniteloopcounter1 > 100 Then
            'Abort generation...
            MsgBox "WARNING: Unable to continue building this Sudoku grid anymore." & vbCrLf & "Random Sudoku generation has failed." & vbCrLf & "Please try to adjust the settings so as to make it harder to fail." & vbCrLf & "We will abort the generation and reset the game later.", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
            Call CmdStartReset_Click: Exit Sub
        End If

        'Locate random block...
        LabelStatusbar.Caption = "Generating... " & sudokutotalgenerated & "/" & sudokutotalfixed & " --- Locating a new block..."
        lotterytotal = 9: lotterynumber = 0
        Do
            Call RandomNumberGenerator: sudokucurrentrow = lotterynumber
            Call RandomNumberGenerator: sudokucurrentcolumn = lotterynumber
        Loop Until sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0

        'Check rules...
        Call LargeBlockChecker
        If tempvariant = 444 Then GoTo SudokuGenerator_Relocate_

        'Prevent infinite loop...
        preventinfiniteloopcounter2 = 0

SudokuGenerator_Refill_:

        'Maximum toleration of infinite loop...
        preventinfiniteloopcounter2 = preventinfiniteloopcounter2 + 1
        If preventinfiniteloopcounter2 > 10 Then
            sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0
            sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 0
            GoTo SudokuGenerator_Relocate_
        End If

        'Fill and fix random block...
        LabelStatusbar.Caption = "Generating... " & sudokutotalgenerated & "/" & sudokutotalfixed & " --- Filling with random number..."
        lotterytotal = 9: lotterynumber = 0
        Do
            Call RandomNumberGenerator
            sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = lotterynumber
            sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 1
            Call Refresher
        Loop Until sudokufillednumbercount(lotterynumber) <= 9

        'Check rules...
        LabelStatusbar.Caption = "Generating... " & sudokutotalgenerated & "/" & sudokutotalfixed & " --- Checking..."
        Call SudokuChecker
        Call Refresher
        If sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 3 Then GoTo SudokuGenerator_Refill_

        'OK...
        preventinfiniteloopcounter1 = 0: preventinfiniteloopcounter2 = 0
        LabelStatusbar.Caption = "Generating... " & sudokutotalgenerated & "/" & sudokutotalfixed & " --- OK..."
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerProgressbarAnimation_Timer()
        labelrowindicatoranimationtarget = 364 + (sudokucurrentrow / 9) * (6270 - 364)
        labelcolumnindicatoranimationtarget = 289 + (sudokucurrentcolumn / 9) * (6195 - 289)

        If ShapeProgressbar.Width = progressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
        If ShapeProgressbar.Width > progressbaranimationtarget Then ShapeProgressbar.Width = ShapeProgressbar.Width - Abs(ShapeProgressbar.Width - progressbaranimationtarget) / 4
        If ShapeProgressbar.Width < progressbaranimationtarget Then ShapeProgressbar.Width = ShapeProgressbar.Width + Abs(ShapeProgressbar.Width - progressbaranimationtarget) / 4
        If Abs(ShapeProgressbar.Width - progressbaranimationtarget) < 10 Then ShapeProgressbar.Width = progressbaranimationtarget
TimerProgressbarAnimation_Skip1_:

        If LabelRowIndicator.Top = labelrowindicatoranimationtarget Then GoTo TimerProgressbarAnimation_Skip2_
        If LabelRowIndicator.Top > labelrowindicatoranimationtarget Then LabelRowIndicator.Top = LabelRowIndicator.Top - Abs(LabelRowIndicator.Top - labelrowindicatoranimationtarget) / 4
        If LabelRowIndicator.Top < labelrowindicatoranimationtarget Then LabelRowIndicator.Top = LabelRowIndicator.Top + Abs(LabelRowIndicator.Top - labelrowindicatoranimationtarget) / 4
        If Abs(LabelRowIndicator.Top - labelrowindicatoranimationtarget) < 10 Then LabelRowIndicator.Top = labelrowindicatoranimationtarget
TimerProgressbarAnimation_Skip2_:

        If LabelColumnIndicator.Left = labelcolumnindicatoranimationtarget Then GoTo TimerProgressbarAnimation_Skip3_
        If LabelColumnIndicator.Left > labelcolumnindicatoranimationtarget Then LabelColumnIndicator.Left = LabelColumnIndicator.Left - Abs(LabelColumnIndicator.Left - labelcolumnindicatoranimationtarget) / 4
        If LabelColumnIndicator.Left < labelcolumnindicatoranimationtarget Then LabelColumnIndicator.Left = LabelColumnIndicator.Left + Abs(LabelColumnIndicator.Left - labelcolumnindicatoranimationtarget) / 4
        If Abs(LabelColumnIndicator.Left - labelcolumnindicatoranimationtarget) < 10 Then LabelColumnIndicator.Left = labelcolumnindicatoranimationtarget
TimerProgressbarAnimation_Skip3_:

    End Sub

    Public Sub TimerSudokuBlockAnimation_Timer()
        'DISABLED LINE: If Not (gamestatus = 2) Then Exit Sub

        'Highlight current row and column with light blue color, animated. And highlight the row and column number...
        Select Case gameinputstep
            Case 0
                Exit Sub
            Case 1
                Exit Sub
            Case 2
                LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentrowanimation).BackColor = &HFFF0E0
                If sudokucurrentrowanimation < 9 Then sudokucurrentrowanimation = sudokucurrentrowanimation + 1
                LabelRow(sudokucurrentrow).ForeColor = &H0&
            Case 3
                LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentrowanimation).BackColor = &HFFF0E0
                If sudokucurrentrowanimation < 9 Then sudokucurrentrowanimation = sudokucurrentrowanimation + 1
                LabelRow(sudokucurrentrow).ForeColor = &H0&
                LabelSudokuBlock(sudokucurrentcolumnanimation * 9 - 9 + sudokucurrentcolumn).BackColor = &HFFF0E0
                If sudokucurrentcolumnanimation < 9 Then sudokucurrentcolumnanimation = sudokucurrentcolumnanimation + 1
                LabelColumn(sudokucurrentcolumn).ForeColor = &H0&
            Case Else
                MsgBox "ERROR: Game input step is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select

        'Highlight current block with light green or red color...
        If gameinputstep = 3 Then
            'Show block border...
            LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BorderStyle = 1

            'Change backcolor...
            Select Case sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn)
                Case 0
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0FFC0
                Case 1
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0C0FF
                Case 2
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0FFC0
                Case 3
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0C0FF
                Case Else
                    MsgBox "ERROR: Sudoku block fixed-or-not data is out of range." & vbCrLf & "Please send a feedback to us so as to help solve the problem. Thank you very much.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
            End Select
        End If
    End Sub
