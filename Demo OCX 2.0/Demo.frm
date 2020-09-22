VERSION 5.00
Begin VB.Form Demo 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Demo Ocx"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10500
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Demo.frx":030A
   ScaleHeight     =   7470
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin Progetto1.HSlider HSlider3 
      Height          =   135
      Left            =   480
      TabIndex        =   142
      Top             =   6420
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   238
      MaxValue        =   100
      Picture         =   "Demo.frx":FF876
      PicCursor_Selected=   "Demo.frx":100DE2
      PictureCursor   =   "Demo.frx":101106
   End
   Begin Progetto1.HSlider HSlider2 
      Height          =   255
      Left            =   360
      TabIndex        =   141
      Top             =   4680
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      Value           =   50
      MaxValue        =   100
      Picture         =   "Demo.frx":10142A
      PicCursor_Selected=   "Demo.frx":102996
      PictureCursor   =   "Demo.frx":102D6E
   End
   Begin Progetto1.HSlider HSlider1 
      Height          =   255
      Left            =   360
      TabIndex        =   140
      Top             =   4080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      Value           =   50
      MaxValue        =   100
      Picture         =   "Demo.frx":103146
      PicCursor_Selected=   "Demo.frx":1046B2
      PictureCursor   =   "Demo.frx":104A8A
   End
   Begin Progetto1.Led LedBlink 
      Height          =   255
      Index           =   0
      Left            =   4740
      TabIndex        =   138
      Top             =   620
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Colore          =   1
   End
   Begin VB.TextBox Txt_VSlider1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   6600
      TabIndex        =   137
      Text            =   "0"
      Top             =   640
      Width           =   615
   End
   Begin VB.Frame FrmBarra_R 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   7080
      TabIndex        =   124
      Top             =   4080
      Width           =   255
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   125
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   126
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   127
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   128
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   129
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   130
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   131
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   132
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   3
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   133
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   3
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   134
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin Progetto1.Led Led_Barra_R 
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   135
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
   End
   Begin VB.Frame FrmBarra_L 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   6240
      TabIndex        =   112
      Top             =   4080
      Width           =   255
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   113
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   114
         Top             =   2160
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   115
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   116
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   117
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   118
         Top             =   1200
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   119
         Top             =   960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   1
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   120
         Top             =   720
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   3
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   121
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Colore          =   3
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   122
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin Progetto1.Led Led_Barra_L 
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   123
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
   End
   Begin Progetto1.V_UpDown V_UpDown3 
      Height          =   1575
      Left            =   3840
      TabIndex        =   111
      Top             =   5040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2778
      ForeColor       =   4210752
      MaxValue        =   10
   End
   Begin VB.Timer Timer_Led 
      Interval        =   50
      Left            =   2640
      Top             =   120
   End
   Begin Progetto1.VSlider VSlider1 
      Height          =   2655
      Index           =   0
      Left            =   5865
      TabIndex        =   110
      ToolTipText     =   "DownToUp = False"
      Top             =   4080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4683
      Value           =   5
      DownToUp        =   0   'False
      Picture         =   "Demo.frx":104E62
      PictureCursor   =   "Demo.frx":106506
   End
   Begin Progetto1.H_UpDown H_UpDown2 
      Height          =   255
      Left            =   2640
      TabIndex        =   109
      Top             =   4380
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      MaxValue        =   2
      LoopValue       =   -1  'True
   End
   Begin Progetto1.V_UpDown V_UpDown1 
      Height          =   855
      Index           =   0
      Left            =   3855
      TabIndex        =   105
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1508
      BackColor       =   16659743
      ForeColor       =   16777152
      MaxValue        =   9
      LoopValue       =   -1  'True
   End
   Begin VB.CheckBox CheckBinary 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   3080
      TabIndex        =   103
      Top             =   660
      Width           =   200
   End
   Begin VB.Frame FrmByte 
      Appearance      =   0  'Flat
      BackColor       =   &H00C8B4AC&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   360
      TabIndex        =   88
      Top             =   5040
      Width           =   3375
      Begin VB.Frame FrmLed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   89
         Top             =   960
         Width           =   2895
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   90
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   91
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   92
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   3
            Left            =   1080
            TabIndex        =   93
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   4
            Left            =   1440
            TabIndex        =   94
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   5
            Left            =   1800
            TabIndex        =   95
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   6
            Left            =   2160
            TabIndex        =   96
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
         Begin Progetto1.Led Led1 
            Height          =   375
            Index           =   7
            Left            =   2520
            TabIndex        =   97
            Top             =   0
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            Colore          =   1
         End
      End
      Begin Progetto1.H_UpDown H_UpDown9 
         Height          =   255
         Left            =   1920
         TabIndex        =   98
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16761024
         MaxValue        =   255
         LoopValue       =   -1  'True
      End
      Begin Progetto1.H_UpDown H_UpDown10 
         Height          =   255
         Left            =   1920
         TabIndex        =   101
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BackColor       =   16761024
         Value           =   1
         MaxValue        =   3
         LoopValue       =   -1  'True
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Led Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   360
         Width           =   1695
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   120
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Binary Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0     1      2      3     4      5      6      7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   390
         TabIndex        =   99
         Top             =   720
         Width           =   2655
      End
      Begin VB.Shape Shape4 
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   3375
      End
   End
   Begin Progetto1.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   1560
      TabIndex        =   87
      Top             =   3480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      ValueColor      =   16659743
      PictureBackG    =   "Demo.frx":106806
      PictureForG     =   "Demo.frx":10B9A2
   End
   Begin Progetto1.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   86
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      ValueColor      =   16659743
      PictureBackG    =   "Demo.frx":110B3E
      PictureForG     =   "Demo.frx":115CDA
   End
   Begin VB.TextBox TxtUpDown 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   84
      Text            =   "0"
      Top             =   1100
      Width           =   735
   End
   Begin Progetto1.H_UpDown H_UpDown7 
      Height          =   255
      Left            =   6960
      TabIndex        =   81
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777152
      ForeColor       =   16659743
      MaxValue        =   3
   End
   Begin VB.Timer TimerDisplay_User 
      Interval        =   50
      Left            =   2160
      Top             =   120
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   61
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin VB.TextBox Txt_User 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "Put here your  text"
      Top             =   1130
      Width           =   4695
   End
   Begin VB.TextBox Txt_VSlider1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   5640
      TabIndex        =   59
      Text            =   "0"
      Top             =   640
      Width           =   615
   End
   Begin Progetto1.H_UpDown H_UpDown6 
      Height          =   255
      Left            =   6960
      TabIndex        =   36
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777152
      ForeColor       =   16659743
      Value           =   1
      MinValue        =   1
   End
   Begin VB.Timer TimerDisplay 
      Interval        =   60
      Left            =   1680
      Top             =   120
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed DisplayLed2 
      Height          =   1320
      Left            =   8640
      TabIndex        =   13
      Top             =   2280
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Zoom            =   4
   End
   Begin Progetto1.H_UpDown H_UpDown5 
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777152
      ForeColor       =   16659743
      Value           =   4
      MinValue        =   1
      MaxValue        =   4
   End
   Begin Progetto1.H_UpDown H_UpDown4 
      Height          =   255
      Left            =   6960
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777152
      ForeColor       =   16659743
      MaxValue        =   6
      LoopValue       =   -1  'True
   End
   Begin Progetto1.H_UpDown H_UpDown3 
      Height          =   255
      Left            =   6960
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   16777152
      ForeColor       =   16659743
      MaxValue        =   9
      LoopValue       =   -1  'True
   End
   Begin Progetto1.DisplayLed DisplayLed1 
      Height          =   660
      Index           =   0
      Left            =   3840
      TabIndex        =   3
      Top             =   3080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      Zoom            =   2
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.H_UpDown H_UpDown1 
      Height          =   255
      Left            =   8910
      TabIndex        =   1
      Top             =   675
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   450
      MaxValue        =   10
      LoopValue       =   -1  'True
   End
   Begin Progetto1.DisplayLed DisplayLed1 
      Height          =   660
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   3080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      Zoom            =   2
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed DisplayLed1 
      Height          =   660
      Index           =   2
      Left            =   4800
      TabIndex        =   5
      Top             =   3080
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      Zoom            =   2
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   1
      Left            =   720
      TabIndex        =   15
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   16
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   3
      Left            =   1200
      TabIndex        =   17
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   4
      Left            =   1440
      TabIndex        =   18
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   5
      Left            =   1680
      TabIndex        =   19
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   6
      Left            =   1920
      TabIndex        =   20
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   7
      Left            =   2160
      TabIndex        =   21
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   8
      Left            =   2400
      TabIndex        =   22
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   9
      Left            =   2640
      TabIndex        =   23
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   10
      Left            =   2880
      TabIndex        =   24
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   11
      Left            =   3120
      TabIndex        =   25
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   12
      Left            =   3360
      TabIndex        =   26
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   13
      Left            =   3600
      TabIndex        =   27
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   14
      Left            =   3840
      TabIndex        =   28
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   15
      Left            =   4080
      TabIndex        =   29
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   16
      Left            =   4320
      TabIndex        =   30
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   17
      Left            =   4560
      TabIndex        =   31
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   18
      Left            =   4800
      TabIndex        =   32
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   19
      Left            =   5040
      TabIndex        =   33
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   20
      Left            =   480
      TabIndex        =   34
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   21
      Left            =   720
      TabIndex        =   35
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   22
      Left            =   960
      TabIndex        =   38
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   23
      Left            =   1200
      TabIndex        =   39
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   24
      Left            =   1440
      TabIndex        =   40
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   25
      Left            =   1680
      TabIndex        =   41
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   26
      Left            =   1920
      TabIndex        =   42
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   27
      Left            =   2160
      TabIndex        =   43
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   28
      Left            =   2400
      TabIndex        =   44
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   29
      Left            =   2640
      TabIndex        =   45
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   30
      Left            =   2880
      TabIndex        =   46
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   31
      Left            =   3120
      TabIndex        =   47
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   32
      Left            =   3360
      TabIndex        =   48
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   33
      Left            =   3600
      TabIndex        =   49
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   34
      Left            =   3840
      TabIndex        =   50
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   35
      Left            =   4080
      TabIndex        =   51
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   36
      Left            =   4320
      TabIndex        =   52
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   37
      Left            =   4560
      TabIndex        =   53
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   38
      Left            =   4800
      TabIndex        =   54
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   39
      Left            =   5040
      TabIndex        =   55
      Top             =   2160
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   1
      Left            =   720
      TabIndex        =   62
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   2
      Left            =   960
      TabIndex        =   63
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   3
      Left            =   1200
      TabIndex        =   64
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   4
      Left            =   1440
      TabIndex        =   65
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   5
      Left            =   1680
      TabIndex        =   66
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   6
      Left            =   1920
      TabIndex        =   67
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   7
      Left            =   2160
      TabIndex        =   68
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   8
      Left            =   2400
      TabIndex        =   69
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   9
      Left            =   2640
      TabIndex        =   70
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   10
      Left            =   2880
      TabIndex        =   71
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   11
      Left            =   3120
      TabIndex        =   72
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   12
      Left            =   3360
      TabIndex        =   73
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   13
      Left            =   3600
      TabIndex        =   74
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   14
      Left            =   3840
      TabIndex        =   75
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   15
      Left            =   4080
      TabIndex        =   76
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   16
      Left            =   4320
      TabIndex        =   77
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   17
      Left            =   4560
      TabIndex        =   78
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   18
      Left            =   4800
      TabIndex        =   79
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.DisplayLed Led_User 
      Height          =   330
      Index           =   19
      Left            =   5040
      TabIndex        =   80
      Top             =   2640
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      Colore          =   16551548
      Serie           =   1
   End
   Begin Progetto1.V_UpDown V_UpDown1 
      Height          =   855
      Index           =   1
      Left            =   4335
      TabIndex        =   106
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1508
      BackColor       =   16659743
      ForeColor       =   16777152
      MaxValue        =   9
      LoopValue       =   -1  'True
   End
   Begin Progetto1.V_UpDown V_UpDown1 
      Height          =   855
      Index           =   2
      Left            =   4815
      TabIndex        =   107
      Top             =   4080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1508
      BackColor       =   16659743
      ForeColor       =   16777152
      MaxValue        =   9
      LoopValue       =   -1  'True
   End
   Begin Progetto1.V_UpDown V_UpDown2 
      Height          =   855
      Left            =   6600
      TabIndex        =   108
      Top             =   5040
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1508
      Value           =   1
      MaxValue        =   2
      LoopValue       =   -1  'True
   End
   Begin Progetto1.VSlider VSlider1 
      Height          =   2655
      Index           =   1
      Left            =   7440
      TabIndex        =   136
      ToolTipText     =   "DownToUp = True"
      Top             =   4080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4683
      Value           =   5
      Picture         =   "Demo.frx":11AE76
      PictureCursor   =   "Demo.frx":11C51A
   End
   Begin Progetto1.Led LedBlink 
      Height          =   255
      Index           =   1
      Left            =   5000
      TabIndex        =   139
      Top             =   620
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Colore          =   2
   End
   Begin VB.Image ImgRduci 
      Height          =   150
      Left            =   9960
      Picture         =   "Demo.frx":11C81A
      Top             =   165
      Width           =   180
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Binary"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      TabIndex        =   104
      Top             =   640
      Width           =   1215
   End
   Begin VB.Image ImgFrame 
      Height          =   1095
      Left            =   10275
      Picture         =   "Demo.frx":11C9C6
      Top             =   2235
      Width           =   150
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER TEXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   85
      Top             =   640
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Speed -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FC8E7C&
      Height          =   255
      Left            =   480
      TabIndex        =   83
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H Slider"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   82
      Top             =   4380
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8610
      TabIndex        =   60
      Top             =   1860
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V Sliders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   58
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Speed -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FC8E7C&
      Height          =   255
      Left            =   480
      TabIndex        =   57
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BackGrd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   56
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image ImgExit 
      Height          =   180
      Left            =   10200
      Picture         =   "Demo.frx":11D32A
      Top             =   160
      Width           =   180
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   37
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Image ImgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   7920
      Picture         =   "Demo.frx":11D51E
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2280
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Colore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UpDown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   2
      Top             =   680
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   680
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   5760
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   8490
      Shape           =   4  'Rounded Rectangle
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   4380
      Width           =   1575
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
' Descrizione.....: Demo OCX
' Nome dei Files..: VSlider - HSlider - H_Updown - V_Updown
'                 : DisplayLed - Led
' Data............: 02/01/2005
' Versione........: 1.05
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'===========================================================
'
'                    Not For Commercial Use
'===========================================================
Option Explicit

Private I As Integer
Private I1 As Integer
Private CntDs As Integer
Private StrDisplay As String
'
Private CntDs1 As Integer
'
Private MouseX As Long
Private MouseY As Long
Private OldMouseX As Long
Private OldMouseY As Long
Private MouseDown As Boolean
'
Private CursoreH(2) As Picture
Private CursoreHA(2) As Picture
Private CursoreV(2) As Picture
Private CursFile As String

'
'
'
Private Sub Form_Load()
   
   ProgressBar1.Value = HSlider1.Value  ' ProgressBar1
   ProgressBar2.Value = HSlider2.Value  ' ProgressBar2
   
   
   Txt_VSlider1(0) = VSlider1(0).Value  ' Vslider1(0)
   Txt_VSlider1(1) = VSlider1(1).Value  ' Vslider1(1)
   
   TxtUpDown = H_UpDown1.Value          ' H_UpDown
   
   DisplayLed1(0).Value = Asc("0")      ' Display 1
   DisplayLed1(1).Value = Asc("0")      ' Display 1
   DisplayLed1(2).Value = Asc("0")      ' Display 1
   
   DisplayLed2.Value = Asc("0")         ' Display 2
   '
   StrDisplay = "BY BRUNO CREPALDI - 2004 - "
   StrDisplay = StrDisplay + "HELLO FROM VENICE - ITALY - "
   '
   TimerDisplay.Interval = 110 - HSlider1.Value
   TimerDisplay_User.Interval = 110 - HSlider2.Value
   Timer_Led.Interval = 110 - HSlider3.Value
   '
   For I = 0 To 2                                               '----------------------------------
     CursFile = App.Path + "\Cursore H" + Trim(Str(I)) + " S.BMP" '
     Set CursoreH(I) = LoadPicture(CursFile)                    ' Load Horizontals Cursors Pictures
     CursFile = App.Path + "\Cursore H" + Trim(Str(I)) + " A.BMP" '
     Set CursoreHA(I) = LoadPicture(CursFile)
   Next I                                                       '----------------------------------
   For I = 0 To 2                                               '----------------------------------
     CursFile = App.Path + "\Cursore V" + Trim(Str(I)) + ".BMP" '
     Set CursoreV(I) = LoadPicture(CursFile)                    ' Load Verticals Cursors Pictures
   Next I                                                       '----------------------------------
End Sub

Private Sub HSlider3_Change(Value As Long)
   Timer_Led.Interval = 110 - HSlider3.Value
End Sub

'
Private Sub ImgExit_Click()          ' End
  Unload Me
End Sub
Private Sub ImgRduci_Click()         ' Reduce
  Me.WindowState = 1
End Sub
'-------------------------------------------------------
'                 Horizontal Sliders
'-------------------------------------------------------
Private Sub HSlider1_Change(Value As Long)
  ProgressBar1.Value = Value
  TimerDisplay.Interval = 110 - HSlider1.Value
End Sub
'
Private Sub HSlider2_Change(Value As Long)
  ProgressBar2.Value = Value
  TimerDisplay_User.Interval = 110 - HSlider2.Value

End Sub

Private Sub H_Updown2_Change(Value As Long)
 Set HSlider1.PictureCursor = CursoreH(Value)
 Set HSlider1.PicCursor_Selected = CursoreHA(Value)
 Set HSlider2.PictureCursor = CursoreH(Value)
 Set HSlider2.PicCursor_Selected = CursoreHA(Value)
End Sub

'--------------------------------------------------------
'                  Vertical Sliders
'--------------------------------------------------------
Private Sub VSlider1_Change(Index As Integer, Value As Long)
   Txt_VSlider1(Index).Text = Value
End Sub

Private Sub Txt_VSlider1_Change(Index As Integer)
   VSlider1(Index).Value = Val(Txt_VSlider1(Index).Text)
   Call BarreLed(Index, VSlider1(Index).Value)
End Sub

Private Sub V_UpDown2_Change(Value As Long)
   Set VSlider1(0).PictureCursor = CursoreV(Value)
   Set VSlider1(1).PictureCursor = CursoreV(Value)
End Sub
'
Private Sub BarreLed(Numbarra As Integer, Value As Long)
 Select Case Numbarra
 Case 0
   For I = 10 To 0 Step -1
     If I > Value Then
       Led_Barra_L(I).Status = False
     Else
       Led_Barra_L(I).Status = True
     End If
   Next I
 Case 1
   For I = 10 To 0 Step -1
     If I > Value Then
       Led_Barra_R(I).Status = False
     Else
       Led_Barra_R(I).Status = True
     End If
  Next I
 End Select
End Sub
'--------------------------------------------------------
'                    DisplayLed Esempio 1
'--------------------------------------------------------
Private Sub V_UpDown1_Change(Index As Integer, Value As Long)
  DisplayLed1(Index).Value = Asc(Trim(Str(Value))) ' Ascii Value
End Sub

'--------------------------------------------------------
'                    DisplayLed Esempio 2
'--------------------------------------------------------
Private Sub H_Updown3_Change(Value As Long)
 DisplayLed2.Value = Asc(Trim(Str(Value))) '    Valore Ascii
End Sub
Private Sub H_Updown4_Change(Value As Long)
 Select Case Value
   Case 0
     DisplayLed2.Colore = RGB(168, 255, 0)
   Case 1
     DisplayLed2.Colore = &HFFFFFF
   Case 2
     DisplayLed2.Colore = RGB(255, 90, 0)
   Case 3
     DisplayLed2.Colore = RGB(252, 255, 0)
   Case 4
     DisplayLed2.Colore = RGB(168, 250, 255)
   Case 5
     DisplayLed2.Colore = RGB(255, 150, 200)
   Case 6
     DisplayLed2.Colore = RGB(124, 142, 252)
 End Select
End Sub
Private Sub H_Updown5_Change(Value As Long)
 DisplayLed2.Zoom = Value
End Sub
Private Sub H_Updown6_Change(Value As Long)
 DisplayLed2.Style = Value
End Sub
Private Sub H_Updown7_Change(Value As Long)
 DisplayLed2.Serie = Value
End Sub
'--------------------------------------------------------
'                  Binary Counter
'--------------------------------------------------------

Private Sub H_Updown9_Change(Value As Long)
 Dim Bt As Integer
 '
  Bt = 1                          ' Peso del Bit da controllare
 For I = 0 To 7                   ' da bit 0 a bit 7
   If (Value Or Bt) = Value Then  ' Controla se Bit(n) e ON
     Led1(I).Status = True        ' ON
    Else                          '
     Led1(I).Status = False       ' OFF
   End If                         '
  Bt = Bt * 2                     ' peso del Prossomo Bit da Controllare
 Next I                           '
End Sub
Private Sub H_Updown10_Change(Value As Long)
 For I = 0 To 7
   Led1(I).Colore = Value
 Next I
End Sub
'--------------------------------------------------------
'                     Timer Led  ( Binary )
'--------------------------------------------------------

Private Sub Timer_Led_Timer()
   If CheckBinary.Value = 1 Then
    If H_UpDown9.Value = 255 Then H_UpDown9.Value = -1
    H_UpDown9.Value = H_UpDown9.Value + 1
  End If

   LedBlink(0).Status = Not LedBlink(0).Status
   LedBlink(1).Status = Not LedBlink(0).Status
End Sub
'--------------------------------------------------------
'         Example H_Updown V_Updown
'--------------------------------------------------------
Private Sub H_Updown1_Change(Value As Long)
  TxtUpDown = Value
End Sub
Private Sub V_UpDown3_Change(Value As Long)
  TxtUpDown = Value
End Sub
Private Sub TxtUpDown_Change()
  V_UpDown3.Value = Val(TxtUpDown.Text)
  H_UpDown1.Value = Val(TxtUpDown.Text)
End Sub
'--------------------------------------------------------
'                     Riga Display
'--------------------------------------------------------
Private Sub TimerDisplay_Timer()
  Dim Ch As String * 1
  Dim Lstr As String
     
   Lstr = Len(StrDisplay)
   
   CntDs = CntDs + 1: If CntDs > Lstr Then CntDs = 1
     
     For I1 = 0 To 38  ' 20
       LedRiga(I1).Value = LedRiga(I1 + 1).Value
     Next I1
    
    Ch = Mid$(StrDisplay, CntDs, 1)
    LedRiga(39).Value = Asc(Ch)
   
End Sub

'--------------------------------------------------------
'                     Riga Display USER
'--------------------------------------------------------
Private Sub TimerDisplay_User_Timer()
  Dim Ch As String * 1
  Dim Lstr As String
     
   Lstr = Len(Txt_User.Text)
   
   CntDs1 = CntDs1 + 1: If CntDs1 > Lstr Then CntDs1 = 1
     
     For I1 = 0 To 18
       Led_User(I1).Value = Led_User(I1 + 1).Value
     Next I1
    
    Ch = Mid$(UCase(Txt_User.Text), CntDs1, 1)
    Led_User(19).Value = Asc(Ch)
  
End Sub
'
'
'                 Gestione Spostamento del Form ( Form Move )
'
'
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MouseDown = False
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MouseDown = True
     Call GetCursorPos(MouseX, MouseY)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call SpostaForm
End Sub

Private Sub SpostaForm()
   Dim DifX As Long
   Dim DifY As Long
   
     If MouseDown = False Then Exit Sub ' Tasto non Premuto esci
     
     OldMouseX = Me.ScaleX(MouseX, vbPixels, vbTwips)
     OldMouseY = Me.ScaleY(MouseY, vbPixels, vbTwips)
   
     Call GetCursorPos(MouseX, MouseY)
       
     DifX = OldMouseX - Me.ScaleX(MouseX, vbPixels, vbTwips)
     DifY = OldMouseY - Me.ScaleY(MouseY, vbPixels, vbTwips)

     Me.Left = Me.Left - DifX
     Me.Top = Me.Top - DifY

End Sub


