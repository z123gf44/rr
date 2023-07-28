VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DYNAVOLT TECH."
   ClientHeight    =   12015
   ClientLeft      =   1275
   ClientTop       =   1710
   ClientWidth     =   21915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   12117.82
   ScaleMode       =   0  'User
   ScaleWidth      =   21969.92
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   390
      Index           =   0
      Left            =   18240
      TabIndex        =   339
      Text            =   "Text2"
      Top             =   11640
      Width           =   3015
   End
   Begin VB.Frame ZFrame1 
      Caption         =   "电池状态"
      ForeColor       =   &H8000000D&
      Height          =   855
      Index           =   1
      Left            =   5400
      TabIndex        =   329
      Top             =   10800
      Width           =   15855
   End
   Begin VB.Frame ZFrame1 
      Caption         =   "系统状态"
      Height          =   855
      Index           =   0
      Left            =   5400
      TabIndex        =   327
      Top             =   9960
      Width           =   15855
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H000000FF&
      Caption         =   "重连"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   240
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   8340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCalibration 
      Caption         =   "校准"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "备份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdMcuSysConfig 
      Caption         =   "MCU配置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdAfeReg 
      Caption         =   "AFE寄存器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdPackInfo 
      Caption         =   "PACK信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出程序"
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   9960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCfgSPort 
      Caption         =   "串口配置"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   600
      Top             =   8760
   End
   Begin VB.Frame zFrame7 
      Caption         =   "Frame7"
      Height          =   10095
      Left            =   1560
      TabIndex        =   10
      Top             =   -120
      Width           =   19740
      Begin VB.Frame ZFrame1 
         Height          =   10095
         Index           =   10
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   9615
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   100
            Left            =   6000
            TabIndex        =   76
            Text            =   "√"
            Top             =   10440
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   99
            Left            =   7320
            TabIndex        =   75
            Text            =   "√"
            Top             =   10440
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   98
            Left            =   8640
            TabIndex        =   74
            Text            =   "√"
            Top             =   10440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   97
            Left            =   6000
            TabIndex        =   73
            Text            =   "√"
            Top             =   9960
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   96
            Left            =   7320
            TabIndex        =   72
            Text            =   "√"
            Top             =   9960
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   95
            Left            =   8640
            TabIndex        =   71
            Text            =   "√"
            Top             =   9960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   94
            Left            =   6000
            TabIndex        =   70
            Text            =   "√"
            Top             =   9480
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   93
            Left            =   7320
            TabIndex        =   69
            Text            =   "√"
            Top             =   9480
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   92
            Left            =   8640
            TabIndex        =   68
            Text            =   "√"
            Top             =   9480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   91
            Left            =   6000
            TabIndex        =   67
            Text            =   "√"
            Top             =   9000
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   90
            Left            =   7320
            TabIndex        =   66
            Text            =   "√"
            Top             =   9000
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   53
            Left            =   8640
            TabIndex        =   65
            Text            =   "√"
            Top             =   9000
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   52
            Left            =   6000
            TabIndex        =   64
            Text            =   "√"
            Top             =   8520
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   15
            Left            =   7320
            TabIndex        =   63
            Text            =   "√"
            Top             =   8520
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   13
            Left            =   8640
            TabIndex        =   62
            Text            =   "√"
            Top             =   8520
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   12
            Left            =   6000
            TabIndex        =   61
            Text            =   "√"
            Top             =   8040
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   11
            Left            =   7320
            TabIndex        =   60
            Text            =   "√"
            Top             =   8040
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   10
            Left            =   8640
            TabIndex        =   59
            Text            =   "√"
            Top             =   8040
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   31
            Left            =   8640
            TabIndex        =   58
            Text            =   "√"
            Top             =   6120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   34
            Left            =   8640
            TabIndex        =   57
            Text            =   "√"
            Top             =   5640
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   37
            Left            =   8640
            TabIndex        =   56
            Text            =   "√"
            Top             =   5160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   40
            Left            =   8640
            TabIndex        =   55
            Text            =   "√"
            Top             =   4680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   43
            Left            =   8640
            TabIndex        =   54
            Text            =   "√"
            Top             =   4200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   46
            Left            =   8640
            TabIndex        =   53
            Text            =   "√"
            Top             =   3720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   49
            Left            =   8640
            TabIndex        =   52
            Text            =   "√"
            Top             =   3240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   50
            Left            =   8640
            TabIndex        =   51
            Text            =   "√"
            Top             =   2760
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   54
            Left            =   8640
            TabIndex        =   50
            Text            =   "√"
            Top             =   7560
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   55
            Left            =   8640
            TabIndex        =   49
            Text            =   "√"
            Top             =   7080
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   56
            Left            =   8640
            TabIndex        =   48
            Text            =   "√"
            Top             =   6600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   57
            Left            =   8640
            TabIndex        =   47
            Text            =   "√"
            Top             =   2280
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   58
            Left            =   8640
            TabIndex        =   46
            Text            =   "√"
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   59
            Left            =   8640
            TabIndex        =   45
            Text            =   "√"
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   60
            Left            =   7320
            TabIndex        =   44
            Text            =   "√"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   61
            Left            =   7320
            TabIndex        =   43
            Text            =   "√"
            Top             =   7560
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   62
            Left            =   6000
            TabIndex        =   42
            Text            =   "√"
            Top             =   7560
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   63
            Left            =   7320
            TabIndex        =   41
            Text            =   "√"
            Top             =   7080
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   64
            Left            =   6000
            TabIndex        =   40
            Text            =   "√"
            Top             =   7080
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   65
            Left            =   7320
            TabIndex        =   39
            Text            =   "√"
            Top             =   6600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   66
            Left            =   6000
            TabIndex        =   38
            Text            =   "√"
            Top             =   6600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   67
            Left            =   7320
            TabIndex        =   37
            Text            =   "√"
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   68
            Left            =   6000
            TabIndex        =   36
            Text            =   "√"
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   69
            Left            =   7320
            TabIndex        =   35
            Text            =   "√"
            Top             =   5640
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   70
            Left            =   6000
            TabIndex        =   34
            Text            =   "√"
            Top             =   5640
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   71
            Left            =   7320
            TabIndex        =   33
            Text            =   "√"
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   72
            Left            =   6000
            TabIndex        =   32
            Text            =   "√"
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   73
            Left            =   7320
            TabIndex        =   31
            Text            =   "√"
            Top             =   4680
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   74
            Left            =   6000
            TabIndex        =   30
            Text            =   "√"
            Top             =   4680
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   75
            Left            =   7320
            TabIndex        =   29
            Text            =   "√"
            Top             =   4200
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   76
            Left            =   6000
            TabIndex        =   28
            Text            =   "√"
            Top             =   4200
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   77
            Left            =   7320
            TabIndex        =   27
            Text            =   "√"
            Top             =   3720
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   78
            Left            =   6000
            TabIndex        =   26
            Text            =   "√"
            Top             =   3720
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   79
            Left            =   7320
            TabIndex        =   25
            Text            =   "√"
            Top             =   3240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   80
            Left            =   6000
            TabIndex        =   24
            Text            =   "√"
            Top             =   3240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   81
            Left            =   7320
            TabIndex        =   23
            Text            =   "√"
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   82
            Left            =   6000
            TabIndex        =   22
            Text            =   "√"
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   83
            Left            =   7320
            TabIndex        =   21
            Text            =   "√"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   84
            Left            =   6000
            TabIndex        =   20
            Text            =   "√"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   85
            Left            =   7320
            TabIndex        =   19
            Text            =   "√"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   86
            Left            =   6000
            TabIndex        =   18
            Text            =   "√"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   87
            Left            =   7320
            TabIndex        =   17
            Text            =   "√"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   420
            Index           =   88
            Left            =   6000
            TabIndex        =   16
            Text            =   "√"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   89
            Left            =   6000
            TabIndex        =   15
            Text            =   "√"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Index           =   51
            Left            =   8640
            TabIndex        =   14
            Text            =   "√"
            Top             =   1800
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   199
            Left            =   5640
            TabIndex        =   187
            Top             =   10440
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   198
            Left            =   6960
            TabIndex        =   186
            Top             =   10440
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯21电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   197
            Left            =   120
            TabIndex        =   185
            Top             =   10440
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   196
            Left            =   2760
            TabIndex        =   184
            Top             =   10440
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   195
            Left            =   4320
            TabIndex        =   183
            Top             =   10440
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   194
            Left            =   5640
            TabIndex        =   182
            Top             =   9960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   193
            Left            =   6960
            TabIndex        =   181
            Top             =   9960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯20电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   192
            Left            =   120
            TabIndex        =   180
            Top             =   9960
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   191
            Left            =   2760
            TabIndex        =   179
            Top             =   9960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   190
            Left            =   4320
            TabIndex        =   178
            Top             =   9960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   189
            Left            =   5640
            TabIndex        =   177
            Top             =   9480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   188
            Left            =   6960
            TabIndex        =   176
            Top             =   9480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯19电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   187
            Left            =   120
            TabIndex        =   175
            Top             =   9480
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   186
            Left            =   2760
            TabIndex        =   174
            Top             =   9480
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   185
            Left            =   4320
            TabIndex        =   173
            Top             =   9480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   184
            Left            =   5640
            TabIndex        =   172
            Top             =   9000
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   183
            Left            =   6960
            TabIndex        =   171
            Top             =   9000
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯18电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   182
            Left            =   120
            TabIndex        =   170
            Top             =   9000
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   181
            Left            =   2760
            TabIndex        =   169
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   180
            Left            =   4320
            TabIndex        =   168
            Top             =   9000
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   179
            Left            =   5640
            TabIndex        =   167
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   178
            Left            =   6960
            TabIndex        =   166
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯17电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   177
            Left            =   120
            TabIndex        =   165
            Top             =   8520
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   176
            Left            =   2760
            TabIndex        =   164
            Top             =   8520
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   175
            Left            =   4320
            TabIndex        =   163
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   174
            Left            =   5640
            TabIndex        =   162
            Top             =   8040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   173
            Left            =   6960
            TabIndex        =   161
            Top             =   8040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯16电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   172
            Left            =   120
            TabIndex        =   160
            Top             =   8040
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   171
            Left            =   2760
            TabIndex        =   159
            Top             =   8040
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   170
            Left            =   4320
            TabIndex        =   158
            Top             =   8040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   95
            Left            =   4320
            TabIndex        =   157
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   96
            Left            =   2760
            TabIndex        =   156
            Top             =   7560
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯15电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   97
            Left            =   120
            TabIndex        =   155
            Top             =   7560
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   98
            Left            =   4320
            TabIndex        =   154
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   99
            Left            =   2760
            TabIndex        =   153
            Top             =   7080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯14电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   100
            Left            =   120
            TabIndex        =   152
            Top             =   7080
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   101
            Left            =   4320
            TabIndex        =   151
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   102
            Left            =   2760
            TabIndex        =   150
            Top             =   6600
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯13电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   103
            Left            =   120
            TabIndex        =   149
            Top             =   6600
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   104
            Left            =   4320
            TabIndex        =   148
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   105
            Left            =   2760
            TabIndex        =   147
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯12电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   106
            Left            =   120
            TabIndex        =   146
            Top             =   6120
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   107
            Left            =   4320
            TabIndex        =   145
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   108
            Left            =   2760
            TabIndex        =   144
            Top             =   5640
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯11电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   109
            Left            =   120
            TabIndex        =   143
            Top             =   5640
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   110
            Left            =   4320
            TabIndex        =   142
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   111
            Left            =   2760
            TabIndex        =   141
            Top             =   5160
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯10电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   112
            Left            =   120
            TabIndex        =   140
            Top             =   5160
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   113
            Left            =   4320
            TabIndex        =   139
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   114
            Left            =   2760
            TabIndex        =   138
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯09电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   115
            Left            =   120
            TabIndex        =   137
            Top             =   4680
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   116
            Left            =   4320
            TabIndex        =   136
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   117
            Left            =   2760
            TabIndex        =   135
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯08电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   118
            Left            =   120
            TabIndex        =   134
            Top             =   4200
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   119
            Left            =   4320
            TabIndex        =   133
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   120
            Left            =   2760
            TabIndex        =   132
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯07电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   121
            Left            =   120
            TabIndex        =   131
            Top             =   3720
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   122
            Left            =   4320
            TabIndex        =   130
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   505
            Left            =   2760
            TabIndex        =   129
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯06电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   124
            Left            =   120
            TabIndex        =   128
            Top             =   3240
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   125
            Left            =   4320
            TabIndex        =   127
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   504
            Left            =   2760
            TabIndex        =   126
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯05电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   127
            Left            =   120
            TabIndex        =   125
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   128
            Left            =   4320
            TabIndex        =   124
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   503
            Left            =   2760
            TabIndex        =   123
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯04电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   130
            Left            =   120
            TabIndex        =   122
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   131
            Left            =   4320
            TabIndex        =   121
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   502
            Left            =   2760
            TabIndex        =   120
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯03电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   133
            Left            =   120
            TabIndex        =   119
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   134
            Left            =   4320
            TabIndex        =   118
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   501
            Left            =   2760
            TabIndex        =   117
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯02电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   136
            Left            =   120
            TabIndex        =   116
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   137
            Left            =   4320
            TabIndex        =   115
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   500
            Left            =   2760
            TabIndex        =   114
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯01电压"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   139
            Left            =   120
            TabIndex        =   113
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   140
            Left            =   6960
            TabIndex        =   112
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   141
            Left            =   5640
            TabIndex        =   111
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   142
            Left            =   6960
            TabIndex        =   110
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   143
            Left            =   5640
            TabIndex        =   109
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   144
            Left            =   6960
            TabIndex        =   108
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   145
            Left            =   5640
            TabIndex        =   107
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   146
            Left            =   6960
            TabIndex        =   106
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   147
            Left            =   5640
            TabIndex        =   105
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   148
            Left            =   6960
            TabIndex        =   104
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   149
            Left            =   5640
            TabIndex        =   103
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   150
            Left            =   6960
            TabIndex        =   102
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   151
            Left            =   5640
            TabIndex        =   101
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   152
            Left            =   6960
            TabIndex        =   100
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   153
            Left            =   5640
            TabIndex        =   99
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   154
            Left            =   6960
            TabIndex        =   98
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   155
            Left            =   5640
            TabIndex        =   97
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   156
            Left            =   6960
            TabIndex        =   96
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   157
            Left            =   5640
            TabIndex        =   95
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   158
            Left            =   6960
            TabIndex        =   94
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   159
            Left            =   5640
            TabIndex        =   93
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   160
            Left            =   6960
            TabIndex        =   92
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   161
            Left            =   5640
            TabIndex        =   91
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   162
            Left            =   6960
            TabIndex        =   90
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   163
            Left            =   5640
            TabIndex        =   89
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   164
            Left            =   6960
            TabIndex        =   88
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   165
            Left            =   5640
            TabIndex        =   87
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   166
            Left            =   6960
            TabIndex        =   86
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   167
            Left            =   5640
            TabIndex        =   85
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   168
            Left            =   6960
            TabIndex        =   84
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   169
            Left            =   5640
            TabIndex        =   83
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "均衡"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   8280
            TabIndex        =   82
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "√扫描"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   6960
            TabIndex        =   81
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "√记录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   5640
            TabIndex        =   80
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "单位"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   4320
            TabIndex        =   79
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "数值"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   2760
            TabIndex        =   78
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "名称"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame ZFrame1 
         Height          =   10095
         Index           =   11
         Left            =   10200
         TabIndex        =   188
         Top             =   0
         Width           =   8415
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   28
            Left            =   6000
            TabIndex        =   226
            Text            =   "√"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   25
            Left            =   7320
            TabIndex        =   225
            Text            =   "√"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   22
            Left            =   6000
            TabIndex        =   224
            Text            =   "√"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   19
            Left            =   7320
            TabIndex        =   223
            Text            =   "√"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   16
            Left            =   6000
            TabIndex        =   222
            Text            =   "√"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   7
            Left            =   7320
            TabIndex        =   221
            Text            =   "√"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   9
            Left            =   6000
            TabIndex        =   220
            Text            =   "√"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   14
            Left            =   7320
            TabIndex        =   219
            Text            =   "√"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   17
            Left            =   6000
            TabIndex        =   218
            Text            =   "√"
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   18
            Left            =   7320
            TabIndex        =   217
            Text            =   "√"
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   20
            Left            =   6000
            TabIndex        =   216
            Text            =   "√"
            Top             =   3240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   21
            Left            =   7320
            TabIndex        =   215
            Text            =   "√"
            Top             =   3240
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   23
            Left            =   6000
            TabIndex        =   214
            Text            =   "√"
            Top             =   3720
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   24
            Left            =   7320
            TabIndex        =   213
            Text            =   "√"
            Top             =   3720
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   26
            Left            =   6000
            TabIndex        =   212
            Text            =   "√"
            Top             =   4200
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   27
            Left            =   7320
            TabIndex        =   211
            Text            =   "√"
            Top             =   4200
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   29
            Left            =   6000
            TabIndex        =   210
            Text            =   "√"
            Top             =   4680
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   30
            Left            =   7320
            TabIndex        =   209
            Text            =   "√"
            Top             =   4680
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   32
            Left            =   6000
            TabIndex        =   208
            Text            =   "√"
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   33
            Left            =   7320
            TabIndex        =   207
            Text            =   "√"
            Top             =   5160
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   35
            Left            =   6000
            TabIndex        =   206
            Text            =   "√"
            Top             =   5640
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   36
            Left            =   7320
            TabIndex        =   205
            Text            =   "√"
            Top             =   5640
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   38
            Left            =   6000
            TabIndex        =   204
            Text            =   "√"
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   39
            Left            =   7320
            TabIndex        =   203
            Text            =   "√"
            Top             =   6120
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   41
            Left            =   6000
            TabIndex        =   202
            Text            =   "√"
            Top             =   6600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   42
            Left            =   7320
            TabIndex        =   201
            Text            =   "√"
            Top             =   6600
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   44
            Left            =   6000
            TabIndex        =   200
            Text            =   "√"
            Top             =   7080
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   45
            Left            =   7320
            TabIndex        =   199
            Text            =   "√"
            Top             =   7080
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   47
            Left            =   6000
            TabIndex        =   198
            Text            =   "√"
            Top             =   7560
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   48
            Left            =   7320
            TabIndex        =   197
            Text            =   "√"
            Top             =   7560
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   6
            Left            =   6000
            TabIndex        =   196
            Text            =   "√"
            Top             =   8040
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   5
            Left            =   7320
            TabIndex        =   195
            Text            =   "√"
            Top             =   8040
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   8
            Left            =   6000
            TabIndex        =   194
            Text            =   "√"
            Top             =   8520
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   4
            Left            =   7320
            TabIndex        =   193
            Text            =   "√"
            Top             =   8520
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   3
            Left            =   6000
            TabIndex        =   192
            Text            =   "√"
            Top             =   9000
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   2
            Left            =   7320
            TabIndex        =   191
            Text            =   "√"
            Top             =   9000
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   1
            Left            =   6000
            TabIndex        =   190
            Text            =   "√"
            Top             =   9480
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   0
            Left            =   7320
            TabIndex        =   189
            Text            =   "√"
            Top             =   9480
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "名称"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   326
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "数值"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   2760
            TabIndex        =   325
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "单位"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   4320
            TabIndex        =   324
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "√记录"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5640
            TabIndex        =   323
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            Caption         =   "√扫描"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   6960
            TabIndex        =   322
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   0
            Left            =   5640
            TabIndex        =   321
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   1
            Left            =   6960
            TabIndex        =   320
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "实时电流值"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   319
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   3
            Left            =   2760
            TabIndex        =   318
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mA"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   4
            Left            =   4320
            TabIndex        =   317
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   5
            Left            =   5640
            TabIndex        =   316
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   6
            Left            =   6960
            TabIndex        =   315
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "环境温度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   314
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   8
            Left            =   2760
            TabIndex        =   313
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   9
            Left            =   4320
            TabIndex        =   312
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   311
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   11
            Left            =   6960
            TabIndex        =   310
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   309
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   13
            Left            =   2760
            TabIndex        =   308
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   14
            Left            =   4320
            TabIndex        =   307
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   15
            Left            =   5640
            TabIndex        =   306
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   16
            Left            =   6960
            TabIndex        =   305
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度2"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   304
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   18
            Left            =   2760
            TabIndex        =   303
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   19
            Left            =   4320
            TabIndex        =   302
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   20
            Left            =   5640
            TabIndex        =   301
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   21
            Left            =   6960
            TabIndex        =   300
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度3"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   22
            Left            =   120
            TabIndex        =   299
            Top             =   2760
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   23
            Left            =   2760
            TabIndex        =   298
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   24
            Left            =   4320
            TabIndex        =   297
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   25
            Left            =   5640
            TabIndex        =   296
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   26
            Left            =   6960
            TabIndex        =   295
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   27
            Left            =   120
            TabIndex        =   294
            Top             =   3240
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   28
            Left            =   2760
            TabIndex        =   293
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   29
            Left            =   4320
            TabIndex        =   292
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   30
            Left            =   5640
            TabIndex        =   291
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   31
            Left            =   6960
            TabIndex        =   290
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度5"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   32
            Left            =   120
            TabIndex        =   289
            Top             =   3720
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   33
            Left            =   2760
            TabIndex        =   288
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   34
            Left            =   4320
            TabIndex        =   287
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   35
            Left            =   5640
            TabIndex        =   286
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   36
            Left            =   6960
            TabIndex        =   285
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度6"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   37
            Left            =   120
            TabIndex        =   284
            Top             =   4200
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   38
            Left            =   2760
            TabIndex        =   283
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   39
            Left            =   4320
            TabIndex        =   282
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   40
            Left            =   5640
            TabIndex        =   281
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   41
            Left            =   6960
            TabIndex        =   280
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "电芯温度7"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   42
            Left            =   120
            TabIndex        =   279
            Top             =   4680
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   43
            Left            =   2760
            TabIndex        =   278
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   44
            Left            =   4320
            TabIndex        =   277
            Top             =   4680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   45
            Left            =   5640
            TabIndex        =   276
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   46
            Left            =   6960
            TabIndex        =   275
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "功率温度"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   47
            Left            =   120
            TabIndex        =   274
            Top             =   5160
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   48
            Left            =   2760
            TabIndex        =   273
            Top             =   5160
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "℃"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   49
            Left            =   4320
            TabIndex        =   272
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   50
            Left            =   5640
            TabIndex        =   271
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   51
            Left            =   6960
            TabIndex        =   270
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "剩余电量百分比"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   52
            Left            =   120
            TabIndex        =   269
            Top             =   5640
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   53
            Left            =   2760
            TabIndex        =   268
            Top             =   5640
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   54
            Left            =   4320
            TabIndex        =   267
            Top             =   5640
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   55
            Left            =   5640
            TabIndex        =   266
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   56
            Left            =   6960
            TabIndex        =   265
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "系统满充容量"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   57
            Left            =   120
            TabIndex        =   264
            Top             =   6120
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   58
            Left            =   2760
            TabIndex        =   263
            Top             =   6120
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mAH"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   59
            Left            =   4320
            TabIndex        =   262
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   63
            Left            =   5640
            TabIndex        =   261
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   64
            Left            =   6960
            TabIndex        =   260
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "当前剩余电量 "
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   65
            Left            =   120
            TabIndex        =   259
            Top             =   6600
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   66
            Left            =   2760
            TabIndex        =   258
            Top             =   6600
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mAH"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   67
            Left            =   4320
            TabIndex        =   257
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   68
            Left            =   5640
            TabIndex        =   256
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   69
            Left            =   6960
            TabIndex        =   255
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "循环放电次数"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   70
            Left            =   120
            TabIndex        =   254
            Top             =   7080
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   71
            Left            =   2760
            TabIndex        =   253
            Top             =   7080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "次"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   72
            Left            =   4320
            TabIndex        =   252
            Top             =   7080
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   73
            Left            =   5640
            TabIndex        =   251
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   74
            Left            =   6960
            TabIndex        =   250
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "最高电压值"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   75
            Left            =   120
            TabIndex        =   249
            Top             =   7560
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   76
            Left            =   2760
            TabIndex        =   248
            Top             =   7560
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   77
            Left            =   4320
            TabIndex        =   247
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   60
            Left            =   5640
            TabIndex        =   246
            Top             =   8040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   61
            Left            =   6960
            TabIndex        =   245
            Top             =   8040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "最高电压单体序号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   62
            Left            =   120
            TabIndex        =   244
            Top             =   8040
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   78
            Left            =   2760
            TabIndex        =   243
            Top             =   8040
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   79
            Left            =   4320
            TabIndex        =   242
            Top             =   8040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   80
            Left            =   5640
            TabIndex        =   241
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   81
            Left            =   6960
            TabIndex        =   240
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "RTC通讯状态"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   82
            Left            =   120
            TabIndex        =   239
            Top             =   8520
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   83
            Left            =   2760
            TabIndex        =   238
            Top             =   8520
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   84
            Left            =   4320
            TabIndex        =   237
            Top             =   8520
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   85
            Left            =   5640
            TabIndex        =   236
            Top             =   9000
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   86
            Left            =   6960
            TabIndex        =   235
            Top             =   9000
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "EEPROM通讯状态"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   87
            Left            =   120
            TabIndex        =   234
            Top             =   9000
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   88
            Left            =   2760
            TabIndex        =   233
            Top             =   9000
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   89
            Left            =   4320
            TabIndex        =   232
            Top             =   9000
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   90
            Left            =   5640
            TabIndex        =   231
            Top             =   9480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   91
            Left            =   6960
            TabIndex        =   230
            Top             =   9480
            Width           =   1215
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "AFE通讯状态"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   92
            Left            =   120
            TabIndex        =   229
            Top             =   9480
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   93
            Left            =   2760
            TabIndex        =   228
            Top             =   9480
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H80000014&
            Caption         =   "mV"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   94
            Left            =   4320
            TabIndex        =   227
            Top             =   9480
            Width           =   1215
         End
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   9735
         Index           =   0
         LargeChange     =   50
         Left            =   9600
         Max             =   100
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   9975
         Index           =   1
         LargeChange     =   50
         Left            =   18720
         Max             =   100
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame zPackFrame 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   9975
      Left            =   1560
      TabIndex        =   8
      Top             =   0
      Width           =   19695
      Begin VB.OLE zOLE1 
         AutoActivate    =   1  'GetFocus
         BackColor       =   &H80000018&
         Class           =   "Excel.Sheet.12"
         Height          =   9780
         Left            =   0
         OleObjectBlob   =   "frmMain.frx":0000
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Users\Administrator\Desktop\第5章  电机驱动监控系统\zgf 配置.xlsx"
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   20130
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   1560
      TabIndex        =   328
      Top             =   9960
      Width           =   3855
      Begin VB.Label Label4 
         Caption         =   "MinCell:电芯05 4.5V "
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   338
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "MaxCell:电芯05 4.5V "
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   337
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "SOC:"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   336
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "实时电流:"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   335
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "总电压:"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   334
         Top             =   0
         Width           =   3255
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "未打开端口"
      ForeColor       =   &H80000010&
      Height          =   255
      Index           =   4
      Left            =   10500
      TabIndex        =   340
      Top             =   11760
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "0A"
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   333
      Top             =   11760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "下位机设备地址："
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   332
      Top             =   11760
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "未能成功连接下位机，请重新检查！"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   331
      Top             =   11760
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RichPower"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   330
      Top             =   11760
      Width           =   3135
   End
   Begin VB.Menu MNU_File 
      Caption         =   "文件(&F)"
      Begin VB.Menu MNU_File_Save 
         Caption         =   "保存(&S)..."
      End
      Begin VB.Menu MNU_File_Close 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu MNU_File_Open 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_File_Exit 
         Caption         =   "退出(&x)"
      End
   End
   Begin VB.Menu MNU_Edit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu MNU_Edit_Edit 
         Caption         =   "编辑(&E)"
      End
      Begin VB.Menu MNU_Edit_EditDone 
         Caption         =   "编辑完成(&F)"
      End
      Begin VB.Menu MNU_Edit_Clear 
         Caption         =   "清除(&C)"
      End
   End
   Begin VB.Menu MNU_Option 
      Caption         =   "运行(&R)"
      Begin VB.Menu MNU_Run_Run 
         Caption         =   "运行(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MNU_Config 
      Caption         =   "设置(&C)"
      Begin VB.Menu MNU_Config_Port 
         Caption         =   "设置串口(&P)"
      End
      Begin VB.Menu MNU_Config_Code 
         Caption         =   "设置控制码(&C)"
      End
   End
   Begin VB.Menu MNU_Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MNU_About 
         Caption         =   "关于本软件&A"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    '用于接收的变量
Dim ComSwitch As Boolean                                                            ' 串口开关状态判断
Dim FileData As String                                                              ' 要发送的文件暂存
Dim SendCount As Long                                                               ' 发送数据字节计数器
Dim ReceiveCount As Long                                                            ' 接收数据字节计数器
Dim InputSignal As String                                                           ' 接收缓冲暂存
Dim OutputSignal As String                                                          ' 发送数据暂存
Dim DisplaySwitch As Boolean                                                        ' 显示开关
Dim ModeSend As Boolean                                                             ' 发送方式判断
Dim Savetime As Single                                                              ' 时间数据暂存 延时用
Dim SaveTextPath As String


Private Sub cmdCfgSPort_Click(Index As Integer)
  chuanshu.Show 1, frmMain
End Sub

Private Sub cmdConnect_Click(Index As Integer)
    chuanshu.MSComm1.PortOpen = False
    chuanshu.MSComm1.PortOpen = True
End Sub

Private Sub cmdMcuSysConfig_Click(Index As Integer)
zFrame7.Visible = False
ZFrame1(10).Visible = False
ZFrame1(11).Visible = False
zPackFrame.Visible = True
zOLE1.Visible = True
End Sub

Private Sub cmdPackInfo_Click(Index As Integer)
zPackFrame.Visible = False
zOLE1.Visible = False
zFrame7.Visible = True
ZFrame1(10).Visible = True
ZFrame1(11).Visible = True
End Sub

Private Sub Form_Load()

Call chuanshu.myInitForm_Load

End Sub

Private Sub MNU_Config_Port_Click()
  chuanshu.Show 1, frmMain
End Sub

Private Sub Timer_Timer()
  Text2(0).Text = Now()
End Sub
 
 
