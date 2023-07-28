VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "名贝科技."
   ClientHeight    =   12720
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   23490
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H00800000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   848
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1566
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command3CAP 
      Caption         =   "CAP参数"
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
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   307
      Top             =   6360
      Width           =   1496
   End
   Begin VB.Frame 生产解码 
      Caption         =   "自动生产解码中"
      Height          =   11535
      Left            =   19080
      TabIndex        =   293
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton Command3 
         Caption         =   "清除"
         Height          =   1335
         Left            =   0
         TabIndex        =   308
         Top             =   6720
         Width           =   495
      End
      Begin VB.TextBox 自动配置结果 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4815
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   296
         Text            =   "frmMain.frx":10CA
         Top             =   6240
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaskColor       =   &H0000C000&
         TabIndex        =   295
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton 配置下发开始按钮 
         Caption         =   "开始"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   24
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         MaskColor       =   &H0000C000&
         TabIndex        =   294
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label AUTO_NUM 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2520
         TabIndex        =   309
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label 记录擦除 
         BackColor       =   &H80000014&
         Caption         =   "√记录擦除"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   435
         Left            =   480
         TabIndex        =   306
         Top             =   5158
         Width           =   3015
      End
      Begin VB.Label 系统参数2下发 
         BackColor       =   &H80000014&
         Caption         =   "√系统参数2下发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   305
         Top             =   2778
         Width           =   3015
      End
      Begin VB.Label 时间校正 
         BackColor       =   &H80000014&
         Caption         =   "√时间校正"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   304
         Top             =   4206
         Width           =   3015
      End
      Begin VB.Label 电流校零 
         BackColor       =   &H80000014&
         Caption         =   "√电流校零"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   303
         Top             =   4682
         Width           =   3015
      End
      Begin VB.Label CAP下发 
         BackColor       =   &H80000014&
         Caption         =   "√CAP下发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   302
         Top             =   3254
         Width           =   3015
      End
      Begin VB.Label OCV下发 
         BackColor       =   &H80000014&
         Caption         =   "√OCV下发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   301
         Top             =   3730
         Width           =   3015
      End
      Begin VB.Label 解码 
         BackColor       =   &H80000014&
         Caption         =   "√解码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   435
         Left            =   480
         TabIndex        =   300
         Top             =   1350
         Width           =   3015
      End
      Begin VB.Label 系统参数1下发 
         BackColor       =   &H80000014&
         Caption         =   "√系统参数1下发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   299
         Top             =   2302
         Width           =   3015
      End
      Begin VB.Label 硬件配置下发 
         BackColor       =   &H80000014&
         Caption         =   "√硬件配置下发"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   480
         TabIndex        =   298
         Top             =   1826
         Width           =   3015
      End
      Begin VB.Label 连续 
         BackColor       =   &H80000014&
         Caption         =   "√单台配置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   480
         TabIndex        =   297
         Top             =   5640
         Width           =   3015
      End
   End
   Begin VB.Frame Framecap 
      Caption         =   "cap参数"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   10575
      Left            =   21960
      TabIndex        =   284
      Top             =   0
      Width           =   1440
      Begin VB.CommandButton Command_reand_cap 
         BackColor       =   &H0000C000&
         Caption         =   "读取cap"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   288
         Top             =   9900
         Width           =   2115
      End
      Begin VB.CommandButton Command_writeCAP 
         BackColor       =   &H0000C000&
         Caption         =   "设置cap"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3540
         TabIndex        =   287
         Top             =   9900
         Width           =   2115
      End
      Begin VB.CommandButton Commandopencap 
         BackColor       =   &H0000C000&
         Caption         =   "载入cap"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13200
         TabIndex        =   286
         Top             =   9840
         Width           =   2115
      End
      Begin VB.CommandButton CommandSavecap 
         BackColor       =   &H0000C000&
         Caption         =   "保存cap"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   16740
         TabIndex        =   285
         Top             =   9780
         Width           =   2115
      End
      Begin MSComDlg.CommonDialog CommonDialog_cap 
         Left            =   8820
         Top             =   9900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridcap 
         Height          =   9455
         Left            =   300
         TabIndex        =   289
         Top             =   480
         Width           =   18555
         _ExtentX        =   32729
         _ExtentY        =   16669
         _Version        =   393216
         Rows            =   22
         Cols            =   6
         ForeColor       =   12582912
         ForeColorFixed  =   32768
         MousePointer    =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label get_real_cap 
         Caption         =   "Label5"
         Height          =   615
         Left            =   6840
         TabIndex        =   290
         Top             =   9960
         Width           =   5655
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog_jilv 
      Left            =   1800
      Top             =   11040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame_SYSMUC2 
      Caption         =   "系统参数2"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   10815
      Left            =   1920
      TabIndex        =   241
      Top             =   240
      Width           =   11295
      Begin MSComDlg.CommonDialog CommonDialogsys2 
         Left            =   3120
         Top             =   8760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command_Sys2Write 
         Caption         =   "下发配置"
         Height          =   435
         Left            =   5280
         TabIndex        =   247
         Top             =   7560
         Width           =   1455
      End
      Begin VB.CommandButton Command_sys2load 
         Caption         =   "加载配置"
         Height          =   435
         Left            =   180
         TabIndex        =   246
         Top             =   7620
         Width           =   1155
      End
      Begin VB.CommandButton Command_Sys2Read 
         Caption         =   "读取配置"
         Height          =   435
         Left            =   3360
         TabIndex        =   245
         Top             =   7620
         Width           =   1455
      End
      Begin VB.CommandButton Command_Sys2Save 
         Caption         =   "保存配置"
         Height          =   435
         Left            =   1560
         TabIndex        =   244
         Top             =   7680
         Width           =   1335
      End
      Begin VB.TextBox TexSys2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   242
         Text            =   "01"
         Top             =   420
         Width           =   4035
      End
      Begin VB.Label LabeSYS2 
         BackColor       =   &H8000000E&
         Caption         =   "保护板地址"
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
         Left            =   -300
         TabIndex        =   243
         Top             =   660
         Width           =   1815
      End
   End
   Begin VB.Frame FrameSOC_OCV 
      Caption         =   "SOC-OCV参数"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   10575
      Left            =   20760
      TabIndex        =   235
      Top             =   0
      Width           =   600
      Begin MSComDlg.CommonDialog CommonDialog_SOCOCV 
         Left            =   8820
         Top             =   9900
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton CommandSaveOCV 
         BackColor       =   &H0000C000&
         Caption         =   "保存OCV"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   16740
         TabIndex        =   240
         Top             =   9780
         Width           =   2115
      End
      Begin VB.CommandButton CommandopenOCV 
         BackColor       =   &H0000C000&
         Caption         =   "载入OCV"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13200
         TabIndex        =   239
         Top             =   9840
         Width           =   2115
      End
      Begin VB.CommandButton CommandwriteOCV 
         BackColor       =   &H0000C000&
         Caption         =   "设置OCV"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3540
         TabIndex        =   238
         Top             =   9900
         Width           =   2115
      End
      Begin VB.CommandButton CommandreadOCV 
         BackColor       =   &H0000C000&
         Caption         =   "读取OCV"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   480
         TabIndex        =   237
         Top             =   9900
         Width           =   2115
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   7455
         Left            =   300
         TabIndex        =   236
         Top             =   480
         Width           =   18555
         _ExtentX        =   32729
         _ExtentY        =   13150
         _Version        =   393216
         Rows            =   22
         Cols            =   6
         ForeColor       =   12582912
         ForeColorFixed  =   32768
         MousePointer    =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame_Record 
      BackColor       =   &H00E0E0E0&
      Caption         =   "记录查询"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   11715
      Left            =   20640
      TabIndex        =   228
      Top             =   -120
      Width           =   615
      Begin MSComDlg.CommonDialog CommonDialog_Record 
         Left            =   17820
         Top             =   9360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command_SaveBlackupData 
         Caption         =   "保存"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   16920
         TabIndex        =   234
         Top             =   10920
         Width           =   1455
      End
      Begin VB.CommandButton Command_clearBlackup 
         Caption         =   "清除界面"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   9240
         TabIndex        =   233
         Top             =   10920
         Width           =   1995
      End
      Begin VB.CommandButton Command_EraseBalckUp 
         BackColor       =   &H00000080&
         Caption         =   "擦除记录"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   13800
         MaskColor       =   &H000000FF&
         OLEDropMode     =   1  'Manual
         TabIndex        =   232
         Top             =   10920
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "暂停"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5040
         TabIndex        =   231
         Top             =   10920
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000C000&
         Caption         =   "读取记录"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2160
         TabIndex        =   230
         Top             =   10920
         Width           =   2115
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   9435
         Left            =   3840
         TabIndex        =   229
         Top             =   300
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   16642
         _Version        =   393216
         ScrollTrack     =   -1  'True
      End
   End
   Begin VB.Frame Frame_Calib 
      Caption         =   "校准界面"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   11715
      Left            =   13920
      TabIndex        =   158
      Top             =   0
      Width           =   1785
      Begin VB.Frame Frame7 
         Caption         =   "校正复位"
         Height          =   1815
         Left            =   600
         TabIndex        =   291
         Top             =   720
         Width           =   6375
         Begin VB.CommandButton Command_resetall 
            BackColor       =   &H000000C0&
            Caption         =   "全部校正复位"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1200
            MaskColor       =   &H00C00000&
            TabIndex        =   292
            Top             =   480
            Width           =   2955
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1875
         Left            =   660
         TabIndex        =   264
         Top             =   8220
         Width           =   6195
         Begin VB.TextBox Text_TimeRTC 
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   435
            Left            =   720
            TabIndex        =   267
            Text            =   "Text6"
            Top             =   660
            Width           =   3075
         End
         Begin VB.CommandButton Command_rtcTIME 
            BackColor       =   &H000000C0&
            Caption         =   "更新"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4620
            MaskColor       =   &H00C00000&
            TabIndex        =   266
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "RTC时钟校正"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   19
            Left            =   1860
            TabIndex        =   265
            Top             =   -60
            Width           =   2595
         End
      End
      Begin VB.Frame Frame5 
         Height          =   9615
         Index           =   2
         Left            =   7980
         TabIndex        =   177
         Top             =   600
         Width           =   7635
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   9
            Left            =   1560
            TabIndex        =   218
            Text            =   "25.0"
            Top             =   6660
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   9
            Left            =   4080
            TabIndex        =   215
            Text            =   "25.0"
            Top             =   6660
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   8
            Left            =   1560
            TabIndex        =   211
            Text            =   "25.0"
            Top             =   6002
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   8
            Left            =   4080
            TabIndex        =   210
            Text            =   "25.0"
            Top             =   6002
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   7
            Left            =   1560
            TabIndex        =   207
            Text            =   "25.0"
            Top             =   5316
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   7
            Left            =   4080
            TabIndex        =   206
            Text            =   "25.0"
            Top             =   5316
            Width           =   1575
         End
         Begin VB.CommandButton Command_Temp 
            BackColor       =   &H000000C0&
            Caption         =   "下发校正"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2700
            MaskColor       =   &H00C00000&
            TabIndex        =   205
            Top             =   8820
            Width           =   1815
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   6
            Left            =   1560
            TabIndex        =   203
            Text            =   "25.0"
            Top             =   4630
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   6
            Left            =   4080
            TabIndex        =   202
            Text            =   "25.0"
            Top             =   4630
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   5
            Left            =   1560
            TabIndex        =   199
            Text            =   "25.0"
            Top             =   3944
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   5
            Left            =   4080
            TabIndex        =   198
            Text            =   "25.0"
            Top             =   3944
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   4
            Left            =   1560
            TabIndex        =   195
            Text            =   "25.0"
            Top             =   3258
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   4
            Left            =   4080
            TabIndex        =   194
            Text            =   "25.0"
            Top             =   3258
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            HelpContextID   =   3
            Index           =   3
            Left            =   1560
            TabIndex        =   191
            Text            =   "25.0"
            Top             =   2572
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   3
            Left            =   4080
            TabIndex        =   190
            Text            =   "25.0"
            Top             =   2572
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   2
            Left            =   1560
            TabIndex        =   187
            Text            =   "25.0"
            Top             =   1886
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   2
            Left            =   4080
            TabIndex        =   186
            Text            =   "25.0"
            Top             =   1886
            Width           =   1575
         End
         Begin VB.TextBox disTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   1
            Left            =   1560
            TabIndex        =   179
            Text            =   "25.0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox RealTemptext 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   1
            Left            =   4080
            TabIndex        =   178
            Text            =   "25.0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "全选"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   435
            Index           =   18
            Left            =   5400
            TabIndex        =   258
            Top             =   7500
            Width           =   795
         End
         Begin VB.Label LabelwenduCaliball 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Left            =   6300
            TabIndex        =   257
            Top             =   7440
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   8
            Left            =   6300
            TabIndex        =   256
            Top             =   6720
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   7
            Left            =   6300
            TabIndex        =   255
            Top             =   6034
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   6
            Left            =   6300
            TabIndex        =   254
            Top             =   5352
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   5
            Left            =   6300
            TabIndex        =   253
            Top             =   4670
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   4
            Left            =   6300
            TabIndex        =   252
            Top             =   3988
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   3
            Left            =   6300
            TabIndex        =   251
            Top             =   3306
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   2
            Left            =   6300
            TabIndex        =   250
            Top             =   2624
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   1
            Left            =   6300
            TabIndex        =   249
            Top             =   1942
            Width           =   495
         End
         Begin VB.Label LabelwenduCalib 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   0
            Left            =   6300
            TabIndex        =   248
            Top             =   1260
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度7"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   17
            Left            =   240
            TabIndex        =   227
            Top             =   6780
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   28
            Left            =   5700
            TabIndex        =   226
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   27
            Left            =   5700
            TabIndex        =   225
            Top             =   1942
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   26
            Left            =   5700
            TabIndex        =   224
            Top             =   2624
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   25
            Left            =   5700
            TabIndex        =   223
            Top             =   3306
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   24
            Left            =   5700
            TabIndex        =   222
            Top             =   3988
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   23
            Left            =   5700
            TabIndex        =   221
            Top             =   4670
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   22
            Left            =   5700
            TabIndex        =   220
            Top             =   5352
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   15
            Left            =   5700
            TabIndex        =   219
            Top             =   6034
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   21
            Left            =   3180
            TabIndex        =   217
            Top             =   4670
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   20
            Left            =   3180
            TabIndex        =   216
            Top             =   3306
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度6"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   16
            Left            =   240
            TabIndex        =   214
            Top             =   6090
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   19
            Left            =   3180
            TabIndex        =   213
            Top             =   6720
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   18
            Left            =   5700
            TabIndex        =   212
            Top             =   6720
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度5"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   15
            Left            =   240
            TabIndex        =   209
            Top             =   5400
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   17
            Left            =   3180
            TabIndex        =   208
            Top             =   6034
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度4"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   14
            Left            =   240
            TabIndex        =   204
            Top             =   4710
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度3"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   13
            Left            =   240
            TabIndex        =   201
            Top             =   4020
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   13
            Left            =   3180
            TabIndex        =   200
            Top             =   5352
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度2"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   12
            Left            =   240
            TabIndex        =   197
            Top             =   3330
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   11
            Left            =   3180
            TabIndex        =   196
            Top             =   3988
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "电芯温度1"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   11
            Left            =   240
            TabIndex        =   193
            Top             =   2640
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   9
            Left            =   3180
            TabIndex        =   192
            Top             =   2624
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "环境温度"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   10
            Left            =   240
            TabIndex        =   189
            Top             =   1950
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   7
            Left            =   3180
            TabIndex        =   188
            Top             =   1942
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "MOS温度"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   435
            Index           =   9
            Left            =   240
            TabIndex        =   185
            Top             =   1260
            Width           =   1395
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   5
            Left            =   3180
            TabIndex        =   184
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "℃"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   4
            Left            =   1140
            TabIndex        =   183
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "温度校正"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   8
            Left            =   2760
            TabIndex        =   182
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "当前值"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   7
            Left            =   1620
            TabIndex        =   181
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "实际值"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   375
            Index           =   6
            Left            =   4260
            TabIndex        =   180
            Top             =   720
            Width           =   1395
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3435
         Index           =   1
         Left            =   600
         TabIndex        =   168
         Top             =   4560
         Width           =   6315
         Begin VB.CommandButton Command_Currentz 
            BackColor       =   &H000000C0&
            Caption         =   "零点"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4800
            MaskColor       =   &H00C00000&
            TabIndex        =   260
            Top             =   2460
            Width           =   1155
         End
         Begin VB.TextBox RealCurrentzText 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   2760
            TabIndex        =   259
            Text            =   "0.0"
            Top             =   2580
            Width           =   1575
         End
         Begin VB.TextBox DisCurrentText 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   171
            Text            =   "19.0"
            Top             =   1740
            Width           =   1575
         End
         Begin VB.TextBox RealCurrentkText 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   2760
            TabIndex        =   170
            Text            =   "22.0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton Command_Currentk 
            BackColor       =   &H000000C0&
            Caption         =   "标定"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4800
            MaskColor       =   &H00C00000&
            TabIndex        =   169
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "零点电流"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   375
            Index           =   20
            Left            =   2820
            TabIndex        =   262
            Top             =   2100
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "A"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   8
            Left            =   4380
            TabIndex        =   261
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "A"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   3
            Left            =   1860
            TabIndex        =   176
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "A"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   2
            Left            =   4380
            TabIndex        =   175
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "电流校正"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   5
            Left            =   1920
            TabIndex        =   174
            Top             =   -60
            Width           =   1875
         End
         Begin VB.Label Label3 
            Caption         =   "当前值"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   4
            Left            =   360
            TabIndex        =   173
            Top             =   1260
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "实际值"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   375
            Index           =   3
            Left            =   2820
            TabIndex        =   172
            Top             =   720
            Width           =   1395
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Index           =   0
         Left            =   600
         TabIndex        =   159
         Top             =   2520
         Width           =   6315
         Begin VB.CommandButton Command1_Volte 
            BackColor       =   &H000000C0&
            Caption         =   "校正"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4800
            MaskColor       =   &H00C00000&
            TabIndex        =   167
            Top             =   1080
            Width           =   1155
         End
         Begin VB.TextBox RealVolteText 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   2760
            TabIndex        =   163
            Text            =   "48.0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox DisVolteText 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   160
            Text            =   "48.0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "实际值"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   375
            Index           =   2
            Left            =   2820
            TabIndex        =   166
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "当前值"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   1
            Left            =   240
            TabIndex        =   165
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label3 
            Caption         =   "总电压校正"
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   555
            Index           =   0
            Left            =   1740
            TabIndex        =   164
            Top             =   -60
            Width           =   2355
         End
         Begin VB.Label Label2 
            Caption         =   "V"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   1
            Left            =   4380
            TabIndex        =   162
            Top             =   1260
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "V"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   0
            Left            =   1740
            TabIndex        =   161
            Top             =   1260
            Width           =   375
         End
      End
   End
   Begin VB.CommandButton cmdBQSYS 
      Caption         =   "系统参数2"
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
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   157
      Top             =   3390
      Width           =   1496
   End
   Begin VB.CommandButton cmdSOC 
      Caption         =   "SOC参数"
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
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   156
      Top             =   5450
      Width           =   1496
   End
   Begin VB.Frame FrameREG 
      Caption         =   "硬件配置"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   10200
      Left            =   6720
      TabIndex        =   37
      Top             =   120
      Width           =   4755
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1035
         Left            =   15600
         MultiLine       =   -1  'True
         TabIndex        =   155
         Text            =   "frmMain.frx":10CF
         Top             =   7680
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.ComboBox ComboRtemp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         ItemData        =   "frmMain.frx":10D5
         Left            =   8760
         List            =   "frmMain.frx":10E5
         TabIndex        =   154
         Top             =   9120
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ComboBox ComboRtemp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         ItemData        =   "frmMain.frx":10FD
         Left            =   8760
         List            =   "frmMain.frx":1131
         TabIndex        =   153
         Top             =   9060
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ComboBox ComboRtemp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         ItemData        =   "frmMain.frx":118D
         Left            =   8700
         List            =   "frmMain.frx":11C1
         TabIndex        =   152
         Top             =   9000
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ComboBox ComboRtemp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         ItemData        =   "frmMain.frx":121E
         Left            =   8760
         List            =   "frmMain.frx":1252
         TabIndex        =   151
         Top             =   9000
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ComboBox ComboRtemp 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         ItemData        =   "frmMain.frx":12AE
         Left            =   8640
         List            =   "frmMain.frx":12E2
         TabIndex        =   150
         Top             =   9000
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Frame FrameS1 
         BackColor       =   &H8000000B&
         Caption         =   "系统寄存器配置"
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
         Height          =   8500
         Left            =   60
         TabIndex        =   124
         Top             =   480
         Width           =   5175
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   148
            Text            =   "电池节数"
            Top             =   480
            Width           =   1515
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   147
            Text            =   "过流MOS控制"
            Top             =   2442
            Width           =   2235
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   146
            Text            =   "放电恢复充电MOS"
            Top             =   1788
            Width           =   2355
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   145
            Text            =   "MCU控制平衡"
            Top             =   3096
            Width           =   2535
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   144
            Text            =   "预充电控制"
            Top             =   1134
            Width           =   2175
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   143
            Text            =   "二级过压保护使能"
            Top             =   3750
            Width           =   2415
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   142
            Text            =   "过流保护定时恢复"
            Top             =   6366
            Width           =   2475
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   7
            Left            =   120
            TabIndex        =   141
            Text            =   "负载释放延迟"
            Top             =   5058
            Width           =   2115
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   140
            Text            =   "CTL管脚控制"
            Top             =   5712
            Width           =   2055
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   139
            Text            =   "低压充电功能"
            Top             =   4404
            Width           =   1995
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   138
            Text            =   "过放负载锁定 "
            Top             =   7020
            Width           =   1935
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   11
            Left            =   120
            TabIndex        =   137
            Text            =   "欠压关闭CHG"
            Top             =   7605
            Width           =   1995
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   0
            ItemData        =   "frmMain.frx":1345
            Left            =   2640
            List            =   "frmMain.frx":136D
            TabIndex        =   136
            Top             =   540
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   1
            ItemData        =   "frmMain.frx":139C
            Left            =   2640
            List            =   "frmMain.frx":13A6
            TabIndex        =   135
            Top             =   1185
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   2
            ItemData        =   "frmMain.frx":13CA
            Left            =   2640
            List            =   "frmMain.frx":13D4
            TabIndex        =   134
            Top             =   1845
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   3
            ItemData        =   "frmMain.frx":1408
            Left            =   2640
            List            =   "frmMain.frx":1412
            TabIndex        =   133
            Top             =   2490
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   4
            ItemData        =   "frmMain.frx":143E
            Left            =   2640
            List            =   "frmMain.frx":1448
            TabIndex        =   132
            Top             =   3135
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   5
            ItemData        =   "frmMain.frx":146B
            Left            =   2640
            List            =   "frmMain.frx":1475
            TabIndex        =   131
            Top             =   3780
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   6
            ItemData        =   "frmMain.frx":14A1
            Left            =   2640
            List            =   "frmMain.frx":14AB
            TabIndex        =   130
            Top             =   4440
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   7
            ItemData        =   "frmMain.frx":14CB
            Left            =   2640
            List            =   "frmMain.frx":14DB
            TabIndex        =   129
            Top             =   5100
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   8
            ItemData        =   "frmMain.frx":14FD
            Left            =   2640
            List            =   "frmMain.frx":150D
            TabIndex        =   128
            Top             =   5730
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   9
            ItemData        =   "frmMain.frx":1542
            Left            =   2640
            List            =   "frmMain.frx":154C
            TabIndex        =   127
            Top             =   6375
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   10
            ItemData        =   "frmMain.frx":155E
            Left            =   2640
            List            =   "frmMain.frx":1568
            TabIndex        =   126
            Top             =   7035
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   300
            Index           =   11
            ItemData        =   "frmMain.frx":159E
            Left            =   2640
            List            =   "frmMain.frx":15A8
            TabIndex        =   125
            Top             =   7620
            Width           =   2400
         End
      End
      Begin VB.Frame FrameS2 
         BackColor       =   &H8000000B&
         Caption         =   "电压保护"
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
         Height          =   8500
         Left            =   5350
         TabIndex        =   92
         Top             =   480
         Width           =   5340
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   14
            ItemData        =   "frmMain.frx":15D4
            Left            =   2835
            List            =   "frmMain.frx":1608
            TabIndex        =   123
            Top             =   5685
            Width           =   2400
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   13
            ItemData        =   "frmMain.frx":1662
            Left            =   2835
            List            =   "frmMain.frx":1696
            TabIndex        =   122
            Top             =   2415
            Width           =   2400
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   13
            Left            =   180
            TabIndex        =   121
            Text            =   "低压禁止充电 "
            Top             =   6960
            Width           =   2595
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   14
            Left            =   180
            TabIndex        =   120
            Text            =   "欠压保护恢复"
            Top             =   4368
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   15
            Left            =   180
            TabIndex        =   119
            Text            =   "欠压保护延时"
            Top             =   5664
            Width           =   2055
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   16
            Left            =   180
            TabIndex        =   118
            Text            =   "欠压保护 "
            Top             =   5016
            Width           =   2235
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   17
            Left            =   180
            TabIndex        =   117
            Text            =   "预充电开启 "
            Top             =   6312
            Width           =   2415
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   18
            Left            =   180
            TabIndex        =   116
            Text            =   "平衡开启"
            Top             =   3720
            Width           =   1995
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   19
            Left            =   180
            TabIndex        =   115
            Text            =   "二级压保护延时"
            Top             =   1128
            Width           =   2175
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   20
            Left            =   180
            TabIndex        =   114
            Text            =   "过压保护恢复"
            Top             =   3072
            Width           =   2595
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   21
            Left            =   180
            TabIndex        =   113
            Text            =   "一级过压保护"
            Top             =   1776
            Width           =   1935
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   22
            Left            =   180
            TabIndex        =   112
            Text            =   "一级过压保护延时"
            Top             =   2424
            Width           =   2415
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   23
            Left            =   180
            TabIndex        =   111
            Text            =   "二级过压保护"
            Top             =   480
            Width           =   2355
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   12
            ItemData        =   "frmMain.frx":16F0
            Left            =   2835
            List            =   "frmMain.frx":1700
            TabIndex        =   110
            Top             =   1155
            Width           =   2400
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   0
            Left            =   2835
            TabIndex        =   109
            Text            =   "4400"
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   24
            Left            =   3960
            TabIndex        =   108
            Text            =   "mV"
            Top             =   540
            Width           =   555
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   1
            Left            =   2835
            TabIndex        =   107
            Text            =   "4250"
            Top             =   1740
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   2
            Left            =   2835
            TabIndex        =   106
            Text            =   "4100"
            Top             =   3000
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   3
            Left            =   2835
            TabIndex        =   105
            Text            =   "4000"
            Top             =   3675
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   4
            Left            =   2835
            TabIndex        =   104
            Text            =   "3200"
            Top             =   4350
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   5
            Left            =   2835
            TabIndex        =   103
            Text            =   "3000"
            Top             =   4980
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   6
            Left            =   2835
            TabIndex        =   102
            Text            =   "2900"
            Top             =   6270
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   375
            Index           =   7
            Left            =   2835
            TabIndex        =   101
            Text            =   "2500"
            Top             =   6945
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   25
            Left            =   3960
            TabIndex        =   100
            Text            =   "mV"
            Top             =   1800
            Width           =   555
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   42
            Left            =   3960
            TabIndex        =   99
            Text            =   "mV"
            Top             =   3000
            Width           =   555
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   43
            Left            =   3960
            TabIndex        =   98
            Text            =   "mV"
            Top             =   3720
            Width           =   555
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   44
            Left            =   3960
            TabIndex        =   97
            Text            =   "mV"
            Top             =   4380
            Width           =   555
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   45
            Left            =   3960
            TabIndex        =   96
            Text            =   "mV"
            Top             =   5040
            Width           =   555
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   46
            Left            =   3960
            TabIndex        =   95
            Text            =   "mV"
            Top             =   6300
            Width           =   555
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   47
            Left            =   3960
            TabIndex        =   94
            Text            =   "mV"
            Top             =   6960
            Width           =   555
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   675
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   93
            Text            =   "frmMain.frx":1717
            Top             =   7800
            Visible         =   0   'False
            Width           =   4875
         End
      End
      Begin VB.Frame FrameS3 
         BackColor       =   &H8000000B&
         Caption         =   "电流保护"
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
         Height          =   8500
         Left            =   10805
         TabIndex        =   67
         Top             =   480
         Width           =   4680
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   405
            Index           =   20
            ItemData        =   "frmMain.frx":171D
            Left            =   2880
            List            =   "frmMain.frx":1751
            TabIndex        =   149
            Top             =   3240
            Width           =   1440
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   615
            Index           =   12
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   91
            Text            =   "frmMain.frx":17B4
            Top             =   7440
            Width           =   1635
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   405
            Index           =   26
            ItemData        =   "frmMain.frx":17CB
            Left            =   2880
            List            =   "frmMain.frx":17DB
            TabIndex        =   90
            Top             =   7500
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   25
            ItemData        =   "frmMain.frx":17F3
            Left            =   2880
            List            =   "frmMain.frx":1803
            TabIndex        =   89
            Top             =   6660
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   24
            ItemData        =   "frmMain.frx":181A
            Left            =   2880
            List            =   "frmMain.frx":182A
            TabIndex        =   88
            Top             =   6060
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   23
            ItemData        =   "frmMain.frx":1849
            Left            =   2880
            List            =   "frmMain.frx":187D
            TabIndex        =   87
            Top             =   4920
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   405
            Index           =   22
            ItemData        =   "frmMain.frx":18DC
            Left            =   2880
            List            =   "frmMain.frx":1910
            TabIndex        =   86
            Top             =   4362
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   21
            ItemData        =   "frmMain.frx":196C
            Left            =   2880
            List            =   "frmMain.frx":19A0
            TabIndex        =   85
            Top             =   3805
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   19
            ItemData        =   "frmMain.frx":1A11
            Left            =   2880
            List            =   "frmMain.frx":1A45
            TabIndex        =   84
            Top             =   2691
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   405
            Index           =   18
            ItemData        =   "frmMain.frx":1AA4
            Left            =   2880
            List            =   "frmMain.frx":1AD8
            TabIndex        =   83
            Top             =   2100
            Width           =   1440
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Index           =   17
            ItemData        =   "frmMain.frx":1B35
            Left            =   2880
            List            =   "frmMain.frx":1B69
            TabIndex        =   82
            Top             =   1560
            Width           =   1440
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   59
            Left            =   4080
            TabIndex        =   81
            Text            =   "mΩ"
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
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
            Index           =   16
            Left            =   3240
            TabIndex        =   80
            Text            =   "1"
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   375
            Index           =   58
            Left            =   180
            TabIndex        =   79
            Text            =   "计算电流用采样电阻值"
            Top             =   420
            Width           =   3195
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   495
            Index           =   57
            Left            =   120
            TabIndex        =   78
            Text            =   "过流自恢复延时"
            Top             =   6720
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   55
            Left            =   120
            TabIndex        =   77
            Text            =   "充放电MOS开启延时"
            Top             =   6180
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   54
            Left            =   120
            TabIndex        =   76
            Text            =   "充电过流保护延时"
            Top             =   5040
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   49
            Left            =   120
            TabIndex        =   75
            Text            =   "充电过流保护阀值"
            Top             =   4464
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   53
            Left            =   120
            TabIndex        =   74
            Text            =   "短路保护延时"
            Top             =   3890
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   52
            Left            =   120
            TabIndex        =   73
            Text            =   "短路电流保护阀值"
            Top             =   3316
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   51
            Left            =   120
            TabIndex        =   72
            Text            =   "放电过流2保护延时"
            Top             =   2742
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   50
            Left            =   120
            TabIndex        =   71
            Text            =   "放电过流2保护阀值"
            Top             =   2168
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   48
            Left            =   120
            TabIndex        =   70
            Text            =   "放电过流1保护延时"
            Top             =   1594
            Width           =   2655
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   56
            Left            =   120
            TabIndex        =   69
            Text            =   "放电过流1保护阀值"
            Top             =   1020
            Width           =   2655
         End
         Begin VB.ComboBox ComboR1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   405
            Index           =   16
            ItemData        =   "frmMain.frx":1BC3
            Left            =   2880
            List            =   "frmMain.frx":1BF7
            TabIndex        =   68
            Top             =   1020
            Width           =   1440
         End
      End
      Begin VB.CommandButton CmdRegReadout 
         Caption         =   "读取配置"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10800
         TabIndex        =   66
         Top             =   9180
         Width           =   1575
      End
      Begin VB.CommandButton CmdRegSent 
         Caption         =   "下发配置"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   15000
         TabIndex        =   65
         Top             =   9180
         Width           =   1575
      End
      Begin VB.CommandButton CmdRegSave 
         Caption         =   "保存配置"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3660
         TabIndex        =   64
         Top             =   9200
         Width           =   1575
      End
      Begin VB.CommandButton CmdRegjiazai 
         Caption         =   "加载配置"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   63
         Top             =   9200
         Width           =   1635
      End
      Begin VB.Frame FrameS4 
         BackColor       =   &H8000000B&
         Caption         =   "温度保护"
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
         Height          =   8500
         Left            =   15600
         TabIndex        =   38
         Top             =   480
         Width           =   4080
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   41
            Left            =   3600
            TabIndex        =   62
            Text            =   "℃"
            Top             =   4920
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   420
            Index           =   15
            Left            =   2520
            TabIndex        =   61
            Text            =   "-5"
            Top             =   4180
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   40
            Left            =   3600
            TabIndex        =   60
            Text            =   "℃"
            Top             =   4260
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   420
            Index           =   14
            Left            =   2520
            TabIndex        =   59
            Text            =   "-10"
            Top             =   4800
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   39
            Left            =   3600
            TabIndex        =   58
            Text            =   "℃"
            Top             =   3645
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   420
            Index           =   13
            Left            =   2520
            TabIndex        =   57
            Text            =   "65"
            Top             =   3560
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   38
            Left            =   3600
            TabIndex        =   56
            Text            =   "℃"
            Top             =   3000
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Index           =   12
            Left            =   2520
            TabIndex        =   55
            Text            =   "70"
            Top             =   2985
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   31
            Left            =   3600
            TabIndex        =   54
            Text            =   "℃"
            Top             =   2385
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   420
            Index           =   11
            Left            =   2520
            TabIndex        =   53
            Text            =   "5"
            Top             =   1725
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   29
            Left            =   3600
            TabIndex        =   52
            Text            =   "℃"
            Top             =   1755
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   420
            Index           =   10
            Left            =   2520
            TabIndex        =   51
            Text            =   "0"
            Top             =   2355
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   27
            Left            =   3600
            TabIndex        =   50
            Text            =   "℃"
            Top             =   1110
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Index           =   9
            Left            =   2520
            TabIndex        =   49
            Text            =   "55"
            Top             =   1110
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   26
            Left            =   3600
            TabIndex        =   48
            Text            =   "℃"
            Top             =   480
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   375
            Index           =   8
            Left            =   2520
            TabIndex        =   47
            Text            =   "60"
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   37
            Left            =   60
            TabIndex        =   46
            Text            =   "充电高温保护"
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   36
            Left            =   60
            TabIndex        =   45
            Text            =   "充电低温保护释放"
            Top             =   1748
            Width           =   2715
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   35
            Left            =   60
            TabIndex        =   44
            Text            =   "充电低温保护"
            Top             =   2382
            Width           =   1815
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   34
            Left            =   60
            TabIndex        =   43
            Text            =   "放电高温保护"
            Top             =   3016
            Width           =   1515
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   33
            Left            =   60
            TabIndex        =   42
            Text            =   "充电高温保护释放"
            Top             =   1114
            Width           =   2595
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   32
            Left            =   60
            TabIndex        =   41
            Text            =   "放电高温保护释放"
            Top             =   3650
            Width           =   2535
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   30
            Left            =   120
            TabIndex        =   40
            Text            =   "放电低温保护释放"
            Top             =   4284
            Width           =   2475
         End
         Begin VB.TextBox TEXTRE1 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "楷体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000012&
            Height          =   315
            Index           =   28
            Left            =   60
            TabIndex        =   39
            Text            =   "放电低温保护"
            Top             =   4920
            Width           =   1995
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   660
      Top             =   10980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame_SYSCONFIG 
      Caption         =   "系统参数1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   11655
      Left            =   3960
      TabIndex        =   30
      Top             =   0
      Width           =   3735
      Begin VB.Frame Frame6 
         Height          =   1275
         Left            =   15840
         TabIndex        =   270
         Top             =   9960
         Width           =   4695
         Begin VB.Label Label4 
            Caption         =   "电芯4 温度"
            Height          =   435
            Index           =   6
            Left            =   4080
            TabIndex        =   282
            Top             =   720
            Width           =   555
         End
         Begin VB.Label templab 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   435
            Index           =   5
            Left            =   4080
            TabIndex        =   281
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "电芯3 温度"
            Height          =   435
            Index           =   4
            Left            =   3360
            TabIndex        =   280
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label4 
            Caption         =   "电芯2 温度"
            Height          =   435
            Index           =   3
            Left            =   2520
            TabIndex        =   279
            Top             =   720
            Width           =   555
         End
         Begin VB.Label templab 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   278
            Top             =   240
            Width           =   495
         End
         Begin VB.Label templab 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   435
            Index           =   1
            Left            =   915
            TabIndex        =   277
            Top             =   240
            Width           =   495
         End
         Begin VB.Label templab 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   435
            Index           =   2
            Left            =   1710
            TabIndex        =   276
            Top             =   240
            Width           =   495
         End
         Begin VB.Label templab 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   435
            Index           =   3
            Left            =   2505
            TabIndex        =   275
            Top             =   240
            Width           =   495
         End
         Begin VB.Label templab 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   435
            Index           =   4
            Left            =   3300
            TabIndex        =   274
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "环境温度"
            Height          =   435
            Index           =   1
            Left            =   240
            TabIndex        =   273
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "MOS 温度"
            Height          =   435
            Index           =   0
            Left            =   960
            TabIndex        =   272
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label4 
            Caption         =   "电芯1 温度"
            Height          =   435
            Index           =   2
            Left            =   1740
            TabIndex        =   271
            Top             =   720
            Width           =   555
         End
      End
      Begin VB.CommandButton CmdSYSWrite 
         Caption         =   "下发配置"
         Height          =   435
         Left            =   5880
         TabIndex        =   36
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton CmdSysSaveclink 
         Caption         =   "保存配置"
         Height          =   435
         Left            =   2460
         TabIndex        =   35
         Top             =   3420
         Width           =   1335
      End
      Begin VB.CommandButton CmdSYSRead 
         Caption         =   "读取配置"
         Height          =   435
         Left            =   4140
         TabIndex        =   34
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton CmdSYSjiazai 
         Caption         =   "加载配置"
         Height          =   435
         Left            =   1200
         TabIndex        =   33
         Top             =   3420
         Width           =   1155
      End
      Begin VB.TextBox Textsys 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   32
         Text            =   "01"
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label LabelSYS 
         BackColor       =   &H8000000E&
         Caption         =   "保护板地址"
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
         Left            =   -60
         TabIndex        =   31
         Top             =   420
         Width           =   1815
      End
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   435
      Left            =   120
      TabIndex        =   22
      Text            =   "记录"
      Top             =   9420
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "PACK状态"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   10155
      Left            =   10440
      TabIndex        =   13
      Top             =   0
      Width           =   915
      Begin VB.HScrollBar HScrollMain1 
         Height          =   315
         LargeChange     =   10
         Left            =   12720
         Max             =   100
         SmallChange     =   10
         TabIndex        =   25
         Top             =   5100
         Width           =   6915
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "状态标志"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   120
         TabIndex        =   15
         Top             =   5640
         Width           =   21495
         Begin VB.Frame Frame3 
            Caption         =   "AFEReg状态"
            ForeColor       =   &H000080FF&
            Height          =   5235
            Index           =   5
            Left            =   17220
            TabIndex        =   268
            Top             =   2160
            Width           =   3000
            Begin VB.Label LabelBitG 
               BackColor       =   &H80000014&
               Caption         =   "单体过压"
               BeginProperty Font 
                  Name            =   "楷体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000C&
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   269
               Top             =   300
               Width           =   2955
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "电压状态"
            ForeColor       =   &H000080FF&
            Height          =   4155
            Index           =   0
            Left            =   1140
            TabIndex        =   20
            Top             =   600
            Width           =   3000
            Begin VB.Label LabelBitV 
               BackColor       =   &H80000014&
               Caption         =   "单体过压"
               BeginProperty Font 
                  Name            =   "楷体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF8080&
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   21
               Top             =   240
               Width           =   2955
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "保护板状态"
            ForeColor       =   &H000080FF&
            Height          =   4155
            Index           =   3
            Left            =   12120
            TabIndex        =   19
            Top             =   300
            Width           =   3000
            Begin VB.Label LabelBitA 
               BackColor       =   &H80000014&
               Caption         =   "单体过压"
               BeginProperty Font 
                  Name            =   "楷体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000C&
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   28
               Top             =   360
               Width           =   2955
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "温度状态"
            ForeColor       =   &H000080FF&
            Height          =   4155
            Index           =   2
            Left            =   8340
            TabIndex        =   18
            Top             =   480
            Width           =   3000
            Begin VB.Label LabelBitT 
               BackColor       =   &H80000014&
               Caption         =   "单体过压"
               BeginProperty Font 
                  Name            =   "楷体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000C&
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   300
               Width           =   2955
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "FET控制状态"
            ForeColor       =   &H000080FF&
            Height          =   5235
            Index           =   4
            Left            =   15300
            TabIndex        =   17
            Top             =   420
            Width           =   3000
            Begin VB.Label LabelBitF 
               BackColor       =   &H80000014&
               Caption         =   "单体过压"
               BeginProperty Font 
                  Name            =   "楷体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000C&
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   29
               Top             =   120
               Width           =   2955
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "电流状态"
            ForeColor       =   &H000080FF&
            Height          =   5955
            Index           =   1
            Left            =   4740
            TabIndex        =   16
            Top             =   480
            Width           =   3000
            Begin VB.Label LabelBitC 
               BackColor       =   &H80000014&
               Caption         =   "单体过压"
               BeginProperty Font 
                  Name            =   "楷体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   315
               Index           =   0
               Left            =   0
               TabIndex        =   26
               Top             =   300
               Width           =   2955
            End
         End
      End
      Begin VB.Label Labelpinheng 
         Alignment       =   2  'Center
         Caption         =   "均衡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   24
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Labelpinhenguse 
         BackColor       =   &H80000014&
         Caption         =   "√"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Index           =   0
         Left            =   4260
         TabIndex        =   23
         Top             =   420
         Width           =   375
      End
      Begin VB.Label LabelV82 
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
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   4095
      End
   End
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
      Height          =   360
      Index           =   0
      Left            =   18240
      TabIndex        =   11
      Text            =   "MCU-RTC-TIME"
      Top             =   11700
      Width           =   3015
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H000000FF&
      Caption         =   "CAN设置"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   0
      MaskColor       =   &H000000FF&
      TabIndex        =   7
      Top             =   8880
      Width           =   1395
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
      Top             =   7320
      Width           =   1496
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
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   5
      Top             =   4420
      Width           =   1496
   End
   Begin VB.CommandButton cmdMcuSysConfig 
      Caption         =   "系统参数1"
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
      Top             =   2360
      Width           =   1496
   End
   Begin VB.CommandButton cmdAfeReg 
      Caption         =   "硬件配置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Index           =   2
      Left            =   0
      MaskColor       =   &H00404040&
      TabIndex        =   3
      Top             =   1270
      Width           =   1496
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
      Width           =   1496
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出程序"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   10080
      Width           =   1397
   End
   Begin VB.CommandButton cmdCfgSPort 
      BackColor       =   &H00008000&
      Caption         =   "串口配置"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   60
      MaskColor       =   &H000040C0&
      TabIndex        =   0
      Top             =   8160
      Width           =   1365
   End
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   1320
      Top             =   10980
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   180
      Top             =   10980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   8520
      TabIndex        =   310
      Top             =   11640
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
      Max             =   5
   End
   Begin VB.Label tongxunerror 
      Caption         =   "0/0"
      Height          =   375
      Left            =   21360
      TabIndex        =   283
      Top             =   11640
      Width           =   1695
   End
   Begin VB.Label Labeljilu 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   960
      TabIndex        =   263
      Top             =   9420
      Width           =   495
   End
   Begin VB.Label Label_strdis 
      Alignment       =   2  'Center
      Caption         =   "未打开端口"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   345
      Left            =   3420
      TabIndex        =   12
      Top             =   11700
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   347
      Index           =   3
      Left            =   7260
      TabIndex        =   10
      Top             =   11700
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "主机地址"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   347
      Index           =   2
      Left            =   6360
      TabIndex        =   9
      Top             =   11700
      Width           =   855
   End
   Begin VB.Label Label_dis1 
      Alignment       =   2  'Center
      Caption         =   "未能连接下位机"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   347
      Left            =   60
      TabIndex        =   8
      Top             =   11700
      Width           =   3255
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
      Caption         =   "操作(&E)"
      Begin VB.Menu MNU_Edit_Edit 
         Caption         =   "FET手动开关(&F)"
         Begin VB.Menu MNU_Edit_cls_CDMOSFET 
            Caption         =   "关闭充放电及预充MOSFET"
         End
         Begin VB.Menu MNU_Edit_op_DMOSFET 
            Caption         =   "打开放电MOSFET"
         End
         Begin VB.Menu MNU_Edit_op_CMOSFET 
            Caption         =   "打开充电MOSFET"
         End
         Begin VB.Menu MNU_Edit_op_YMOSFET 
            Caption         =   "打开预充MOSFET"
         End
         Begin VB.Menu MNU_Edit_op_CDMOSFET 
            Caption         =   "打开充放电MOSFET"
         End
         Begin VB.Menu MNU_Edit_op_close_manual 
            Caption         =   "关闭手动____恢复自动控制"
         End
      End
      Begin VB.Menu blue_name 
         Caption         =   "修改蓝牙模块名称"
      End
      Begin VB.Menu moni_weixin 
         Caption         =   "模拟微信小程序"
         Begin VB.Menu moni_weixin_Read_soc 
            Caption         =   "读SOC"
         End
         Begin VB.Menu moni_weixin_Read_volt_curr 
            Caption         =   "读总压总电流"
         End
      End
      Begin VB.Menu ReSet_MCU 
         Caption         =   "复位MCU"
      End
      Begin VB.Menu BMS_POWEROFF 
         Caption         =   "BMS关机"
      End
      Begin VB.Menu BMS_POWERON 
         Caption         =   "BMS开机"
      End
   End
   Begin VB.Menu MNU_Option 
      Caption         =   "记录(&R)"
      Begin VB.Menu setRecordtime 
         Caption         =   "设置记录时间"
      End
      Begin VB.Menu MNU_RecordNow 
         Caption         =   "记录(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu MNU_Config 
      Caption         =   "设置(&C)"
      Begin VB.Menu MNU_Config_Port 
         Caption         =   "设置串口(&P)"
      End
      Begin VB.Menu setCyTime 
         Caption         =   "设置采样周期"
      End
      Begin VB.Menu MNU_Config_Code 
         Caption         =   "程序升级(&C)"
      End
      Begin VB.Menu JIHOU_BMS 
         Caption         =   "进入生产激活模式"
      End
      Begin VB.Menu admin 
         Caption         =   "管理员设置"
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
Public Function Frame_Visible_off()
    Frame_SYSCONFIG.Visible = False
    FrameREG.Visible = False
    Frame2.Visible = False
    Frame_Calib.Visible = False
    Frame_Record.Visible = False
    FrameSOC_OCV.Visible = False
    Framecap.Visible = False
    Frame_SYSMUC2.Visible = False
End Function

Private Sub admin_Click()
    Dim aa As String


    aa = InputBox("输入打开BMS激活模式密码")
    If aa = "16874162" Then
        BMS_admin_mode = 1
    ElseIf aa = "25649813" Then
        BMS_admin_mode = 2
    ElseIf aa = "36546123" Then
        BMS_admin_mode = 3
    ElseIf aa = "44654658" Then
        BMS_admin_mode = 4
    ElseIf aa = "52342342" Then
        BMS_admin_mode = 5
    ElseIf aa = "66546546" Then
        BMS_admin_mode = 6
    ElseIf aa = "75212354" Then
        BMS_admin_mode = 7
    Else
        My_msgbox "密码错误，请重新输入密码！"
    End If
 
End Sub

Private Sub blue_name_Click()
    Dialog_bluetooth.Visible = True
End Sub

Private Sub BMS_POWEROFF_Click() '按键下发 BMS 关机命令
    NextSentCmd = CMD_Enter_Sleep_Mode
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub BMS_POWERON_Click()
    NextSentCmd = V82_SET_POWERON
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub Command6_cap_Click()

End Sub





Private Sub CAP下发_Click()
 
If mode_bit5 = 22 Then
  mode_bit5 = 0
   CAP下发.Caption = "√CAP下发"
   CAP下发.ForeColor = &HC0&
Else
  mode_bit5 = 22
   CAP下发.Caption = "×CAP下发"
   CAP下发.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Command_reand_cap_Click()
    Call clear_Thecap
    Delay_dis_Readcap = 4  ' 延时处理 回复数据
    NextSentCmd = CMD_Readcap
    CMD_cmd_No = 0
    Record_Num = 0
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command_resetall_Click()
    NextSentCmd = CMD_ReSet_OFFSET
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub Command_writeCAP_Click()
    If (BMS_admin_mode = 4) Or (BMS_admin_mode = 7) Then
        Delay_dis_Writecap = 4 ' 延时处理 回复数据
        NextSentCmd = CMD_Writecap
        CMD_cmd_No = 1
        Record_Num = 0
        manual_time = 5 ' 500ms 发送间隔
    End If
End Sub

Private Sub Command3_Click()
自动配置结果.Text = ""
End Sub

Private Sub Command3CAP_Click()
If (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
    Framecap.Visible = True
    FrameSOC_OCV.Visible = False
End If
End Sub

Private Sub Command4_Click()
        BMS_active_mode = 0
        生产解码.Visible = False
       jingdu1 = 0
End Sub

Private Sub Commandopencap_Click()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s, kkk    As Integer
    Dim slen As Integer
    Dim bith, bitl  As Integer
    Dim rst(8) As Byte
    ' show open box
      CommonDialog_cap.ShowOpen
      filelocation = CommonDialog_cap.FileName
    ' input files into text1.text
    If CommonDialog_cap.FileName = "" Then
    Else
        Open filelocation For Input As #1
          I = 0
          Do Until EOF(1)
            Line Input #1, strsss
            For j = 0 To 3
              slen = Len(strsss)
              kkk = InStr(strsss, "$")                   ' 找到第一个空格
              capData(I * 4 + j) = Val(Mid(strsss, 1, kkk - 1)) ' 读出空格前数据
              strsss = Mid(strsss, kkk + 1, slen)            ' 清出这个读出数据 继续
            Next j
            If I < 26 Then
                I = I + 1
            End If
           Loop
        Call PrintfThecap
        Close #1
    End If
End Sub

Private Sub CommandSavecap_Click()
    Dim I, j As Integer
    Dim filelocation As String
    ' loads save as box
      CommonDialog_cap.ShowSave
      filelocation = CommonDialog_cap.FileName
       If CommonDialog_cap.FileName = "" Then
       Else
       
    ' append saves over file if it assists
    '  Open filelocation For Append As #1
      Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的
      For I = 0 To 26
        For j = 0 To 3
            Print #1, str(capData(I * 4 + j)) & "$"; '& vbCrLf 有分号 不分行
        Next j
        Print #1, ""   ' 下一行
      Next I
      Close #1
      End If
End Sub

 Private Sub Form_Resize()
    Dim xx  As Long
    Dim yy As Long
    xx = Screen.Width / Screen.TwipsPerPixelX '获取屏幕bai宽度
    yy = Screen.Height / Screen.TwipsPerPixelY '获取屏幕高度
    If Me.WindowState = 2 Then
        Call form_allresize(xx, yy)
    Else
        Call form_allresize(xx, yy)
    End If
 End Sub
Private Sub cmdAfeReg_Click(Index As Integer)
    Call Frame_Visible_off
    If (BMS_admin_mode = 1) Or (BMS_admin_mode = 2) Or (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
        FrameREG.Visible = True
    End If
End Sub
' 初始化一个理论的 备份表
 Public Function load_backup()
    Dim I, j, n As Integer
    Dim iCol    As Integer
    MSFlexGrid1.Rows = 801
    MSFlexGrid1.Cols = 51
    MSFlexGrid1.Left = 0
    MSFlexGrid1.Top = 0
    MSFlexGrid1.Width = 19740
    MSFlexGrid1.Height = 11800 - 1000  ' 比整页少500 留部分放按键
    For j = 1 To 50
        MSFlexGrid1.ColWidth(j) = 600 ' 第一排列宽 设置窄一点
    Next
    MSFlexGrid1.FillStyle = flexFillRepeat '把更改应用到所有选定单元
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.ColSel = MSFlexGrid1.Cols - 1 '和col属性联用选中第1到第5列,
    MSFlexGrid1.CellFontSize = 8 '经过fillstyle设置效果应用到选定单元
    MSFlexGrid1.AllowUserResizing = flexResizeColumns '用户可以 修改整列大小
    MSFlexGrid1.ScrollBars = flexScrollBarBoth  ' MSHFlexGrid 有水平和竖直的滚动条。这是缺省设置值
    MSFlexGrid1.ColWidth(0) = 400 ' 第一排列宽 设置窄一点
    MSFlexGrid1.ColWidth(2) = 1500 '时间 要宽一些
    MSFlexGrid1.ColWidth(3) = 1500 '时间 要宽一些
   ' MSFlexGrid1.ColWidth = 500
    MSFlexGrid1.TextMatrix(0, 0) = "Num"
    MSFlexGrid1.TextMatrix(0, 1) = "Time"
    MSFlexGrid1.ColWidth(1) = 1800 '时间 要宽一些
    MSFlexGrid1.TextMatrix(0, 2) = "Event"
    MSFlexGrid1.TextMatrix(0, 3) = "Protect_state"
    MSFlexGrid1.TextMatrix(0, 4) = "PackStatus"
    MSFlexGrid1.TextMatrix(0, 5) = "BatteryStatus"
    MSFlexGrid1.TextMatrix(0, 6) = "FCC(mAh)"
    MSFlexGrid1.TextMatrix(0, 7) = "RC(mAh)"
    MSFlexGrid1.TextMatrix(0, 8) = "SOC‰"
    MSFlexGrid1.TextMatrix(0, 9) = "Total(V)"
    MSFlexGrid1.TextMatrix(0, 10) = "Current(mA)"
    MSFlexGrid1.TextMatrix(0, 11) = "Cell1(mV)"
    MSFlexGrid1.TextMatrix(0, 12) = "Cell2(mV)"
    MSFlexGrid1.TextMatrix(0, 13) = "Cell3(mV)"
    MSFlexGrid1.TextMatrix(0, 14) = "Cell4(mV)"
    MSFlexGrid1.TextMatrix(0, 15) = "Cell5(mV)"
    MSFlexGrid1.TextMatrix(0, 16) = "Cell6(mV)"
    MSFlexGrid1.TextMatrix(0, 17) = "Cell7(mV)"
    MSFlexGrid1.TextMatrix(0, 18) = "Cell8(mV)"
    MSFlexGrid1.TextMatrix(0, 19) = "Cell9(mV)"
    MSFlexGrid1.TextMatrix(0, 20) = "Cell10(mV)"
    MSFlexGrid1.TextMatrix(0, 21) = "Cell11(mV)"
    MSFlexGrid1.TextMatrix(0, 22) = "Cell12(mV)"
    MSFlexGrid1.TextMatrix(0, 23) = "Cell13(mV)"
    MSFlexGrid1.TextMatrix(0, 24) = "Cell14(mV)"
    MSFlexGrid1.TextMatrix(0, 25) = "Cell15(mV)"
    MSFlexGrid1.TextMatrix(0, 26) = "Cell16(mV)"
    MSFlexGrid1.TextMatrix(0, 27) = "AmbientTemp"
    MSFlexGrid1.TextMatrix(0, 28) = "PowerTemp"
    MSFlexGrid1.TextMatrix(0, 29) = "Cell1Temp"
    MSFlexGrid1.TextMatrix(0, 30) = "Cell2Temp"
    MSFlexGrid1.TextMatrix(0, 31) = "Cell3Temp"
    MSFlexGrid1.TextMatrix(0, 32) = "Cell4Temp"
    MSFlexGrid1.TextMatrix(0, 33) = "Cell5Temp"
    MSFlexGrid1.TextMatrix(0, 34) = "Cell6Temp"
    MSFlexGrid1.TextMatrix(0, 35) = "Cell7Temp"
    For I = 1 To 800
        MSFlexGrid1.TextMatrix(I, 0) = I
    Next I
 End Function
  ' 加载时，禁止AFE 那温度参数变化
 Public Function AFEreg_temp_text_nodo()
'    Text1(8).Enabled = False
'    Text1(9).Enabled = False
'    Text1(10).Enabled = False
'    Text1(11).Enabled = False
'    Text1(12).Enabled = False
'    Text1(13).Enabled = False
'    Text1(14).Enabled = False
'    Text1(15).Enabled = False
 End Function
 ' 初始化一个理论的 备份表
 Public Function load_SOC_COVGrid()
    Dim I   As Integer
    MSFlexGrid2.Rows = 24
    MSFlexGrid2.Cols = 10
    MSFlexGrid2.Left = 300
    MSFlexGrid2.Top = 480
    MSFlexGrid2.Width = 18555
    MSFlexGrid2.Height = 7455 ' 比整页少500 留部分放按键
    MSFlexGrid2.AllowUserResizing = flexResizeColumns '用户可以 修改整列大小
    MSFlexGrid2.ScrollBars = flexScrollBarBoth  ' MSHFlexGrid 有水平和竖直的滚动条。这是缺省设置值
    MSFlexGrid2.ColWidth(0) = 400 ' 第一排列宽 设置窄一点
    For I = 0 To 5
         MSFlexGrid2.ColWidth(I) = 2200
    Next I
 End Function
  ' 初始化一个理论的 备份表
 Public Function load_capGrid()
    Dim I   As Integer
    MSFlexGridcap.Rows = 30
    MSFlexGridcap.Cols = 10
    MSFlexGridcap.Left = 300
    MSFlexGridcap.Top = 480
    MSFlexGridcap.Width = 18555
    MSFlexGridcap.Height = 9455 ' 比整页少500 留部分放按键
    MSFlexGridcap.AllowUserResizing = flexResizeColumns '用户可以 修改整列大小
    MSFlexGridcap.ScrollBars = flexScrollBarBoth  ' MSHFlexGrid 有水平和竖直的滚动条。这是缺省设置值
    MSFlexGridcap.ColWidth(0) = 400 ' 第一排列宽 设置窄一点
    For I = 0 To 5
         MSFlexGridcap.ColWidth(I) = 2200
    Next I
 End Function
 
Private Sub cmdBackup_Click()
    Dim I, j  As Integer
    Call Frame_Visible_off
      If (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
        Frame_Record.Visible = True
    End If
End Sub
Private Sub cmdBQSYS_Click()
    Call Frame_Visible_off
    If (BMS_admin_mode = 2) Or (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 6) Or (BMS_admin_mode = 7) Then
        Frame_SYSMUC2.Visible = True
    End If
End Sub
Private Sub cmdCalibration_Click(Index As Integer)
    Call Frame_Visible_off
     If (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
    Frame_Calib.Visible = True
    End If
End Sub
Private Sub cmdCfgSPort_Click(Index As Integer)
    chuanshu.Show 1, frmMain
End Sub
Private Sub cmdConnect_Click(Index As Integer)
    Form1_can.Show 1, frmMain
End Sub

Private Sub cmdExit_Click()
End
    chuanshu.Enabled = False
    
      VCI_CloseDevice m_devtype, 0
    
    Unload chuanshu
    Unload Me
End Sub

Private Sub cmdMcuSysConfig_Click(Index As Integer)
    Call Frame_Visible_off
    If (BMS_admin_mode = 2) Or (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 6) Or (BMS_admin_mode = 7) Then
        Frame_SYSCONFIG.Visible = True
    End If
End Sub

Private Sub cmdPackInfo_Click(Index As Integer)
    Call Frame_Visible_off
    Frame2.Visible = True
    manual_time = cyInfoTime
End Sub

 ' 系统参数 界面 改变显示位置
Private Sub FrameSYSconfig_Load()
    Dim I As Byte
    Dim j, y1, x1, x2, x3 As Integer
    Dim float As Single
    Dim zhens, xiaos As Long
    Dim huadong As Integer
    For I = 1 To 100
        LabelSYS(I).Visible = False
        Textsys(I).Visible = False
    Next
    Call frist_sys1load_ptch
    For I = 0 To 72 - 1
        LabelSYS(I).Visible = True
        LabelSYS(I).Height = 375
        LabelSYS(I).Width = 3600
        LabelSYS(I).Left = huadong + 60 + 6800 * Fix(I / 25)
        LabelSYS(I).Top = 360 + (I - Fix(I / 25) * 25) * (375 + 45)
        LabelSYS(I).Caption = sysCaption(I)
        LabelSYS(I).Visible = True
        Textsys(I).ForeColor = &H40C0&
        Textsys(I).Text = "---"
        Textsys(I).Visible = True
        Textsys(I).Height = 375
        Textsys(I).Width = 2050
        Textsys(I).Left = huadong + 60 + 6800 * Fix(I / 25) + 3700
        Textsys(I).Top = 360 + (I - Fix(I / 25) * 25) * (375 + 45)
    Next
End Sub
 ' 系统参数 界面 改变显示位置
Private Sub FrameSYS2config_Load()
    Dim I As Byte
    Dim j, y1, x1, x2, x3 As Integer
    Dim float As Single
    Dim zhens, xiaos As Long
    Dim huadong As Integer
    For I = 1 To 100
        LabeSYS2(I).Visible = False
        TexSys2(I).Visible = False
    Next
    Call frist_sys2load_ptch
    For I = 0 To 24 - 1
        LabeSYS2(I).Visible = True
        LabeSYS2(I).Height = 375
        LabeSYS2(I).Width = 3600
        LabeSYS2(I).Left = 60 + 6800 * Fix(I / 25)
        LabeSYS2(I).Top = 360 + (I - Fix(I / 25) * 25) * (375 + 45)
        LabeSYS2(I).Caption = sys2Caption(I)
        LabeSYS2(I).Visible = True
        TexSys2(I).ForeColor = &H40C0&
        TexSys2(I).Text = "---"
        TexSys2(I).Visible = True
        TexSys2(I).Height = 375
        TexSys2(I).Width = 4100
        TexSys2(I).Left = 60 + 6800 * Fix(I / 25) + 3700
        TexSys2(I).Top = 360 + (I - Fix(I / 25) * 25) * (375 + 45)
    Next
End Sub
Private Sub Frame2_Load()
    Dim I As Byte
    Dim j, y1, x1, x2, x3 As Integer
    Dim float As Single
    Dim zhens, xiaos As Long
    Dim huadong As Integer
    Dim strtemp As String
    huadong = HScrollMain1.Value * (-150)
    LOAD_CELLmun = myRealV82Info.Vcell_num
    LOAD_Tempmun = myRealV82Info.RealTempNum
    For I = 1 To 100
        LabelV82(I).Visible = False
    Next I
    For I = 1 To 50
        Labelpinhenguse(I).Visible = False
    Next I
    If myRealV82Info.Vcell_num > 0 Then
    Else
        myRealV82Info.Vcell_num = 1
    End If
    For I = 0 To myRealV82Info.Vcell_num - 1
        LabelV82(I).Visible = True
        LabelV82(I).Height = 375
        LabelV82(I).Width = 4095
        LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
        LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
        If (I) < 9 Then
            LabelV82(I).Caption = "电芯0" & (I + 1) & "：" & myRealV82Info.Vcell(I) & "   V"
        Else
            LabelV82(I).Caption = "电芯" & (I + 1) & "：" & myRealV82Info.Vcell(I) & "   V"
        End If
        Labelpinhenguse(I).Visible = True
        Labelpinhenguse(I).Left = huadong + 4260 + 4800 * Fix(I / 10)
        Labelpinhenguse(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
        xiaos = 0
        Labelpinhenguse(I).Caption = " "
    Next I
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "电池总电压：" & myRealV82Info.Vbat & "   V"
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "实时电流 ：" & myRealV82Info.Curr & "   A"
 
 
    For j = 0 To 4
        If myRealV82Info.RealTempNum And (2 ^ j) Then
            I = I + 1
            LabelV82(I).Visible = True
            LabelV82(I).Height = 375
            LabelV82(I).Width = 4095
            LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
            LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
            LabelV82(I).Caption = " 温度 " & j + 1 & "  ： " & myRealV82Info.temp(j) & "   °C"
        End If
    Next j
    
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    xiaos = myRealV82Info.NUM_VOV + 1 ' '单体高压对应的电池的序号，例如 5 表示第 5 节高压
    LabelV82(I).Caption = "最高电芯序号：   " & xiaos
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    xiaos = myRealV82Info.NUM_VUV + 1 ' '单体高压对应的电池的序号，例如 5 表示第 5 节高压
    LabelV82(I).Caption = "最低电芯序号：   " & xiaos
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    xiaos = myRealV82Info.NUM_WARN_VHIGH  ' '单体高压警告对应的电池的序号
    LabelV82(I).Caption = "最高电芯电压：   " & xiaos
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    xiaos = myRealV82Info.NUM_WARN_VLOW  ' '单体低压警告对应的电池的序号
    LabelV82(I).Caption = "最低电芯电压：   " & xiaos
    I = I + 1
    y1 = I
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    xiaos = myRealV82Info.DchgNum  '：放电次数'
    LabelV82(I).Caption = "累计放电次数：   " & xiaos
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    xiaos = myRealV82Info.mcu_powerStatu  '：放电次数'
          strtemp = ""
        
         If xiaos >= 64 Then
             xiaos = xiaos - 64
            strtemp = "MCU状态：激活  "
         Else
            strtemp = "MCU状态：待激活"
         End If
         
         If xiaos >= 32 Then
             xiaos = xiaos - 32
            strtemp = strtemp & ""
         Else
            strtemp = strtemp & "未配置"
         End If
         
         If xiaos >= 16 Then
             xiaos = xiaos - 16
            strtemp = strtemp & ""
         Else
            strtemp = strtemp & "未开机"
         End If
         
        If xiaos = 0 Then
            LabelV82(I).Caption = strtemp & "上电未初始化"
        End If
        If xiaos = 1 Then
            LabelV82(I).Caption = strtemp & "工作中"
        End If
        If xiaos = 2 Then
            LabelV82(I).Caption = strtemp & "进入休眠"
        End If
        If xiaos = 3 Then
            LabelV82(I).Caption = strtemp & "休眠"
        End If
        If xiaos = 4 Then
            LabelV82(I).Caption = strtemp & "关机中"
        End If
            If xiaos = 5 Then
        LabelV82(I).Caption = strtemp & "关机"
        End If

    ' BlanceState As Long ' ： 均衡状态，表示那一节电压开启均衡
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "剩余电量百分比SOC：   " & myRealV82Info.SOC & "  %"  '电池 soc ，百分比 0-1000
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    '：放电次数'
    LabelV82(I).Caption = "当前剩余容量：   " & myRealV82Info.CapNow & "  AH"   ' 当前容量 (0.1AH)
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "满充容量：   " & myRealV82Info.CapFull & "  AH"  '满充容量(0.1AH)
    
     I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "run_code：   " & Int_to_hex(myRealV82Info.FET_code)   '
    
     I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "AFE_TEMP1：   " & myRealV82Info.afe_Temp(1) & "   °C"
     I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "AFE_TEMP2：   " & myRealV82Info.afe_Temp(2) & "   °C"
    I = I + 1
    LabelV82(I).Visible = True
    LabelV82(I).Height = 375
    LabelV82(I).Width = 4095
    LabelV82(I).Left = huadong + 60 + 4800 * Fix(I / 10)
    LabelV82(I).Top = 360 + (I - Fix(I / 10) * 10) * (375 + 45)
    LabelV82(I).Caption = "AFE_TEMP3：   " & myRealV82Info.afe_Temp(3) & "   °C"
    
    
End Sub
Private Sub CmdRegjiazai_Click()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s, kkk As Integer
    Dim bith, bitl  As Integer
    Dim rst(8) As Byte
    ' show open box
      CommonDialog1.ShowOpen
      filelocation = CommonDialog1.FileName
    ' input files into text1.text
    If CommonDialog1.FileName = "" Then
    Else
        Open filelocation For Input As #1
        For I = 0 To 25
            Line Input #1, strsss
            strsss = Replace(strsss, " ", "")
            RegEERPOM(I) = Mid(strsss, 1, 2)
        Next I
        Call PrintfTheReg
        Close #1
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         手动读取AFE寄存器按钮函数
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CmdRegReadout_Click()
    Delay_dis_ReadRegAfe = 4  ' 延时处理 回复数据
    NextSentCmd = CMD_ReadAFEseg
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub CmdRegSave_Click()
    Dim I As Integer
    Dim filelocation As String
    ' loads save as box
    CommonDialog1.ShowSave
    filelocation = CommonDialog1.FileName
    If CommonDialog1.FileName = "" Then
    Else
        ' append saves over file if it assists
        '  Open filelocation For Append As #1
        Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的
        Call frmMain.ReadTheRegchang
        For I = 0 To 25
            Print #1, RegEERPOM(I)
        Next I
        Close #1
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         手动下发设置AFE寄存器按钮函数
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CmdRegSent_Click()
    If (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
        Delay_dis_WriteRegAfe = 44
        NextSentCmd = CMD_WriteAFEseg
        Delay_waite_muc_back_cmd = 20
        manual_time = 12 ' 500ms 发送间隔
    End If

End Sub

Private Sub cmdSOC_Click()
    Call Frame_Visible_off
     If (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
       
         FrameSOC_OCV.Visible = True
        Framecap.Visible = False
      

       
        
    End If
End Sub
Private Sub cmdcap_Click()
    Call Frame_Visible_off
     If (BMS_admin_mode = 3) Or (BMS_admin_mode = 4) Or (BMS_admin_mode = 5) Or (BMS_admin_mode = 7) Then
        Framecap.Visible = True
    End If
End Sub
Private Sub CmdSYSjiazai_Click()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s As Integer
    ' show open box
    CommonDialog1.ShowOpen
    filelocation = CommonDialog1.FileName
    If CommonDialog1.FileName = "" Then
    Else
        ' input files into text1.text
        Open filelocation For Input As #1
        Do Until EOF(1)
        Line Input #1, strsss
        strsss = Replace(strsss, " ", "")
        s = Len(strsss)
        j = InStr(strsss, "=")
        sysCaption(I) = Mid(strsss, 1, j - 1)
        LabelSYS(I).Caption = sysCaption(I)
        Textsys(I) = Mid(strsss, j + 1, s)
        I = I + 1
        Loop
        Close #1
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         手动读取V82 MCU 设置参数 按钮函数
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CmdSYSRead_Click()
    Delay_dis_Readsysconfig = 4  '2S 等待
    NextSentCmd = CMD_ReadSysConfig
    manual_time = 5 ' 500ms 发送间隔
End Sub
' 先读取 界面 txt值 到变量中
Public Sub CmdSYSWrite_Click()
Dim temp As Long
Dim I As Integer
I = I + 0:    McuV82SysConfig.Addr = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CellNum = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.TempsetNum = Val("&H" & Textsys(I))
I = I + 1:    McuV82SysConfig.EngDesign = Val(Textsys(I)) * 10
I = I + 1:    McuV82SysConfig.BalanceCur = Val(Textsys(I))   '  //"均衡启动最小充电电流(mA)"    原来这个不用    采样电阻大小    0_01mR）
I = I + 1:    McuV82SysConfig.BalanceDelay = Val(Textsys(I))   '    //  软件均衡延时(S）    原来这个不用    参考电压    mv  10
I = I + 1:
If Textsys(I) = "不均衡" Then
    McuV82SysConfig.B_Mode = 0
End If
If Textsys(I) = "充电均衡" Then
    McuV82SysConfig.B_Mode = 1
End If
If Textsys(I) = "充电+静态均衡" Then
    McuV82SysConfig.B_Mode = 2
End If
I = I + 1:    McuV82SysConfig.B_THDIS = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.B_TLDIS = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.B_VStart = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.B_Vdiff = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.W_Vcell_H = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.W_VCell_L = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.W_VBAT_H = Val(Textsys(I)) ' * McuV82SysConfig.CellNum / 2
I = I + 1:    McuV82SysConfig.W_VBAT_L = Val(Textsys(I)) ' * McuV82SysConfig.CellNum / 2
I = I + 1:    McuV82SysConfig.W_Tcell_H = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.W_Tcell_L = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.W_Tenv_H = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.W_Tenv_L = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.W_Tfet_H = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.W_Tfet_L = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.W_CURR_C = Val(Textsys(I)) * 100
I = I + 1:    McuV82SysConfig.W_CURR_D = Val(Textsys(I)) * 100
I = I + 1:    McuV82SysConfig.W_VDIFF_H = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.W_VDIFF_L = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.OVPVal = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.OVPDly = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.OVPRel = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.UVPVal = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.UVPDly = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.UVPRel = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.BOVPVal = Val(Textsys(I)) '* McuV82SysConfig.CellNum / 2
I = I + 1:    McuV82SysConfig.BOVPDly = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.BOVPRel = Val(Textsys(I)) '* McuV82SysConfig.CellNum / 2
I = I + 1:    McuV82SysConfig.BUVPVal = Val(Textsys(I)) '* McuV82SysConfig.CellNum / 2
I = I + 1:    McuV82SysConfig.BUVPDly = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.BUVPRel = Val(Textsys(I)) '* McuV82SysConfig.CellNum / 2
I = I + 1:    McuV82SysConfig.CTcellHPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.CTcellHRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.CTcellLPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.CTcellLRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.DTcellHPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.DTcellHRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.DTcellLPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.DTcellLRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TenvHPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TenvHRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TenvLPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TenvLRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TfetHPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TfetHRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TfetLPro = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.TfetLRel = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.CC_PRO_VAL = Val(Textsys(I)) * 100
I = I + 1:    McuV82SysConfig.CC_PRO_PDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CC_PRO_RDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CC_PRO_LOCK = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CD1_PRO_VAL = Val(Textsys(I)) * 100
I = I + 1:    McuV82SysConfig.CD1_PRO_PDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CD1_PRO_RDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CD1_PRO_LOCK = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CD2_PRO_VAL = Val(Textsys(I)) * 100
I = I + 1:    McuV82SysConfig.CD2_PRO_PDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CD2_PRO_RDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.CD2_PRO_LOCK = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.SHORT_VAL = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.SHORT_RDLY = Val(Textsys(I))
I = I + 1:    McuV82SysConfig.SHORT_LOCK = Val(Textsys(I))
I = I + 1:
If Textsys(I) = "使能" Then
       McuV82SysConfig.HEAT_EN = 1
Else
       McuV82SysConfig.HEAT_EN = 0
End If
I = I + 1:    McuV82SysConfig.HEAT_TSTART = Val(Textsys(I)) + 40
I = I + 1:    McuV82SysConfig.HEAT_TEND = Val(Textsys(I)) + 40

If (BMS_admin_mode = 4) Or (BMS_admin_mode = 6) Or (BMS_admin_mode = 7) Then
    Delay_dis_Writesysconfig = 10  '2S 等待
     NextSentCmd = CMD_WriteSysConfig
     manual_time = 5 ' 500ms 发送间隔
End If

End Sub
Private Sub Command_clearBlackup_Click()
    Dim I, j As Integer
    For I = 1 To 800
        For j = 1 To 50
            MSFlexGrid1.TextMatrix(I, j) = ""
        Next j
    Next I
End Sub
Private Sub Command_Currentk_Click()
    Dim strm As String
    If (BMS_admin_mode = 4) Or (BMS_admin_mode = 7) Then
        Delay_dis_CALIB_CURRENT = 4
        NextSentCmd = CMD_CALIB_CURRENT
        manual_time = 10 ' 500ms 发送间隔  ' 加上 放电 负值 校正，当有 负号时，最高位置1
        strm = RealCurrentkText.Text
        sent_result = strm * 1000
        CMD_cmd_No = 1
    End If
End Sub

Private Sub Command_Currentz_Click()
   If (BMS_admin_mode = 4) Or (BMS_admin_mode = 7) Then
    Delay_dis_CALIB_CURRENT = 4
    NextSentCmd = CMD_CALIB_CURRENT
    manual_time = 10 ' 500ms 发送间隔
    sent_result = RealCurrentzText.Text * 1000
    CMD_cmd_No = 2
    End If
End Sub
Private Sub Command_rtcTIME_Click()
    Delay_dis_CALIB_RTC = 4
    NextSentCmd = CMD_CALIB_RTC
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub Command_sys2load_Click()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s As Integer
    ' show open box
     CommonDialogsys2.ShowOpen
     
     filelocation = CommonDialogsys2.FileName
     If CommonDialogsys2.FileName = "" Then
     Else
    ' input files into text1.text
      Open filelocation For Input As #1
        Do Until EOF(1)
            Line Input #1, strsss
            strsss = Replace(strsss, " ", "")
            s = Len(strsss)
            j = InStr(strsss, "=")
            sys2Caption(I) = Mid(strsss, 1, j - 1)
             LabeSYS2(I).Caption = sys2Caption(I)
            TexSys2(I) = Mid(strsss, j + 1, s)
            I = I + 1
        Loop
      Close #1
   End If
End Sub
' 读取默认文件夹内的 SYSCONF2 文件
Private Sub frist_sys2load_ptch()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s As Integer
    ' show open box
    filelocation = App.Path + "\配置文件\系统参数2"
    Open filelocation For Input As #1 '打开文本文件zhi，读取。
 
        Do Until EOF(1)
            Line Input #1, strsss
            strsss = Replace(strsss, " ", "")
            s = Len(strsss)
            j = InStr(strsss, "=")
            sys2Caption(I) = Mid(strsss, 1, j - 1)
             LabeSYS2(I).Caption = sys2Caption(I)
          '  TexSys2(i) = Mid(strsss, j + 1, s)
            I = I + 1
        Loop
      Close #1
   
End Sub
' 读取默认文件夹内的 SYSCONF1 文件
Private Sub frist_sys1load_ptch()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s As Integer
    ' show open box
    Open App.Path + "\配置文件\系统参数1" For Input As #1 '打开文本文件zhi，读取。
        Do Until EOF(1)
            Line Input #1, strsss
            strsss = Replace(strsss, " ", "")
            s = Len(strsss)
            j = InStr(strsss, "=")
            sysCaption(I) = Mid(strsss, 1, j - 1)
              LabelSYS(I).Caption = sysCaption(I)
           ' Textsys(i) =
            I = I + 1
        Loop
      Close #1
   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'         手动读取V82 MCUsys2 设置参数 按钮函数
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command_Sys2Read_Click()
    Delay_dis_Readsys2config = 4  '2S 等待
    NextSentCmd = CMD_ReadSys2Config
    manual_time = 5 ' 500ms 发送间隔
End Sub
Private Sub Command_Sys2Save_Click()
   Dim I As Integer
    Dim filelocation As String
    ' loads save as box
    CommonDialogsys2.ShowSave
    filelocation = CommonDialogsys2.FileName
    If CommonDialogsys2.FileName = "" Then
    Else
    ' append saves over file if it assists
    '  Open filelocation For Append As #1
      Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的
      For I = 0 To 70
        Print #1, sys2Caption(I) & "           =            " & TexSys2(I) '& vbCrLf
      Next I
      Close #1
    End If
End Sub

'  读取 界面 txt值 到 SYS2CONFIG变量中
Public Sub Command_Sys2Write_Click()
Dim temp As Long
Dim I, j As Integer
Dim strrrr  As String
I = I + 0:    McuSys2Config.DesignVol = Val(TexSys2(I)) * 10
I = I + 1:    McuSys2Config.PackConfigMap = Val(TexSys2(I)) Mod 65536
I = I + 1:    McuSys2Config.FCC = Val(TexSys2(I)) * 10
I = I + 1:    McuSys2Config.CycleThreshold = Val(TexSys2(I)) * 10
I = I + 1:    McuSys2Config.CycleCount = Val(TexSys2(I))
I = I + 1:    McuSys2Config.NearFCC = Val(TexSys2(I))
I = I + 1:    McuSys2Config.DfilterCur = Val(TexSys2(I))
I = I + 1:    McuSys2Config.LearnLowTemp = Val(TexSys2(I))
I = I + 1:    McuSys2Config.SWVersion = Val(TexSys2(I)) * 100
I = I + 1:    McuSys2Config.HWVersion = Val(TexSys2(I)) * 100
I = I + 1:    McuSys2Config.ShutDownDelay = Val(TexSys2(I))
I = I + 1:    McuSys2Config.SelfDsgRate = Val(TexSys2(I)) Mod 256
I = I + 1:    McuSys2Config.CommOffDelay = Val(TexSys2(I)) Mod 256
I = I + 1
strrrr = Mid(TexSys2(I), 1, 8)
For j = Len(strrrr) To 7  ' 不足8位 补空格
 strrrr = strrrr & " "
Next j
I = I + 1: McuSys2Config.MNFDate = strrrr
strrrr = Mid(TexSys2(I), 1, 16)
For j = Len(strrrr) To 16
    strrrr = strrrr & " "
Next j
I = I + 1:    McuSys2Config.MNFName = strrrr
strrrr = Mid(TexSys2(I), 1, 16)
For j = Len(strrrr) To 16
 strrrr = strrrr & " "
Next j
I = I + 1:    McuSys2Config.DeviceName = strrrr
strrrr = Mid(TexSys2(I), 1, 16)
For j = Len(strrrr) To 16
 strrrr = strrrr & " "
Next j
McuSys2Config.SN = strrrr

I = I + 1: McuSys2Config.SOH = Val(TexSys2(I))
I = I + 1
strrrr = Mid(TexSys2(I), 1, 16)
I = I + 1:
McuSys2Config.MCU_ID = strrrr
strrrr = TexSys2(I)
If jingdu1 = 4 Then
Else

 McuSys2Config.KEY_CODE = strrrr
End If

    If (BMS_admin_mode = 4) Or (BMS_admin_mode = 6) Or (BMS_admin_mode = 7) Then
        Delay_dis_Writesys2config = 4  '2S 等待
        NextSentCmd = CMD_WriteSys2Config  '2023.1.23  改CMD_WriteSys2Config
        manual_time = 5 ' 500ms 发送间隔
    End If
End Sub

Private Sub Command_Temp_Click()
Dim I As Integer
   If (BMS_admin_mode = 4) Or (BMS_admin_mode = 7) Then
    For I = 0 To 8
        If LabelwenduCalib(I).Caption = "√" Then
            Claib_temp(I) = 1
        Else
            Claib_temp(I) = 0
        End If
    Next I
    Delay_dis_CALIB_Temp = 4
    NextSentCmd = CMD_CALIB_TEMPE
    manual_time = 5 ' 500ms 发送间隔
    ' sent_result = RealVolteText.Text(i) * 1000
     CMD_cmd_No = 0
     End If
End Sub

Private Sub Command1_Click()
 Dim I, j As Integer
 For I = 1 To 800
  For j = 1 To 50
  MSFlexGrid1.TextMatrix(I, j) = ""
  Next j
 Next I

   
    
    puse_blackup_button = 0
    Delay_dis_ReadBalckUp = 20
    NextSentCmd = CMD_ReadBalckUp
    manual_time = 3 ' 500ms 发送间隔
    CMD_cmd_No = 0 ' 读第一条
    Record_Num = 0
End Sub

Private Sub Command1_Volte_Click()
   ' Delay_dis_CALIB_VOLTAGE = 4
   ' NextSentCmd = CMD_CALIB_VOLTAGE
   ' manual_time = 5 ' 500ms 发送间隔
    'sent_result = RealVolteText.Text * 1000
End Sub

Private Sub Command2_Click()
     puse_blackup_button = 1
     Delay_dis_ReadBalckUp = 0
     manual_time = 0
     NextSentCmd = CMD_ReadInfo  '接收正确 后立马换读数据，不然会出现 再次读最后一条记录
End Sub

Private Sub Command_EraseBalckUp_Click()
If (BMS_admin_mode = 7) Then
    Delay_dis_EraseBalckUp = 4
    NextSentCmd = CMD_EraseBalckUp
    manual_time = 5 ' 500ms 发送间隔
End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  保存 读取到的备份数据  到DATA
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command_SaveBlackupData_Click()
    Dim I, j As Integer
    Dim filelocation As String
    Dim strrg  As String
    ' loads save as box
      CommonDialog_Record.ShowSave
      filelocation = CommonDialog_Record.FileName
       If CommonDialog_Record.FileName = "" Then
       Else
       
    ' append saves over file if it assists
    '  Open filelocation For Append As #1
      Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的
      For I = 0 To Record_Num
        strrg = ""
        For j = 0 To 49
            strrg = strrg & "  " & MSFlexGrid1.TextMatrix(I, j)
        Next j
        Print #1, strrg
      Next I
      Close #1
      End If
      
End Sub

Public Function PrintfTheSOCOCV() ' 显示 读取到的寄存器
    Dim I As Integer
    Dim j, s, kkk As Long
    Dim bith, bitl  As Long
    Dim rst(8) As Byte
    For I = 0 To 21
      For j = 0 To 5
        MSFlexGrid2.TextMatrix(I, j) = SOC_OCVData(I, j)
      Next j
    Next I
End Function
Public Function PrintfThecap() ' 显示 读取到的寄存器
    Dim I As Integer
    Dim j, s, kkk As Long
    Dim bith, bitl  As Long
    Dim rst(8) As Byte
    For I = 1 To 26
      For j = 1 To 3
        MSFlexGridcap.TextMatrix(I, j) = capData(I * 4 + j)
      Next j
    Next I
    
      MSFlexGridcap.TextMatrix(0, 0) = "\"
        MSFlexGridcap.TextMatrix(0, 1) = "InMax" & capData(0 * 4 + 1) & "V"
          MSFlexGridcap.TextMatrix(0, 2) = "InMax" & capData(0 * 4 + 2) & "A"
            MSFlexGridcap.TextMatrix(0, 3) = "OutMax" & capData(0 * 4 + 3) & "A"
    For I = 1 To 10
        MSFlexGridcap.TextMatrix(I, 0) = "~" & (capData(I * 4 + 0) - 40) & "℃"
    Next I
    For I = 18 To 26
        MSFlexGridcap.TextMatrix(I, 0) = "~" & capData(I * 4 + 0) * 20 & "mV"
    Next I
End Function
Public Function clear_TheSOCOCV() '  清除 SOC_OCV 寄存器
    Dim I As Integer
    Dim j   As Long
    For I = 0 To 21
      For j = 0 To 5
         SOC_OCVData(I, j) = ""
      Next j
    Next I
End Function
Public Function clear_Thecap() '  清除 SOC_OCV 寄存器
    Dim I As Integer
    Dim j   As Long
    For I = 0 To 26
      For j = 0 To 3
         capData(I * 4 + j) = 0
      Next j
    Next I
End Function

Private Sub CommandopenOCV_Click()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s, kkk    As Integer
    Dim slen As Integer
    Dim bith, bitl  As Integer
    Dim rst(8) As Byte
    ' show open box
      CommonDialog_SOCOCV.ShowOpen
      filelocation = CommonDialog_SOCOCV.FileName
    ' input files into text1.text
    If CommonDialog_SOCOCV.FileName = "" Then
    Else
        Open filelocation For Input As #1
          I = 0
          Do Until EOF(1)
            Line Input #1, strsss
            For j = 0 To 5
              slen = Len(strsss)
              kkk = InStr(strsss, "$")                   ' 找到第一个空格
              SOC_OCVData(I, j) = Mid(strsss, 1, kkk - 1) ' 读出空格前数据
              strsss = Mid(strsss, kkk + 1, slen)            ' 清出这个读出数据 继续
            Next j
            If I < 49 Then
                I = I + 1
            End If
           Loop
        Call PrintfTheSOCOCV
        Close #1
    End If
End Sub

Private Sub CommandreadOCV_Click()
    Call clear_TheSOCOCV
    Delay_dis_ReadSOC_OCV = 4  ' 延时处理 回复数据
    NextSentCmd = CMD_ReadSOC_OCV
    CMD_cmd_No = 0
    Record_Num = 0
    manual_time = 5 ' 500ms 发送间隔
End Sub

Private Sub CommandSaveOCV_Click()

    Dim I, j As Integer
    Dim filelocation As String
    ' loads save as box
      CommonDialog_SOCOCV.ShowSave
      filelocation = CommonDialog_SOCOCV.FileName
       If CommonDialog_SOCOCV.FileName = "" Then
       Else
       
    ' append saves over file if it assists
    '  Open filelocation For Append As #1
      Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的
      For I = 0 To 21
        For j = 0 To 5
            Print #1, SOC_OCVData(I, j) & "$";  '& vbCrLf 有分号 不分行
        Next j
        Print #1, ""   ' 下一行
      Next I
      Close #1
      End If
      
End Sub

Public Sub CommandwriteOCV_Click()
    If (BMS_admin_mode = 4) Or (BMS_admin_mode = 7) Then
        Delay_dis_WriteSOC_OCV = 4  ' 延时处理 回复数据
        NextSentCmd = CMD_WriteSOC_OCV
        CMD_cmd_No = 1
        Record_Num = 0
        manual_time = 5 ' 500ms 发送间隔
    End If
End Sub
Public Function form_allresize(X As Long, Y As Long)
  Dim Top, Left, Width, Height As Long

'frmMain.Top = 0
'frmMain.Left = 0
'frmMain.Width = 22800
'frmMain.Height = 12800


Top = 0
Left = 100
Width = X * 3 / 4
Height = Y * 778 / 1080 ';781 ' y * 50 * 1024 / 2 ' 22500  781
Frame_SYSCONFIG.Width = Width
Frame_SYSCONFIG.Height = Height
Frame_SYSCONFIG.Top = Top
Frame_SYSCONFIG.Left = Left

FrameREG.Width = Width
FrameREG.Height = Height
FrameREG.Top = Top
FrameREG.Left = Left

Frame2.Width = Width
Frame2.Height = Height
Frame2.Top = Top
Frame2.Left = Left

Frame_Calib.Width = Width
Frame_Calib.Height = Height
Frame_Calib.Top = Top
Frame_Calib.Left = Left

Frame_Record.Width = Width
Frame_Record.Height = Height
Frame_Record.Top = Top
Frame_Record.Left = Left

FrameSOC_OCV.Width = Width
FrameSOC_OCV.Height = Height
FrameSOC_OCV.Top = Top
FrameSOC_OCV.Left = Left

Framecap.Width = Width
Framecap.Height = Height
Framecap.Top = Top
Framecap.Left = Left

Frame_SYSMUC2.Width = Width
Frame_SYSMUC2.Height = Height
Frame_SYSMUC2.Top = Top
Frame_SYSMUC2.Left = Left
End Function
Private Sub Form_Load()
Dim I As Integer
Dim xx As Long
Dim yy As Long

BMS_active_mode = 7
xx = Screen.Width / Screen.TwipsPerPixelX '获取屏幕bai宽度
yy = Screen.Height / Screen.TwipsPerPixelY '获取屏幕高度
 
Call form_allresize(xx, yy)
BMS_admin_mode = 7

'Call regist_work
Frame_SYSCONFIG.Visible = False
 For I = 1 To 100
   Load LabelSYS(I)
   Load Textsys(I)
 Next I

Call FrameSYSconfig_Load
CmdSYSjiazai.Top = 11000
CmdSYSRead.Top = 11000
CmdSYSWrite.Top = 11000
CmdSysSaveclink.Top = 11000

Frame2.Visible = False
Frame_SYSMUC2.Visible = False
FrameSOC_OCV.Visible = False
Framecap.Visible = False
Frame_Record.Visible = False
Frame_Calib.Visible = False

 For I = 1 To 100
   Load LabeSYS2(I)
   Load TexSys2(I)
 Next I
 Call FrameSYS2config_Load
Command_sys2load.Top = 11800 - 800
Command_Sys2Save.Top = 11800 - 800
Command_Sys2Read.Top = 11800 - 800



'jiema_button.Top = 11800 - 800
'jiema_button.BackColor = &H8000000D
Command_Sys2Write.Top = 11800 - 800

CmdSYSjiazai.Top = 11800 - 800
CmdSysSaveclink.Top = 11800 - 800
CmdSYSRead.Top = 11800 - 800
CmdSYSWrite.Top = 11800 - 800

Command_sys2load.Left = 4000 * 0
Command_Sys2Save.Left = 4000 * 1
Command_Sys2Read.Left = 4000 * 2
Command_Sys2Write.Left = 4000 * 3


'jiema_button.Left = 2000 * 1

CmdSYSjiazai.Left = 4000 * 0
CmdSysSaveclink.Left = 4000 * 1
CmdSYSRead.Left = 4000 * 2
CmdSYSWrite.Left = 4000 * 3

Command1.Top = 11800 - 800
Command2.Top = 11800 - 800
Command_EraseBalckUp.Top = 11800 - 800
Command_clearBlackup.Top = 11800 - 800
Command_SaveBlackupData.Top = 11800 - 800

CmdRegjiazai.Top = 11800 - 800
CmdRegSave.Top = 11800 - 800
CmdRegReadout.Top = 11800 - 800
CmdRegSent.Top = 11800 - 800


For I = 0 To 5
Frame3(I).Width = 2700
Frame3(I).Height = 6000
Frame3(I).Left = I * 2700 + I * 100 + 50
Frame3(I).Top = 300
Next I
 myRealV82Info.RealTempNum = 31
 myRealV82Info.Vcell_num = 10
 For I = 1 To 100
   Load LabelV82(I)
   LabelV82(I).Visible = False
 Next I
  For I = 1 To 50
   Load Labelpinhenguse(I)
   Labelpinhenguse(I).Visible = False
 Next I
myRealV82Info.Vcell_num = 16

RecordTime = 5
cyInfoTime = 5
manual_time = 5 ' 500ms 发送间隔
Call chuanshu.myInitForm_Load
Call Form1_can.myInitForm_Load
Call Frame2_Load
Call label_load '状态 窗口 显示初始化
Call load_backup
Call load_SOC_COVGrid
Call load_capGrid
Call AFEreg_temp_text_nodo
Call Frame_Visible_off
Frame2.Visible = True
manual_time = cyInfoTime
End Sub



Private Sub Frame_Callib_DragDrop(Source As Control, X As Single, Y As Single)


End Sub
' 单元格被点击时的事件
Private Sub msflexgrid1_entercll()

  MSFlexGrid1.CellBackColor = vbRed '点中单元格变蓝色

End Sub

Private Sub jilux1_Change(Index As Integer)
If LabelwenduCalib(Index).Caption = "" Then
    LabelwenduCalib(Index).Caption = "√"
Else
    LabelwenduCalib(Index).Caption = ""
End If
End Sub

Private Sub FrameREG_Click()
    Call frmMain.ReadTheRegchang
End Sub

Private Sub FrameS1_Click()
    Call frmMain.ReadTheRegchang
End Sub

Private Sub FrameS2_Click()
    Call frmMain.ReadTheRegchang
End Sub

Private Sub FrameS4_Click()
    Call frmMain.ReadTheRegchang
End Sub

Private Sub get_real_cap_Click()

  get_real_cap.Caption = "InMAXV" & "=" & "*" & "//" & "InMAXA" & "=" & "*" & "//" & "OutMAXA" & "=" & "*" & "//" & "Csg" & "=" & "*" & "//""Dsg" & "=" & "*" & "//"
 
  
End Sub

Private Sub JIHOU_BMS_Click()
    Dim aa As String
    If BMS_active_mode = 22 Then
        BMS_active_mode = 0
        My_msgbox "退出BMS激活模式！"
    Else
        aa = InputBox("输入打开BMS激活模式密码")
        If aa = "shen888" Then
            BMS_active_mode = 22
            生产解码.Visible = True
        Else
            My_msgbox "密码错误，请重新输入密码！"
        End If
    End If

   
 
End Sub

Private Sub Label9_Click()

End Sub



Private Sub Labeljilu_Click()
Dim I As Integer
Dim strrr As String
If Labeljilu.Caption = "" Then
   Labeljilu.Caption = "√"
   jilu_path = "\" & "记录" & "\" & Format(Now, "yyyymmddhhmmss")   ' 统一 时间格式
    If Dir(App.Path & "\" & "记录", vbDirectory) = "" Then   '判断bai文件du夹zhidao是否存zhuan在
        MkDir (App.Path & "\" & "记录") '创建shu文件夹 My_msgbox ("创建完毕")
    End If
    Open App.Path & jilu_path & ".txt" For Output As #1
 
 
strrr = "时间"
strrr = strrr & "    " & "总压"
strrr = strrr & "    " & "电流"
strrr = strrr & "    " & "SOC%"
For I = 0 To 15
strrr = strrr & "    " & "CELL" & I + 1
Next I

    
For I = 0 To 4
   If myRealV82Info.RealTempNum And (2 ^ I) Then
    strrr = strrr & "    " & "Temp" & I + 1
   End If
Next I
strrr = strrr & "    " & "电压状态"
strrr = strrr & "    " & "电流状态"
strrr = strrr & "    " & "温度状态"
strrr = strrr & "    " & "报警状态"
strrr = strrr & "    " & "FET状态"
strrr = strrr & "    " & "最高电芯"
strrr = strrr & "    " & "最低电芯"
strrr = strrr & "    " & "最高电芯电压"
strrr = strrr & "    " & "最低电芯电压"
strrr = strrr & "    " & "均衡值"
strrr = strrr & "    " & "累计放电次数"
strrr = strrr & "    " & "累计充电次数"

strrr = strrr & "    " & "当前剩余容量"
strrr = strrr & "    " & "满充容量"
Print #1, strrr
Close #1
   

Else
    Labeljilu.Caption = ""
End If

End Sub
Private Sub regist_work()
Dim I, j As Integer
Dim strssscal As String
Dim strsssread As String
Dim gettxt, readmadaddress  As String

Dim dddl As Long
Dim start_data1 As Date, end_data2 As Date, lasttime, nsssow As Date
Dim leijiadate, crc_data_read, crc_data_write As Long
Dim filelocation As String
    Shell "cmd.exe /c ipconfig /all >  D:\Program Files (x86)\qfsystem32\DTSNeoPCDLL64.dll"
    gettxt = ""
    
    Open "D:\Program Files (x86)\qfsystem32\DTSNeoPCDLL64.dll" For Input As #1
        For I = 1 To 20
            Line Input #1, gettxt
               j = InStr(gettxt, "物理地址")
               If j >= 1 Then
                gettxt = Replace(gettxt, "物理地址", "")
                gettxt = Replace(gettxt, " ", "")
                gettxt = Replace(gettxt, ".", "")
                gettxt = Replace(gettxt, "-", "")
                gettxt = Replace(gettxt, ":", "")
                readmadaddress = gettxt
                I = 40
               End If
        Next I
    Close #1
    
    filelocation = "D:\Program Files (x86)\qfsystem32\DTSNeoPCDLL32.dll"
        
        Open filelocation For Input As #1  ' 这里清除了 然后 保存自己的
        If EOF(1) Then
        
        Else
         Line Input #1, strsssread
        End If
        Close #1
            For j = 0 To 7
             strssscal = CRC_keycode(j & readmadaddress)     ' 加密时 加上 用户等级
             If strsssread = strssscal Then
                BMS_admin_mode = j
                GoTo ok_too
             End If
            Next j

            Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的7EF55CC5
            strsssread = InputBox("联系19921057745发下框显示码，输入到获取授权码", "软件授权", readmadaddress)
            Print #1, strsssread
            Print #1, strsssread
            Close #1
            For j = 0 To 7
                strssscal = CRC_keycode(j & readmadaddress)   ' 加密时 加上 用户等级
                If strsssread = strssscal Then
                    My_msgbox "授权码成功！"
                End If
            Next j
            
            End         '程序结束

 
ok_too:
   
End Sub
Private Sub LabelwenduCalib_Click(Index As Integer)

If LabelwenduCalib(Index).Caption = "" Then
     
LabelwenduCalib(Index).Caption = "√"

Else
       LabelwenduCalib(Index).Caption = ""
End If
End Sub

Private Sub LabelwenduCaliball_Click()
Dim I  As Integer
If LabelwenduCaliball.Caption = "" Then
    For I = 0 To 8
       LabelwenduCalib(I).Caption = "√"
    Next I
    LabelwenduCaliball.Caption = "√"
Else
    For I = 0 To 8
       LabelwenduCalib(I).Caption = ""
    Next I
    LabelwenduCaliball.Caption = ""
End If
End Sub
Private Sub MNU_Edit_cls_CDMOSFET_Click()
    NextSentCmd = CMD_SetFET
    manual_time = 5 ' 500ms 发送间隔
    sent_result = &H80
    CMD_cmd_No = 0
End Sub

Private Sub MNU_Edit_op_CDMOSFET_Click()
    NextSentCmd = CMD_SetFET
    manual_time = 5 ' 500ms 发送间隔
    sent_result = &H83
    CMD_cmd_No = 0
End Sub

Private Sub MNU_Edit_op_close_manual_Click()
    NextSentCmd = CMD_SetFET
    manual_time = 5 ' 500ms 发送间隔
    sent_result = &H0
    CMD_cmd_No = 0
End Sub

Private Sub MNU_Edit_op_CMOSFET_Click()
    NextSentCmd = CMD_SetFET
    manual_time = 5 ' 500ms 发送间隔
    sent_result = &H82
    CMD_cmd_No = 0
End Sub
Private Sub MNU_Edit_op_DMOSFET_Click()
    NextSentCmd = CMD_SetFET
    manual_time = 5 ' 500ms 发送间隔
    sent_result = &H81
    CMD_cmd_No = 0
End Sub
Private Sub MNU_Edit_op_YMOSFET_Click()
    NextSentCmd = CMD_SetFET
    manual_time = 5 ' 500ms 发送间隔
    sent_result = &H84
    CMD_cmd_No = 0
End Sub
Private Sub MNU_RecordNow_Click()
    If MNU_RecordNow.Checked Then
        MNU_RecordNow.Checked = False
    Else
        MNU_RecordNow.Checked = True
    End If
End Sub

Private Sub moni_weixin_Read_soc_Click()
    Delay_dis_CALIB_CURRENT = 4
    NextSentCmd = CMD_ReadSOCSOP
    manual_time = 5 ' 500ms 发送间隔
  '  sent_result = &H81
   ' CMD_cmd_No = 0
End Sub

Private Sub moni_weixin_Read_volt_curr_Click()
    Delay_dis_CALIB_CURRENT = 4
    NextSentCmd = CMD_ReadVOLTAGE_CURREN
    manual_time = 5 ' 500ms 发送间隔
  '  sent_result = &H81
   ' CMD_cmd_No = 0
End Sub
' 单元格失去光标后事件
Private Sub msflexgrid1_LeaveCell()
  MSFlexGrid1.CellBackColor = vbWhite '变白色
End Sub
Private Sub HScrollMain1_Scroll()
 Call Frame2_Load
End Sub

Private Sub MNU_About_Click()
  frmAbout.Show 1, frmMain
End Sub
Private Sub MNU_Config_Code_Click()
    Dim aa As String
    aa = InputBox("输入程序升级密码")
    If aa = "666666" Then
        IAP.Show 1, frmMain
    Else
        My_msgbox "密码错误，再来一次！"
    End If
End Sub
Private Sub MNU_Config_Port_Click()
  chuanshu.Show 1, frmMain
End Sub
Private Sub CmdSysSaveclink_Click()
   Dim I As Integer
    Dim filelocation As String
    ' loads save as box
      CommonDialog1.ShowSave
      filelocation = CommonDialog1.FileName
    If CommonDialog1.FileName = "" Then
    Else
    ' append saves over file if it assists
    '  Open filelocation For Append As #1
      Open filelocation For Output As #1  ' 这里清除了 然后 保存自己的
      For I = 0 To 70
        Print #1, sysCaption(I) & "           =            " & Textsys(I) '& vbCrLf
      Next I
      Close #1
    End If
End Sub

Private Sub OCV下发_Click()
If mode_bit6 = 22 Then
  mode_bit6 = 0
   OCV下发.Caption = "√OCV下发"
    OCV下发.ForeColor = &HC0&
Else
  mode_bit6 = 22
   OCV下发.Caption = "×OCV下发"
   OCV下发.ForeColor = &HE0E0E0
End If

End Sub

Private Sub ReSet_MCU_Click()
    NextSentCmd = CMD_ReSet_MCU
    manual_time = 5 ' 500ms 发送间隔
End Sub
Private Sub setCyTime_Click()
  setCytimes.cyLabel1.Visible = True
  setCytimes.cyText.Visible = True
  setCytimes.cyText.Top = 1800
  setCytimes.cyLabel1.Top = 1800
  setCytimes.recordLabel1.Visible = False
  setCytimes.RecordText.Visible = False
  setCytimes.Show 1, frmMain
End Sub
Private Sub setRecordtime_Click()
  setCytimes.cyLabel1.Visible = False
  setCytimes.cyText.Visible = False
  setCytimes.recordLabel1.Visible = True
  setCytimes.RecordText.Visible = True
  setCytimes.recordLabel1.Top = 1800
  setCytimes.RecordText.Top = 1800
       setCytimes.RecordText = RecordTime * 100
 setCytimes.Show 1, frmMain
End Sub
Private Sub templab_Click(Index As Integer)
If templab(Index).Caption = "" Then
   templab(Index).Caption = "●"
McuV82SysConfig.TempsetNum = McuV82SysConfig.TempsetNum Xor (2 ^ Index)
Else
templab(Index).Caption = ""
    McuV82SysConfig.TempsetNum = McuV82SysConfig.TempsetNum And (Not (2 ^ Index))
End If
Call dis_McuV82SysConfig
End Sub
'30 A40 A50 A60 A70 A500 A  根据采样电阻 变动
Private Sub Text1_Change(Index As Integer)
Dim temp, k As Double
Dim I As Integer
Dim sstr As String
Dim floatt As Double
temp = Val(Text1(16).Text)

If temp Then
  For I = 0 To ComboRtemp(0).ListCount - 1
    k = Val(Mid(ComboRtemp(0).List(I), 1, InStr(ComboRtemp(0).List(I), "A") - 1))
    k = k / temp
  ComboR1(16).List(I) = CStr(k) & "A"
 Next I
 
  For I = 0 To ComboRtemp(1).ListCount - 1
    k = Val(Mid(ComboRtemp(1).List(I), 1, InStr(ComboRtemp(1).List(I), "A") - 1))
    k = k / temp
  ComboR1(18).List(I) = CStr(k) & "A"
 Next I
  For I = 0 To ComboRtemp(2).ListCount - 1
    k = Val(Mid(ComboRtemp(2).List(I), 1, InStr(ComboRtemp(2).List(I), "A") - 1))
    k = k / temp
  ComboR1(20).List(I) = CStr(k) & "A"
 Next I
  For I = 0 To ComboRtemp(3).ListCount - 1
    k = Val(Mid(ComboRtemp(3).List(I), 1, InStr(ComboRtemp(3).List(I), "A") - 1))
    k = k / temp
  ComboR1(22).List(I) = CStr(k) & "A"
 Next I
 
   For I = 0 To ComboRtemp(4).ListCount - 1
    k = Val(Mid(ComboRtemp(4).List(I), 1, InStr(ComboRtemp(4).List(I), "A") - 1))
    k = k / temp
  ComboR1(26).List(I) = CStr(k) & "A"
 Next I
 End If
End Sub

'  读取 界面 txt值 到 SYS2CONFIG变量中
Public Sub Printf_McuSys2Config()
Dim temp As Long
Dim I, j As Integer
Dim strrrr  As String
I = 0
I = I + 0:    TexSys2(I) = Format(McuSys2Config.DesignVol / 10, "0.0")
I = I + 1:    TexSys2(I) = Int_to_hex(McuSys2Config.PackConfigMap)
I = I + 1:    TexSys2(I) = Format(McuSys2Config.FCC / 10, "0.0")
I = I + 1:    TexSys2(I) = Format(McuSys2Config.CycleThreshold / 10, "0.0")
I = I + 1:    TexSys2(I) = McuSys2Config.CycleCount
I = I + 1:    TexSys2(I) = McuSys2Config.NearFCC
I = I + 1:    TexSys2(I) = McuSys2Config.DfilterCur
I = I + 1:    TexSys2(I) = McuSys2Config.LearnLowTemp
I = I + 1:    TexSys2(I) = Format(McuSys2Config.SWVersion / 100, "0.00")
I = I + 1:    TexSys2(I) = Format(McuSys2Config.HWVersion / 100, "0.00")
I = I + 1:    TexSys2(I) = McuSys2Config.ShutDownDelay
I = I + 1:    TexSys2(I) = McuSys2Config.SelfDsgRate
I = I + 1:    TexSys2(I) = McuSys2Config.CommOffDelay
I = I + 1
    TexSys2(I) = McuSys2Config.MNFDate
I = I + 1
    TexSys2(I) = McuSys2Config.MNFName
I = I + 1
    TexSys2(I) = McuSys2Config.DeviceName
 I = I + 1
    TexSys2(I) = McuSys2Config.SN
I = I + 1:    TexSys2(I) = McuSys2Config.SOH
I = I + 1:    TexSys2(I) = McuSys2Config.MCU_ID
I = I + 1:    TexSys2(I) = McuSys2Config.KEY_CODE

        If McuSys2Config.KEY_CODE = CRC_keycode(McuSys2Config.MCU_ID) Then
            TexSys2(I).BackColor = &HC000&
            BMS_active_mode = 0 '解码完成 退出解码模式
            'jiema_button.BackColor = &HFF00&
            'jiema_button.Caption = "解码完成"
        Else
            TexSys2(I).BackColor = &HFF&
        End If
End Sub
' 执行自动配置 主要功能 按顺序 模拟人 下发数据 1 读MCU  判断 是否解码
'每 100ms 执行一次
Public Sub readtxt_sys2to_printf()

     Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s As Integer
    Dim kkk     As Integer
    Dim slen As Integer

        If 系统参数2下发.Caption = "√系统参数2下发" Then
          '每次开始 读下文件 到内存中，暂停
            I = 0
            j = 0
            s = 0
            strsss = ""
            filelocation = App.Path + "\自动配置用参数\系统参数2"
            If filelocation = "" Then
            Else
            ' input files into text1.text
             Open filelocation For Input As #1
               Do Until EOF(1)
                   Line Input #1, strsss
                   strsss = Replace(strsss, " ", "")
                   s = Len(strsss)
                   j = InStr(strsss, "=")
                   sys2Caption(I) = Mid(strsss, 1, j - 1)
                    LabeSYS2(I).Caption = sys2Caption(I)
                   TexSys2(I) = Mid(strsss, j + 1, s)
                   I = I + 1
               Loop
             Close #1
            End If
     End If
End Sub
     
Public Sub deal_auto()
Dim strsss As String

auto_500ms = auto_500ms + 1

If auto_500ms < 10 Then

    Exit Sub
End If
auto_500ms = 0
If 配置下发开始按钮.Caption = "暂停" Then
 自动配置结果.SelStart = Len(自动配置结果)
 ' 需要 每 500ms 读一次 直到读到 系统参数2 ，如果解码直接全部完成，如果未下一步
' jingdu1 =1 下发读取MCU ID ，=2 发送给C# =3 从C# 读出 =4 下发解码值 5 读取解码完成否 10 下发 硬件配置
    Select Case jingdu1
    Case 0  '开始第
    
         AUTO_SNUM = AUTO_SNUM + 1
         AUTO_NUM.Caption = AUTO_SNUM
         Call clean_disbox '清除弹出窗
         
         Flag_sys2ok = False
        If 解码.Caption = "√解码" Then
            解码.ForeColor = &H40C0&
        End If
        If 硬件配置下发.Caption = "√硬件配置下发" Then
            硬件配置下发.ForeColor = &H40C0&
        End If
        If 系统参数1下发.Caption = "√系统参数1下发" Then
            系统参数1下发.ForeColor = &H40C0&
        End If
        If 系统参数2下发.Caption = "√系统参数2下发" Then
            系统参数2下发.ForeColor = &H40C0&
        End If
        If CAP下发.Caption = "√CAP下发" Then
            CAP下发.ForeColor = &H40C0&
        End If
        If OCV下发.Caption = "√OCV下发" Then
            OCV下发.ForeColor = &H40C0&
        End If
        If 时间校正.Caption = "√时间校正" Then
            时间校正.ForeColor = &H40C0&
        End If
        If 电流校零.Caption = "√电流校零" Then
            电流校零.ForeColor = &H40C0&
        End If
        If 记录擦除.Caption = "√记录擦除" Then
            记录擦除.ForeColor = &H40C0&
        End If
 
         
         Flag_sys2ok = False
         自动配置结果.Text = 自动配置结果.Text + vbCrLf + "开始第" + str(AUTO_SNUM) + " 次"
         jingdu1 = 1
    Case 1  '下发读取MCU ID
          If 解码.Caption = "√解码" Then
                     ' 有回复好 判断
            If Flag_readmcusys2ok = True Then
                  ' 读出MCU ID 后，发给C# 远程
                  manual_time = 0 '
                   If McuSys2Config.KEY_CODE = CRC_keycode(McuSys2Config.MCU_ID) Then
                    jingdu1 = 10 ' 进入一下步
                    Flag_readmcusys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "1已经解码过" ' 记录本次 解码BMS ，下次BMS 将不再解码
                   Else
                    jingdu1 = 2 ' 进入一下步
                    Flag_readmcusys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "2收到MCUID:" + McuSys2Config.MCU_ID ' 记录本次 解码BMS ，下次BMS 将不再解码
                   End If
            Else
                 If Delay_dis_Readsys2config = 0 Then
                     Flag_readmcusys2ok = False
                     自动配置结果.Text = 自动配置结果.Text + vbCrLf + "1发送读取MCU_ID中"
                     Delay_dis_Readsys2config = 10  ' =0 不等待显示
                     NextSentCmd = CMD_ReadSys2Config
                     manual_time = 0 ' 500ms 发送间隔
                End If
            End If
        Else
            jingdu1 = 10
        End If
    Case 2 '保存文件给C#
        Open "C:\tocsharp_temp2022" For Output As #1  ' 这里清除了 然后 保存自己的
        Print #1, McuSys2Config.MCU_ID & "$$$"   '& vbCrLf 有分号 不分行
        Close #1
        jingdu1 = 3
        自动配置结果.Text = 自动配置结果.Text + vbCrLf + "3发给解码程序处理"
   Case 3 '读取
         
           If Dir("C:\tovb_temp2022") = "" Then
            strsss = ""
           Else
           
            Open "C:\tovb_temp2022" For Input As #1
             If EOF(1) = False Then
              Line Input #1, strsss
             End If
                
            Close #1
           End If

          
          If InStr(strsss, "=") Then
            自动配置结果.Text = 自动配置结果.Text + vbCrLf + "4取到解码值:" + strsss
            strsss = Mid(strsss, 21, 8)
            McuSys2Config.KEY_CODE = strsss
            LAST_MCU_ID = McuSys2Config.MCU_ID ' 记录本次 解码BMS ，下次BMS 将不再解码
            jingdu1 = 4
            Flag_onlysys2ok = False
                Open "C:\tocsharp_temp2022" For Output As #1
                Print #1, "..." ' 这里清除了
                Close #1
          Else
            自动配置结果.Text = 自动配置结果.Text + "."
          End If
 
    
     Case 4   '
                   ' 有回复好 判断
            If Flag_onlysys2ok = True Then
                  jingdu1 = 5 ' 进入一下步
                  Flag_onlysys2ok = False
                  Flag_readckeckjiemasys2ok = False
                  
            Else
                 If Delay_dis_Writesys2config = 0 Then
                     Flag_onlysys2ok = False
                     Delay_dis_Writesys2config = 8  '等待
                     NextSentCmd = CMD_WriteSys2Config
                     manual_time = 5 ' 500ms 发送间隔
                      
                     自动配置结果.Text = 自动配置结果.Text + vbCrLf + "4下发解码给BMS"
                End If
            
            End If
        
    Case 5  '
                   ' 有回复好 判断
            If Flag_readckeckjiemasys2ok = True Then
                If McuSys2Config.KEY_CODE = CRC_keycode(McuSys2Config.MCU_ID) Then
                  '解码对 下一步
                  manual_time = 0 '
                  LAST_MCU_ID = McuSys2Config.MCU_ID ' 记录本次 解码BMS ，下次BMS 将不再解码
                  jingdu1 = 10 ' 进入一下步
                  自动配置结果.Text = 自动配置结果.Text + vbCrLf + "5解码完成"
                  解码.ForeColor = &HC000&
                  Flag_readckeckjiemasys2ok = False
                Else
                    Flag_readckeckjiemasys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "5解码失败,重试中"
                    jingdu1 = 1
                End If
            End If
            
           If Delay_dis_Readsys2config = 0 Then
                Flag_readckeckjiemasys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "5确认解码成功否？"
                Delay_dis_Readsys2config = 20  ' =0 不等待显示
                NextSentCmd = CMD_ReadSys2Config
                manual_time = 0 ' 500ms 发送间隔
           End If

 
    Case 10 ' 需要 每 500ms 下发一次   硬件配置
        If 硬件配置下发.Caption = "√硬件配置下发" Then
                   ' 有回复好 判断
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "10硬件配置OK"
                    硬件配置下发.ForeColor = &HC000&
                    jingdu1 = 20
            End If
            
           If Delay_dis_WriteRegAfe = 0 Then
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "10下发硬件配置"
                Delay_dis_WriteRegAfe = 20  ' =0 不等待显示
                NextSentCmd = CMD_WriteAFEseg
                manual_time = 0 ' 500ms 发送间隔
           End If
        Else
            jingdu1 = 20
        End If
    Case 20
        If 系统参数1下发.Caption = "√系统参数1下发" Then
                   ' 有回复好 判断
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "20系统参数1 OK"
                    jingdu1 = 30
                    系统参数1下发.ForeColor = &HC000&
            End If
            
           If Delay_dis_Writesysconfig = 0 Then
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "20下发系统参数1"
                Delay_dis_Writesysconfig = 40  ' =0 不等待显示
                NextSentCmd = CMD_WriteSysConfig
                manual_time = 5 ' 500ms 发送间隔
           End If
        Else
            jingdu1 = 30
        End If
    Case 30
        If 系统参数2下发.Caption = "√系统参数2下发" Then
    
                   ' 有回复好 判断
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "30统参数2 OK"
                    jingdu1 = 40
                    系统参数2下发.ForeColor = &HC000&
            End If
            

     
         If Delay_dis_Writesys2config = 0 Then
          Call readtxt_sys2to_printf
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "30下发统参数2"
                Delay_dis_Writesys2config = 10  ' =0 不等待显示
                NextSentCmd = CMD_WriteSys2Config
                manual_time = 5 ' 500ms 发送间隔
           End If
        Else
            jingdu1 = 40
        End If
    Case 40
        If CAP下发.Caption = "√CAP下发" Then
                   ' 有回复好 判断
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "40 CAP OK"
                    jingdu1 = 50
                    CAP下发.ForeColor = &HC000&
            End If
            
           If Delay_dis_Writecap = 0 Then
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "40下发CAP"
                Delay_dis_Writecap = 60  ' =0 不等待显示
                NextSentCmd = CMD_Writecap
                manual_time = 0 ' 500ms 发送间隔
           End If
        Else
            jingdu1 = 50
        End If
    Case 50
        If OCV下发.Caption = "√OCV下发" Then
                   ' 有回复好 判断
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "50 OCV OK"
                    jingdu1 = 60
                    OCV下发.ForeColor = &HC000&
            End If
            
           If Delay_dis_WriteSOC_OCV = 0 Then
                Flag_sys2ok = False
                BMS_admin_mode = 7
                NextSentCmd = CMD_WriteSOC_OCV
                CMD_cmd_No = 1
                Record_Num = 0
                manual_time = 5 ' 500ms 发送间隔
                
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "50下发OCV"
                Delay_dis_WriteSOC_OCV = 60  ' =0 不等待显示
          
           End If
           
        Else
            jingdu1 = 60
        End If
    Case 60
        If 时间校正.Caption = "√时间校正" Then
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "60 时间校正 OK"
                    jingdu1 = 70
                    时间校正.ForeColor = &HC000&
            End If
            
            If Delay_dis_CALIB_RTC = 0 Then
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "60下发时间校正"
                Delay_dis_CALIB_RTC = 20  ' =0 不等待显示
                NextSentCmd = CMD_CALIB_RTC
                manual_time = 0 ' 500ms 发送间隔
            End If
        Else
            jingdu1 = 70
        End If
    Case 70
        If 电流校零.Caption = "√电流校零" Then
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "70 电流校零 OK"
                    jingdu1 = 80
                    电流校零.ForeColor = &HC000&
            End If
            
           If Delay_dis_CALIB_CURRENT = 0 Then
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "70下发电流校零"
                Delay_dis_CALIB_CURRENT = 20  ' =0 不等待显示
                NextSentCmd = CMD_CALIB_CURRENT
                CMD_cmd_No = 2
                manual_time = 0 ' 500ms 发送间隔
           End If
        Else
            jingdu1 = 80
        End If
    Case 80
        If 记录擦除.Caption = "√记录擦除" Then
            If Flag_sys2ok = True Then
                    Flag_sys2ok = False
                    自动配置结果.Text = 自动配置结果.Text + vbCrLf + "80 记录擦除 OK"
                    jingdu1 = 80
                    记录擦除.ForeColor = &HC000&
                    NextSentCmd = CMD_ReadInfo
                jingdu1 = 200

                    
            End If
            
           If Delay_dis_EraseBalckUp = 0 Then
                Flag_sys2ok = False
                自动配置结果.Text = 自动配置结果.Text + vbCrLf + "80下发记录擦除"
                Delay_dis_EraseBalckUp = 20  ' =0 不等待显示
                NextSentCmd = CMD_EraseBalckUp
                CMD_cmd_No = 2
                manual_time = 0 ' 500ms 发送间隔
           End If
        Else
            jingdu1 = 200
        End If
    Case 200
                jingdu1 = 0
                Call clean_disbox '清除弹出窗
                If 连续.Caption = "√单台配置" Then
                    配置下发开始按钮.Caption = "开始"
                Else
                    配置下发开始按钮.Caption = "暂停"
                      My_msgbox ("本台BMS已配置好， 请更换下一台")
                End If
                
            自动配置结果.Text = 自动配置结果.Text + vbCrLf + "This BMS OK  Next"
            
    End Select
End If

End Sub

' 定时 显示变量值 给各界面
Private Sub Timer_Timer()
Dim I, j, y1 As Integer
Dim zhens, xiaos As Long
Dim Mcutime As Date
Dim sstrg As String

 
    
   Text_TimeRTC.Text = myRealV82Info.Time_t
   DisVolteText(0).Text = Format(myRealV82Info.Vbat, "0.0")
   DisCurrentText(3).Text = Format(myRealV82Info.Curr, "0.000")
   disTemptext(1).Text = Format(myRealV82Info.temp(0), "0.0")
   disTemptext(2).Text = Format(myRealV82Info.temp(1), "0.0")
   disTemptext(3).Text = Format(myRealV82Info.temp(2), "0.0")
   disTemptext(4).Text = Format(myRealV82Info.temp(3), "0.0")
   disTemptext(5).Text = Format(myRealV82Info.temp(4), "0.0")
   disTemptext(6).Text = Format(myRealV82Info.temp(5), "0.0")
   disTemptext(7).Text = Format(myRealV82Info.temp(6), "0.0")
   disTemptext(8).Text = Format(myRealV82Info.temp(7), "0.0")
   disTemptext(9).Text = Format(myRealV82Info.temp(8), "0.0")
   
   
If BMS_active_mode = 22 Then
 For I = 0 To 18
    TexSys2(I).BackColor = &HC0FFC0
 Next I
Else
 For I = 0 To 18
    TexSys2(I).BackColor = &HFFFFFF
 Next I
End If



 If Delay_dis_Readsysconfig > 0 Then
    Delay_dis_Readsysconfig = Delay_dis_Readsysconfig - 1
    If Delay_dis_Readsysconfig > 20 Then
            Call dis_McuV82SysConfig
           My_msgbox ("读取配置1成功")
            Delay_dis_Readsysconfig = 0
    Else
        If Delay_dis_Readsysconfig = 0 Then
            ' 超时 &H00C0E0FF&
            My_msgbox ("读取配置1失败")
        End If
    End If
 End If
 
  If Delay_dis_Writesysconfig > 0 Then
    Delay_dis_Writesysconfig = Delay_dis_Writesysconfig - 1
    If Delay_dis_Writesysconfig > 20 Then
            My_msgbox ("下发配置1成功")
            Delay_dis_Writesysconfig = 0
    Else
        If Delay_dis_Writesysconfig = 0 Then
            ' 超时
            My_msgbox ("下发配置1失败")
        End If
    End If
 End If
 
 If Delay_dis_ReadRegAfe > 0 Then
    Delay_dis_ReadRegAfe = Delay_dis_ReadRegAfe - 1
    If Delay_dis_ReadRegAfe > 20 Then
           
            My_msgbox ("读取AFE配置0成功")
            Delay_dis_ReadRegAfe = 0
    Else
        If Delay_dis_ReadRegAfe = 0 Then
            ' 超时
            My_msgbox ("读取AFE配置0失败")
        End If
    End If
 End If
 
  If Delay_dis_WriteRegAfe > 0 Then
    Delay_dis_WriteRegAfe = Delay_dis_WriteRegAfe - 1
 
    If Delay_dis_WriteRegAfe > 80 Then
           
            My_msgbox ("下发AFE配0置成功")
            Delay_dis_WriteRegAfe = 0
    Else
        If Delay_dis_WriteRegAfe = 0 Then
            ' 超时
            My_msgbox ("下发AFE配置0失败")
        End If
    End If
 End If

  If LOAD_CELLmun = myRealV82Info.Vcell_num Then ' 如果 下面电芯节数变化，界面更新一次
  Else
     Call Frame2_Load
  End If
  If LOAD_Tempmun = myRealV82Info.RealTempNum Then ' 如果 下面电芯节数变化，界面更新一次
  Else
     Call Frame2_Load
  End If
  
  If havegetTRightData Then
    havegetTRightData = 0
    Call Printf_myRealV82Info ' 显示 myRealV82Info 内存值
  End If
  
   If Delay_dis_Readsys2config > 0 Then
    Delay_dis_Readsys2config = Delay_dis_Readsys2config - 1
    If Delay_dis_Readsys2config > 20 Then
            Call Printf_McuSys2Config
            My_msgbox ("读取配置2成功")
            Delay_dis_Readsys2config = 0
    Else
        If Delay_dis_Readsys2config = 0 Then
            ' 超时
           My_msgbox ("读取配置2失败")
        End If
    End If
  End If
   If Delay_dis_Writesys2config > 0 Then
        Delay_dis_Writesys2config = Delay_dis_Writesys2config - 1
        If Delay_dis_Writesys2config > 20 Then
               My_msgbox ("下发配置2成功")
                Delay_dis_Writesys2config = 0
        Else
            If Delay_dis_Writesys2config = 0 Then
                ' 超时
               My_msgbox ("下发配置2失败")
            End If
        End If
    End If
 
   If Delay_dis_EraseBalckUp > 0 Then
        Delay_dis_EraseBalckUp = Delay_dis_EraseBalckUp - 1
        If Delay_dis_EraseBalckUp > 20 Then
            If BMS_active_mode = 0 Then
                My_msgbox ("擦除成功")
            End If
                
                Delay_dis_EraseBalckUp = 0
        Else
            If Delay_dis_EraseBalckUp = 0 Then
                ' 超时
                My_msgbox ("擦除超时失败")
            End If
        End If
    End If
 
   If Delay_dis_Enter_Sleep_Mode > 0 Then
        Delay_dis_Enter_Sleep_Mode = Delay_dis_Enter_Sleep_Mode - 1
        If Delay_dis_Enter_Sleep_Mode > 20 Then
                My_msgbox ("BMS马上进入待机模式")
                Delay_dis_Enter_Sleep_Mode = 0
        Else
            If Delay_dis_Enter_Sleep_Mode = 0 Then
                ' 超时
                My_msgbox ("下发命令失败")
            End If
        End If
    End If
    If Delay_dis_Enter_WORK_Mode > 0 Then
        Delay_dis_Enter_WORK_Mode = Delay_dis_Enter_WORK_Mode - 1
        If Delay_dis_Enter_WORK_Mode > 20 Then
                My_msgbox ("BMS马上进入开机模式")
                Delay_dis_Enter_WORK_Mode = 0
        Else
            If Delay_dis_Enter_WORK_Mode = 0 Then
                ' 超时
                My_msgbox ("下发命令失败")
            End If
        End If
    End If
   If Delay_dis_SetFET > 0 Then
        Delay_dis_SetFET = Delay_dis_SetFET - 1
        If Delay_dis_SetFET > 20 Then
                My_msgbox ("下发MOSFET命令成功")
                Delay_dis_SetFET = 0
        Else
            If Delay_dis_SetFET = 0 Then
                ' 超时
                My_msgbox ("下发命令失败")
            End If
        End If
    End If
    
    If Delay_dis_Readcap > 0 Then
        Call PrintfThecap
        Delay_dis_Readcap = Delay_dis_Readcap - 1
        If Delay_dis_Readcap > 20 Then
             
                  My_msgbox ("读取cap配置完成")  ' 超时
                   manual_time = 0
                  NextSentCmd = CMD_ReadInfo
                   
                Delay_dis_Readcap = 0
        Else
            If Delay_dis_Readcap = 0 Then
             NextSentCmd = CMD_ReadInfo
              manual_time = 0
                My_msgbox ("读取cap失败")  ' 超时
            End If
        End If
    End If
   If Delay_dis_Writecap > 0 Then
        Delay_dis_Writecap = Delay_dis_Writecap - 1
        If Delay_dis_Writecap > 20 Then
                NextSentCmd = CMD_ReadInfo
                manual_time = 0
                My_msgbox ("下发配置成功")
                Delay_dis_Writecap = 0
        Else
            If Delay_dis_Writecap = 0 Then
                ' 超时
                             NextSentCmd = CMD_ReadInfo
              manual_time = 0
                My_msgbox ("下发配置失败")
            End If
        End If
    End If
  
   If Delay_dis_ReadSOC_OCV > 0 Then
        Call PrintfTheSOCOCV
        Delay_dis_ReadSOC_OCV = Delay_dis_ReadSOC_OCV - 1
        If Delay_dis_ReadSOC_OCV > 20 Then
             
                  My_msgbox ("读取SOC_OCV配置完成")  ' 超时
                   manual_time = 0
                  NextSentCmd = CMD_ReadInfo
                   
                Delay_dis_ReadSOC_OCV = 0
        Else
            If Delay_dis_ReadSOC_OCV = 0 Then
             NextSentCmd = CMD_ReadInfo
              manual_time = 0
                My_msgbox ("读取SOC_OCV失败")  ' 超时
            End If
        End If
    End If
   If Delay_dis_WriteSOC_OCV > 0 Then
        Delay_dis_WriteSOC_OCV = Delay_dis_WriteSOC_OCV - 1
        If Delay_dis_WriteSOC_OCV > 20 Then
                NextSentCmd = CMD_ReadInfo
                manual_time = 0
                My_msgbox ("下发配置成功")
                Delay_dis_WriteSOC_OCV = 0
        Else
            If Delay_dis_WriteSOC_OCV = 0 Then
                ' 超时
                             NextSentCmd = CMD_ReadInfo
              manual_time = 0
                My_msgbox ("下发配置失败")
            End If
        End If
    End If
    
   If Delay_dis_CALIB_RTC > 0 Then
        Delay_dis_CALIB_RTC = Delay_dis_CALIB_RTC - 1
    End If
    
   If Delay_dis_ReadBalckUp > 0 Then
        Delay_dis_ReadBalckUp = Delay_dis_ReadBalckUp - 1
        If Delay_dis_ReadBalckUp > 190 Then
                  My_msgbox ("读到0条记录,eeprom不正常")  ' 超时
                  NextSentCmd = CMD_ReadInfo
                 manual_time = 0
                Delay_dis_ReadBalckUp = 0
        Else
            If Delay_dis_ReadBalckUp > 20 Then
                      My_msgbox ("读取记录完成")  ' 超时
                      NextSentCmd = CMD_ReadInfo
                       manual_time = 0
                    Delay_dis_ReadBalckUp = 0
            Else
                If Delay_dis_ReadBalckUp = 0 Then
                 NextSentCmd = CMD_ReadInfo
                  manual_time = 0
                    My_msgbox ("读取记录失败")  ' 超时
                End If
            End If
        End If
    End If
    
   If Delay_dis_CALIB_VOLTAGE > 0 Then
        Delay_dis_CALIB_VOLTAGE = Delay_dis_CALIB_VOLTAGE - 1
        If Delay_dis_CALIB_VOLTAGE > 20 Then
              '  Call dis_Printf_CALIB_frame
                  My_msgbox ("校正总压成功")
                Delay_dis_CALIB_VOLTAGE = 0
        Else
            If Delay_dis_CALIB_VOLTAGE = 0 Then
                My_msgbox ("校正失败")  ' 超时
            End If
        End If
    End If
   If Delay_dis_CALIB_CURRENT > 0 Then
        Delay_dis_CALIB_CURRENT = Delay_dis_CALIB_CURRENT - 1
        If Delay_dis_CALIB_CURRENT > 20 Then
              '  Call dis_Printf_CALIB_frame
                  My_msgbox ("校正电流成功")
                Delay_dis_CALIB_CURRENT = 0
        Else
            If Delay_dis_CALIB_CURRENT = 0 Then
                My_msgbox ("校正失败")  ' 超时
            End If
        End If
    End If
    
   If Delay_dis_CALIB_Temp > 0 Then
        Delay_dis_CALIB_Temp = Delay_dis_CALIB_Temp - 1
        If Delay_dis_CALIB_Temp > 20 Then
             '   Call dis_Printf_CALIB_frame
                  My_msgbox ("下发校正温度完成")
                Delay_dis_CALIB_Temp = 0
        Else
            If Delay_dis_CALIB_Temp = 0 Then
                My_msgbox ("下发失败")  ' 超时
            End If
        End If
    End If
End Sub
Public Function Printf_myRealV82Info() ' 显示 INFO 内存值
Dim X, I, j, y1 As Integer
Dim zhens, xiaos As Long
Dim xiaoshu As Single
Dim Mcutime As Date
Dim strtemp As String
        For I = 0 To myRealV82Info.Vcell_num - 1
            If (I) < 9 Then
                LabelV82(I).Caption = "电芯0" & (I + 1) & "：" & myRealV82Info.Vcell(I) & "   V"
            Else
                LabelV82(I).Caption = "电芯" & (I + 1) & "：" & myRealV82Info.Vcell(I) & "   V"
            End If
            If (myRealV82Info.BlanceState And (2 ^ I)) Then
                Labelpinhenguse(I).Caption = "√"
            Else
                Labelpinhenguse(I).Caption = ""
            End If
        Next
        LabelV82(I).Caption = "电池总电压：" & myRealV82Info.Vbat & "   V"
        I = I + 1
      '  xiaoshu = myRealV82Info.Curr   'Curr[0]充电电流，Curr[1]放电电流'
      '  xiaoshu = xiaoshu / 1000
        LabelV82(I).Caption = "实时电流 ：" & myRealV82Info.Curr & "A"
'        i = i + 1
'        If (i Mod 10) > 7 Then  ' 让第三行 占首位 排列整齐
'        i = i + 10 - (i Mod 10)
'        End If
'        i = i - 1
        y1 = 0
        If myRealV82Info.RealTempNum And (2 ^ 0) Then
        I = I + 1
        y1 = y1 + 1
            LabelV82(I).Caption = "环境温度  " & "  ：" & myRealV82Info.temp(0) & "  °C"
        Else
        End If
        
        If myRealV82Info.RealTempNum And (2 ^ 1) Then
        I = I + 1
        y1 = y1 + 1
            LabelV82(I).Caption = "功率温度  " & "  ：" & myRealV82Info.temp(1) & "  °C"
        Else
        End If
 
        For X = 2 To 7
            If myRealV82Info.RealTempNum And (2 ^ X) Then
              '  templab(x).Caption = "●"
                I = I + 1
                LabelV82(I).Caption = "电芯温度" & (X - 1) & "  ：" & myRealV82Info.temp(X) & "  °C"
            Else
            End If
        Next X
   '     i = i + 1
    ' '   xiaos = myRealV82Info.Vcell_num  ' '': 电池串数，1-16'
     '   LabelV82(i).Caption = " 电池串数 ：    " & xiaos & "  串"
     '   i = i + 1
      '  xiaos = myRealV82Info.TempNum  ' '' '
     '   LabelV82(i).Caption = "温度采集个数 ：   " & xiaos & "  个"
        I = I + 1
        xiaos = myRealV82Info.NUM_VOV + 1 ' '单体高压对应的电池的序号，例如 5 表示第 5 节高压
        LabelV82(I).Caption = "最高压电芯序号：     " & xiaos
        I = I + 1
        xiaos = myRealV82Info.NUM_VUV + 1 ' '单体高压对应的电池的序号，例如 5 表示第 5 节高压
        LabelV82(I).Caption = "最低压电芯序号：     " & xiaos
        I = I + 1
        xiaos = myRealV82Info.NUM_WARN_VHIGH  ' '单体高压警告对应的电池的序号
        LabelV82(I).Caption = "最高电芯电压：   " & xiaos
        I = I + 1
        xiaos = myRealV82Info.NUM_WARN_VLOW  ' '单体低压警告对应的电池的序号
        LabelV82(I).Caption = "最低电芯电压：   " & xiaos
        I = I + 1
        xiaos = myRealV82Info.DchgNum  '：放电次数'
        LabelV82(I).Caption = "累计放电次数：   " & xiaos
        I = I + 1
        xiaos = myRealV82Info.BatStatus  '：放电次数'
        xiaos = myRealV82Info.mcu_powerStatu  '： MUC状态'
        strtemp = ""
        
         If xiaos >= 64 Then
             xiaos = xiaos - 64
            strtemp = "MCU状态：激活  "
         Else
            strtemp = "MCU状态：待激活"
         End If
         
         If xiaos >= 32 Then
             xiaos = xiaos - 32
            strtemp = strtemp & ""
         Else
            strtemp = strtemp & "未配置"
         End If
         
         If xiaos >= 16 Then
             xiaos = xiaos - 16
            strtemp = strtemp & ""
         Else
            strtemp = strtemp & "未开机"
         End If
         
        If xiaos = 0 Then
            LabelV82(I).Caption = strtemp & "上电未初始化"
        End If
        If xiaos = 1 Then
            LabelV82(I).Caption = strtemp & "工作中"
        End If
        If xiaos = 2 Then
            LabelV82(I).Caption = strtemp & "进入休眠"
        End If
        If xiaos = 3 Then
            LabelV82(I).Caption = strtemp & "休眠"
        End If
        If xiaos = 4 Then
            LabelV82(I).Caption = strtemp & "关机中"
        End If
            If xiaos = 5 Then
        LabelV82(I).Caption = strtemp & "关机"
        End If

        ' BlanceState As Long ' ： 均衡状态，表示那一节电压开启均衡
        For j = 0 To Von - 1
            If (myRealV82Info.vstate And (2 ^ j)) Then
                LabelBitV(j).ForeColor = &HFF
            Else
                LabelBitV(j).ForeColor = &H8000000C
            End If
        Next j
         For j = 0 To Con - 1
            If (myRealV82Info.Cstate And (2 ^ j)) Then
                LabelBitC(j).ForeColor = &HFF
            Else
                LabelBitC(j).ForeColor = &H8000000C
            End If
        Next j
         For j = 0 To Ton - 1
            If (myRealV82Info.Tstate And (2 ^ j)) Then
                LabelBitT(j).ForeColor = &HFF
            Else
                LabelBitT(j).ForeColor = &H8000000C
            End If
        Next j
         For j = 0 To Aon - 1
            If (myRealV82Info.Alarm And (2 ^ j)) Then
                LabelBitA(j).ForeColor = &HFF
            Else
                LabelBitA(j).ForeColor = &H8000000C
            End If
        Next j
        
            j = 0
            If (myRealV82Info.Fetstate And (2 ^ j)) Then
                LabelBitF(j).ForeColor = &HC000&
                LabelBitF(j).Caption = "放电MOSs:导通"
            Else
                LabelBitF(j).ForeColor = &HFF
                LabelBitF(j).Caption = "放电MOSs:不导通"
            End If
            j = 1
            If (myRealV82Info.Fetstate And (2 ^ j)) Then
                LabelBitF(j).ForeColor = &HC000&
                LabelBitF(j).Caption = "充电MOSs:导通"
            Else
                LabelBitF(j).ForeColor = &HFF
                LabelBitF(j).Caption = "充电MOSs:不导通"
            End If
            
            j = 2
            If (myRealV82Info.Fetstate And (2 ^ j)) Then
                LabelBitF(j).ForeColor = &HC000&
                LabelBitF(j).Caption = "放电控制：打开"
            Else
                LabelBitF(j).ForeColor = &HFF
                LabelBitF(j).Caption = "放电控制：关闭"
            End If
            j = 3
            If (myRealV82Info.Fetstate And (2 ^ j)) Then
                LabelBitF(j).ForeColor = &HC000&
                LabelBitF(j).Caption = "充电控制：打开"
            Else
                LabelBitF(j).ForeColor = &HFF
                LabelBitF(j).Caption = "充电控制：关闭"
            End If

         For j = 4 To Fon - 1
            If (myRealV82Info.Fetstate And (2 ^ j)) Then
                LabelBitF(j).ForeColor = &HFF
            Else
                LabelBitF(j).ForeColor = &H8000000C
            End If
        Next j
            j = 6
            If (myRealV82Info.Fetstate And (2 ^ j)) Then
                LabelBitF(j).ForeColor = &HFF
                LabelBitF(j).Caption = "预充MOS：打开"
            Else
                LabelBitF(j).ForeColor = &H8000000C
                LabelBitF(j).Caption = "预充MOS：关闭"
            End If
            j = 7
            If (myRealV82Info.Fetstate And (2 ^ 9)) Then
                    LabelBitF(j).ForeColor = &HFF
                    LabelBitF(j).Caption = "手动控制中"
            Else
                    LabelBitF(j).ForeColor = &H8000000C
                    LabelBitF(j).Caption = "手动控制关闭"
            End If

            

            j = 0
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "OV"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "OV"
            End If
            j = 1
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "UV"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "UV"
            End If
            j = 2
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "OCD1"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "OCD1"
            End If
            j = 3
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "OCD2"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "OCD2"
            End If
                j = 4
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "OCC"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "OCC"
            End If
            j = 5
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "SC"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "SC"
            End If
            j = 6
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "PF"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "PF"
            End If
            j = 7
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "oneCELLMIss"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "oneCELLMIss"
            End If
            j = 8
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "UTC"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "UTC"
            End If
            j = 9
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "UTD"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "UTD"
            End If
            j = 10
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "UTD"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "UTD"
            End If
            j = 11
            If (myRealV82Info.BatStatus And (2 ^ j)) Then
                LabelBitG(j).ForeColor = &HFF&
                LabelBitG(j).Caption = "OPT"
            Else
                LabelBitG(j).ForeColor = &H8000000C
                LabelBitG(j).Caption = "OPT"
            End If
  
        I = I + 1
          '：放电次数'
        LabelV82(I).Caption = "剩余电量百分比SOC：" & myRealV82Info.SOC & " %"  '电池 soc ，百分比 0-100
        I = I + 1
          '：放电次数'

        LabelV82(I).Caption = "当前剩余容量：   " & myRealV82Info.CapNow & "  AH"   ' 当前容量 (0.1AH)
        I = I + 1
          '：放电次数'
     
        LabelV82(I).Caption = "满充容量：   " & myRealV82Info.CapFull & "  AH"     '满充容量(0.1AH)
        LabelV82(I).Visible = True
        I = I + 1
        
        LabelV82(I).Caption = "run_code：   " & Int_to_hex(myRealV82Info.FET_code)    '满充容量(0.1AH)
        LabelV82(I).Visible = True
        I = I + 1
        LabelV82(I).Caption = "AFE_TEMP1:  " & myRealV82Info.afe_Temp(1) & "   °C"
        LabelV82(I).Visible = True
        I = I + 1
        LabelV82(I).Caption = "AFE_TEMP2:  " & myRealV82Info.afe_Temp(2) & "   °C"
        LabelV82(I).Visible = True
        I = I + 1
        LabelV82(I).Caption = "AFE_TEMP3:  " & myRealV82Info.afe_Temp(3) & "   °C"
        LabelV82(I).Visible = True
        
      Text2(0).Text = "RTC:" & myRealV82Info.Time_t
End Function
Public Function PrintfTheReg()        ' 显示  RegEERPOM寄存器内空
    Dim I As Integer
    Dim j, s, kkk As Long
    Dim bith, bitl  As Long
    Dim rst(8) As Byte
               rst(0) = Val("&H" & Mid(RegEERPOM(0), 2, 1))
               rst(1) = Val("&H" & Mid(RegEERPOM(0), 1, 1))
               If (rst(0) >= 5 And rst(0) <= 15) Then
                ComboR1(0).ListIndex = rst(0) - 5
               Else
                ComboR1(0).ListIndex = 11
               End If
                
                ComboR1(1).ListIndex = IIf(Val(rst(1)) And 2 ^ (4 - 1), 1, 0)
                ComboR1(2).ListIndex = IIf(Val(rst(1)) And 2 ^ (3 - 1), 1, 0)
                ComboR1(3).ListIndex = IIf(Val(rst(1)) And 2 ^ (2 - 1), 1, 0)
                ComboR1(4).ListIndex = IIf(Val(rst(1)) And 2 ^ (1 - 1), 1, 0)
                
               rst(0) = Val("&H" & Mid(RegEERPOM(1), 2, 1))
               rst(1) = Val("&H" & Mid(RegEERPOM(1), 1, 1))
                
               
                ComboR1(5).ListIndex = IIf(Val(rst(1)) And 2 ^ (1 - 1), 1, 0)
                ComboR1(6).ListIndex = IIf(Val(rst(1)) And 2 ^ (4 - 1), 1, 0)
                ComboR1(11).ListIndex = IIf(Val(rst(1)) And 2 ^ (2 - 1), 1, 0)
               
                 ComboR1(8).ListIndex = Fix(Fix(rst(0) / 4))
                 ComboR1(9).ListIndex = IIf(Val(rst(0)) And 2 ^ (2 - 1), 1, 0)
                 ComboR1(10).ListIndex = IIf(Val(rst(0)) And 2 ^ (1 - 1), 1, 0)
            rst(0) = Val("&H" & Mid(RegEERPOM(2), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(2), 1, 1))
                
                 ComboR1(7).ListIndex = Fix(rst(0) / 4)
          
                ComboR1(13).ListIndex = rst(1)
                kkk = rst(0) Mod 4
                kkk = kkk * 16
            rst(0) = Val("&H" & Mid(RegEERPOM(3), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(3), 1, 1))
                kkk = kkk + rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(1).Text = kkk * 5
                
            rst(0) = Val("&H" & Mid(RegEERPOM(4), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(4), 1, 1))
                ComboR1(14).ListIndex = rst(1)
                kkk = rst(0) Mod 4
                kkk = kkk * 16
            rst(0) = Val("&H" & Mid(RegEERPOM(5), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(5), 1, 1))
                kkk = kkk + rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(2).Text = kkk * 5
            rst(0) = Val("&H" & Mid(RegEERPOM(6), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(6), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(5).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(7), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(7), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(4).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(8), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(8), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(3).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(9), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(9), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(6).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(10), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(10), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(7).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(11), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(11), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(0).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(12), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(12), 1, 1))
             ComboR1(16).ListIndex = rst(1)
             ComboR1(17).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(13), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(13), 1, 1))
             ComboR1(18).ListIndex = rst(1)
             ComboR1(19).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(14), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(14), 1, 1))
             ComboR1(20).ListIndex = rst(1)
             ComboR1(21).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(15), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(15), 1, 1))
             ComboR1(22).ListIndex = rst(1)
             ComboR1(23).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(16), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(16), 1, 1))
             ComboR1(26).ListIndex = Fix(rst(1) / 4)
             ComboR1(24).ListIndex = rst(1) Mod 4
              ComboR1(25).ListIndex = Fix(rst(0) / 4)
              ComboR1(12).ListIndex = rst(0) Mod 4
              
            rst(0) = Val("&H" & Mid(RegEERPOM(17), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(17), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(8).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(18), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(18), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(9).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(19), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(19), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(10).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(20), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(20), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(11).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(21), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(21), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(12).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(22), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(22), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(13).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(23), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(23), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(14).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(24), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(24), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(15).Text = kkk - 40
End Function
Public Function PrintfThebackup() ' 显示 读取到的寄存器
    Dim I As Integer
    Dim j, s, kkk As Long
    Dim bith, bitl  As Long
    Dim rst(8) As Byte
               rst(0) = Val("&H" & Mid(RegEERPOM(0), 2, 1))
               rst(1) = Val("&H" & Mid(RegEERPOM(0), 1, 1))
               If (rst(0) >= 5 And rst(0) <= 16) Then
                ComboR1(0).ListIndex = rst(0) - 5
               Else
                ComboR1(0).ListIndex = 1
               End If
                ComboR1(1).ListIndex = IIf(Val(rst(1)) And 2 ^ (4 - 1), 1, 0)
                ComboR1(2).ListIndex = IIf(Val(rst(1)) And 2 ^ (3 - 1), 1, 0)
                ComboR1(3).ListIndex = IIf(Val(rst(1)) And 2 ^ (2 - 1), 1, 0)
                ComboR1(4).ListIndex = IIf(Val(rst(1)) And 2 ^ (1 - 1), 1, 0)

               rst(0) = Val("&H" & Mid(RegEERPOM(1), 2, 1))
               rst(1) = Val("&H" & Mid(RegEERPOM(1), 1, 1))
                
               
                ComboR1(5).ListIndex = IIf(Val(rst(1)) And 2 ^ (1 - 1), 1, 0)
                ComboR1(6).ListIndex = IIf(Val(rst(1)) And 2 ^ (4 - 1), 1, 0)
                ComboR1(11).ListIndex = IIf(Val(rst(1)) And 2 ^ (2 - 1), 1, 0)
               
                 ComboR1(8).ListIndex = Fix(Fix(rst(0) / 4))
                 ComboR1(9).ListIndex = IIf(Val(rst(0)) And 2 ^ (2 - 1), 1, 0)
                 ComboR1(10).ListIndex = IIf(Val(rst(0)) And 2 ^ (1 - 1), 1, 0)
            rst(0) = Val("&H" & Mid(RegEERPOM(2), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(2), 1, 1))
                
                 ComboR1(7).ListIndex = Fix(rst(0) / 4)
          
                ComboR1(13).ListIndex = rst(1)
                kkk = rst(0) Mod 4
                kkk = kkk * 16
            rst(0) = Val("&H" & Mid(RegEERPOM(3), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(3), 1, 1))
                kkk = kkk + rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(1).Text = kkk * 5
                
            rst(0) = Val("&H" & Mid(RegEERPOM(4), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(4), 1, 1))
                ComboR1(14).ListIndex = rst(1)
                kkk = rst(0) Mod 4
                kkk = kkk * 16
            rst(0) = Val("&H" & Mid(RegEERPOM(5), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(5), 1, 1))
                kkk = kkk + rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(2).Text = kkk * 5
            rst(0) = Val("&H" & Mid(RegEERPOM(6), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(6), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(5).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(7), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(7), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(4).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(8), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(8), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(3).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(9), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(9), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(6).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(10), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(10), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(7).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(11), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(11), 1, 1))
                kkk = rst(1)
                kkk = kkk * 16
                kkk = kkk + rst(0)
                Text1(0).Text = kkk * 20
            rst(0) = Val("&H" & Mid(RegEERPOM(12), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(12), 1, 1))
             ComboR1(16).ListIndex = rst(1)
             ComboR1(17).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(13), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(13), 1, 1))
             ComboR1(18).ListIndex = rst(1)
             ComboR1(19).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(14), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(14), 1, 1))
             ComboR1(20).ListIndex = rst(1)
             ComboR1(21).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(15), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(15), 1, 1))
             ComboR1(22).ListIndex = rst(1)
             ComboR1(23).ListIndex = rst(0)
            rst(0) = Val("&H" & Mid(RegEERPOM(16), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(16), 1, 1))
             ComboR1(26).ListIndex = Fix(rst(1) / 4)
             ComboR1(24).ListIndex = rst(1) Mod 4
              ComboR1(25).ListIndex = Fix(rst(0) / 4)
              ComboR1(12).ListIndex = rst(0) Mod 4
              
            rst(0) = Val("&H" & Mid(RegEERPOM(17), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(17), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(8).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(18), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(18), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(9).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(19), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(19), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(10).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(20), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(20), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(11).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(21), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(21), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(12).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(22), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(22), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(13).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(23), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(23), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(14).Text = kkk - 40
            rst(0) = Val("&H" & Mid(RegEERPOM(24), 2, 1))
            rst(1) = Val("&H" & Mid(RegEERPOM(24), 1, 1))
            kkk = rst(1)
            kkk = kkk * 16
            kkk = kkk + rst(0)
            Text1(15).Text = kkk - 40
End Function
Private Function dis_temp_num_set()
Dim X As Integer

For X = 0 To 5
    If McuV82SysConfig.TempsetNum And (2 ^ X) Then
        templab(X).Caption = "●"
    Else
        templab(X).Caption = ""
    End If
Next X
End Function

' 将系统配置 显示在界面
Private Function dis_McuV82SysConfig()
Dim I As Integer
Call dis_temp_num_set
I = I + 0: Textsys(I) = byte_to_hex(McuV82SysConfig.Addr)
I = I + 1: Textsys(I) = McuV82SysConfig.CellNum
I = I + 1: Textsys(I) = byte_to_hex(McuV82SysConfig.TempsetNum)
I = I + 1: Textsys(I) = McuV82SysConfig.EngDesign / 10
I = I + 1: Textsys(I) = McuV82SysConfig.BalanceCur           '  //"均衡启动最小充电电流(mA)"    原来这个不用    采样电阻大小    0_01mR）
I = I + 1: Textsys(I) = McuV82SysConfig.BalanceDelay         '    //  软件均衡延时(S）    原来这个不用    参考电压    mv  10
If McuV82SysConfig.B_Mode = 0 Then
    I = I + 1: Textsys(I) = "不均衡"
End If
If McuV82SysConfig.B_Mode = 1 Then
    I = I + 1: Textsys(I) = "充电均衡"
End If
If McuV82SysConfig.B_Mode = 2 Then
    I = I + 1: Textsys(I) = "充电+静态均衡"
End If
I = I + 1: Textsys(I) = McuV82SysConfig.B_THDIS - 40
I = I + 1: Textsys(I) = McuV82SysConfig.B_TLDIS - 40
I = I + 1: Textsys(I) = McuV82SysConfig.B_VStart
I = I + 1: Textsys(I) = McuV82SysConfig.B_Vdiff
I = I + 1: Textsys(I) = McuV82SysConfig.W_Vcell_H
I = I + 1: Textsys(I) = McuV82SysConfig.W_VCell_L
I = I + 1: Textsys(I) = McuV82SysConfig.W_VBAT_H '* 2 / McuV82SysConfig.CellNum
I = I + 1: Textsys(I) = McuV82SysConfig.W_VBAT_L '* 2 / McuV82SysConfig.CellNum
I = I + 1: Textsys(I) = McuV82SysConfig.W_Tcell_H - 40
I = I + 1: Textsys(I) = McuV82SysConfig.W_Tcell_L - 40
I = I + 1: Textsys(I) = McuV82SysConfig.W_Tenv_H - 40
I = I + 1: Textsys(I) = McuV82SysConfig.W_Tenv_L - 40
I = I + 1: Textsys(I) = McuV82SysConfig.W_Tfet_H - 40
I = I + 1: Textsys(I) = McuV82SysConfig.W_Tfet_L - 40

I = I + 1: Textsys(I) = McuV82SysConfig.W_CURR_C / 100
I = I + 1: Textsys(I) = McuV82SysConfig.W_CURR_D / 100
I = I + 1: Textsys(I) = McuV82SysConfig.W_VDIFF_H
I = I + 1: Textsys(I) = McuV82SysConfig.W_VDIFF_L

I = I + 1: Textsys(I) = McuV82SysConfig.OVPVal
I = I + 1: Textsys(I) = McuV82SysConfig.OVPDly
I = I + 1: Textsys(I) = McuV82SysConfig.OVPRel
I = I + 1: Textsys(I) = McuV82SysConfig.UVPVal

I = I + 1: Textsys(I) = McuV82SysConfig.UVPDly
I = I + 1: Textsys(I) = McuV82SysConfig.UVPRel
I = I + 1: Textsys(I) = McuV82SysConfig.BOVPVal '* 2 / McuV82SysConfig.CellNum
I = I + 1: Textsys(I) = McuV82SysConfig.BOVPDly

I = I + 1: Textsys(I) = McuV82SysConfig.BOVPRel ' * 2 / McuV82SysConfig.CellNum
I = I + 1: Textsys(I) = McuV82SysConfig.BUVPVal '* 2 / McuV82SysConfig.CellNum
I = I + 1: Textsys(I) = McuV82SysConfig.BUVPDly
I = I + 1: Textsys(I) = McuV82SysConfig.BUVPRel '* 2 / McuV82SysConfig.CellNum

I = I + 1: Textsys(I) = McuV82SysConfig.CTcellHPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.CTcellHRel - 40
I = I + 1: Textsys(I) = McuV82SysConfig.CTcellLPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.CTcellLRel - 40

I = I + 1: Textsys(I) = McuV82SysConfig.DTcellHPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.DTcellHRel - 40
I = I + 1: Textsys(I) = McuV82SysConfig.DTcellLPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.DTcellLRel - 40

I = I + 1: Textsys(I) = McuV82SysConfig.TenvHPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.TenvHRel - 40
I = I + 1: Textsys(I) = McuV82SysConfig.TenvLPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.TenvLRel - 40

I = I + 1: Textsys(I) = McuV82SysConfig.TfetHPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.TfetHRel - 40
I = I + 1: Textsys(I) = McuV82SysConfig.TfetLPro - 40
I = I + 1: Textsys(I) = McuV82SysConfig.TfetLRel - 40

I = I + 1: Textsys(I) = McuV82SysConfig.CC_PRO_VAL / 100
I = I + 1: Textsys(I) = McuV82SysConfig.CC_PRO_PDLY
I = I + 1: Textsys(I) = McuV82SysConfig.CC_PRO_RDLY
I = I + 1: Textsys(I) = McuV82SysConfig.CC_PRO_LOCK
I = I + 1: Textsys(I) = McuV82SysConfig.CD1_PRO_VAL / 100
I = I + 1: Textsys(I) = McuV82SysConfig.CD1_PRO_PDLY
I = I + 1: Textsys(I) = McuV82SysConfig.CD1_PRO_RDLY
I = I + 1: Textsys(I) = McuV82SysConfig.CD1_PRO_LOCK
I = I + 1: Textsys(I) = McuV82SysConfig.CD2_PRO_VAL / 100
I = I + 1: Textsys(I) = McuV82SysConfig.CD2_PRO_PDLY
I = I + 1: Textsys(I) = McuV82SysConfig.CD2_PRO_RDLY
I = I + 1: Textsys(I) = McuV82SysConfig.CD2_PRO_LOCK
I = I + 1: Textsys(I) = McuV82SysConfig.SHORT_RDLY
I = I + 1: Textsys(I) = McuV82SysConfig.SHORT_LOCK
If McuV82SysConfig.HEAT_EN Then
    I = I + 1: Textsys(I) = "使能"
Else
    I = I + 1: Textsys(I) = "不使能"
End If
I = I + 1: Textsys(I) = McuV82SysConfig.HEAT_TSTART - 40
I = I + 1: Textsys(I) = McuV82SysConfig.HEAT_TEND - 40
End Function
Public Function ReadTheRegchang() ' 读取 界面参数到后台 寄存器 用户修改值
  Dim rst(5) As Long
  Dim kkk As Long
  Dim VOLTE(7) As Long
  Dim temp(7) As Long
  If FrameREG.Visible = True Then
  
    If ComboR1(0).ListIndex <= 10 Then
    rst(0) = ComboR1(0).ListIndex + 5
    Else
    rst(0) = 0
    End If
    
    rst(1) = ComboR1(1).ListIndex * 2 ^ (4 - 1)
    rst(1) = rst(1) + ComboR1(2).ListIndex * 2 ^ (3 - 1)
    rst(1) = rst(1) + ComboR1(3).ListIndex * 2 ^ (2 - 1)
    rst(1) = rst(1) + ComboR1(4).ListIndex * 2 ^ (1 - 1)
    
    RegEERPOM(0) = Hex(rst(1)) & Hex(rst(0))
    rst(1) = ComboR1(5).ListIndex * 2 ^ (1 - 1)
    rst(1) = rst(1) + ComboR1(6).ListIndex * 2 ^ (4 - 1)
    rst(1) = rst(1) + ComboR1(11).ListIndex * 2 ^ (3 - 1)
    rst(0) = ComboR1(8).ListIndex * 4
    rst(0) = rst(0) + ComboR1(9).ListIndex * 2 ^ (2 - 1)
    rst(0) = rst(0) + ComboR1(10).ListIndex * 2 ^ (1 - 1)
    RegEERPOM(1) = Hex(rst(1)) & Hex(rst(0))
    
    VOLTE(0) = Fix(Text1(0).Text)   '二次过压保护电压
    VOLTE(1) = Fix(Val(Text1(1).Text))   '一级过压保护电压
    VOLTE(2) = Fix(Val(Text1(2).Text))    '一级过压恢复电压
    VOLTE(3) = Fix(Text1(4).Text)     '欠压保护恢复电压
    VOLTE(4) = Fix(Text1(5).Text)     '欠压保护电压
    VOLTE(5) = Fix(Text1(6).Text)    '预充开启电压
    VOLTE(6) = Fix(Text1(7).Text)   ' 低压禁止充电电压

   Text4.Visible = False
    If VOLTE(0) <= VOLTE(1) Then
        Text4.Text = "二级过压 > 一级过压 > 过压恢复"
        Text4.Visible = True
    End If
    If VOLTE(1) <= VOLTE(2) Then
        Text4.Text = "一级过压 > 过压恢复 > 欠压恢复 "
        Text4.Visible = True
    End If
    If VOLTE(2) <= VOLTE(3) Then
        Text4.Text = "过压恢复 > 欠压恢复 > 欠压保护  "
        Text4.Visible = True
    End If
    If VOLTE(3) <= VOLTE(4) Then
        Text4.Text = "欠压恢复 > 欠压保护 > 预充开启"
        Text4.Visible = True
    End If
    If VOLTE(4) <= VOLTE(5) Then
        Text4.Text = "欠压保护 > 预充开启 > 低压禁止"
        Text4.Visible = True
    End If
    If VOLTE(5) <= VOLTE(6) Then
        Text4.Text = "欠压保护 > 预充开启 > 低压禁止"
        Text4.Visible = True
    End If
    
kkk = Fix(VOLTE(1) / 5) '一级过压保护电压
    rst(1) = ComboR1(13).ListIndex
    rst(0) = ComboR1(7).ListIndex * 4
    rst(0) = rst(0) + Fix(kkk / 256)
    RegEERPOM(2) = Hex(rst(1)) & Hex(rst(0))
    
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(3) = Hex(rst(1)) & Hex(rst(0))
    
    rst(1) = ComboR1(14).ListIndex
kkk = Fix(VOLTE(2) / 5) '一级过压恢复电压
    rst(0) = Fix(kkk / 256) Mod 4
    RegEERPOM(4) = Hex(rst(1)) & Hex(rst(0))
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(5) = Hex(rst(1)) & Hex(rst(0))
kkk = Fix(VOLTE(4) / 20) '欠压保护电压
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(6) = Hex(rst(1)) & Hex(rst(0))
kkk = Fix(VOLTE(3) / 20) '欠压保护恢复电压
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(7) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(3).Text / 20  '平衡开启电压
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(8) = Hex(rst(1)) & Hex(rst(0))
kkk = Fix(VOLTE(5) / 20) '预充开启电压
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(9) = Hex(rst(1)) & Hex(rst(0))
kkk = Fix(VOLTE(6) / 20) ' 低压禁止充电电压
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(10) = Hex(rst(1)) & Hex(rst(0))
kkk = Fix(VOLTE(0) / 20) '二次过压保护电压
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(11) = Hex(rst(1)) & Hex(rst(0))
    
    rst(1) = ComboR1(16).ListIndex
    rst(0) = ComboR1(17).ListIndex
    RegEERPOM(12) = Hex(rst(1)) & Hex(rst(0))
    rst(1) = ComboR1(18).ListIndex
    rst(0) = ComboR1(19).ListIndex
    RegEERPOM(13) = Hex(rst(1)) & Hex(rst(0))
    rst(1) = ComboR1(20).ListIndex
    rst(0) = ComboR1(21).ListIndex
    RegEERPOM(14) = Hex(rst(1)) & Hex(rst(0))
    rst(1) = ComboR1(22).ListIndex
    rst(0) = ComboR1(23).ListIndex
    RegEERPOM(15) = Hex(rst(1)) & Hex(rst(0))
    rst(1) = ComboR1(26).ListIndex * 4
    rst(1) = rst(1) + ComboR1(24).ListIndex
    rst(0) = ComboR1(25).ListIndex * 4
    rst(0) = rst(0) + ComboR1(12).ListIndex
    RegEERPOM(16) = Hex(rst(1)) & Hex(rst(0))
    
temp(0) = Fix(Text1(8).Text)
temp(1) = Fix(Text1(9).Text)
temp(2) = Fix(Text1(11).Text)
temp(3) = Fix(Text1(10).Text)
temp(4) = Fix(Text1(12).Text)
temp(5) = Fix(Text1(13).Text)
temp(6) = Fix(Text1(15).Text)
temp(7) = Fix(Text1(14).Text)
     Text5.Visible = False
    If temp(0) <= (temp(1) + 4) Then
        Text5.Text = "充电高温 必须> 充电高温恢复高+4℃"
        Text5.Visible = True
    End If
 
    If temp(2) <= (temp(3) + 2) Then
        Text5.Text = "充电低温恢复 必须 > 充电低温高+2℃"
        Text5.Visible = True
    End If
    
    If temp(4) <= (temp(5) + 4) Then
        Text5.Text = "放电高温 必须> 放电高温恢复高+4℃"
        Text5.Visible = True
    End If
 
    If temp(6) <= (temp(7) + 2) Then
        Text5.Text = "放电低温恢复 必须> 放电低温高+2℃ "
        Text5.Visible = True
    End If
 
If Text5.Visible = False Then
    If Text4.Visible = False Then
        CmdRegSent.Enabled = True
    Else
        CmdRegSent.Enabled = False
    End If
Else
    CmdRegSent.Enabled = False
End If
kkk = Text1(8).Text + 40
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(17) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(9).Text
    kkk = kkk Mod 256 + 40
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(18) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(10).Text
    kkk = kkk Mod 256 + 40
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(19) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(11).Text + 40
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(20) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(12).Text + 40
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(21) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(13).Text + 40
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(22) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(14).Text + 40
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(23) = Hex(rst(1)) & Hex(rst(0))
kkk = Text1(15).Text + 40
    kkk = kkk Mod 256
    rst(1) = Fix(kkk / 16)
    rst(0) = kkk Mod 16
    RegEERPOM(24) = Hex(rst(1)) & Hex(rst(0))
  End If
End Function

Public Function label_load()
Dim vstate(Von), Cstate(Con), Tstate(Ton), Alarm(Aon), Fetstate(Fon), Batstate(Gon) As String
Dim I, j As Integer
vstate(0) = "单体过压"
vstate(1) = "单体欠压 "
vstate(2) = "电池组过压 "
vstate(3) = "电池组欠压"
vstate(4) = "单体过压警告 "
vstate(5) = "单体欠压警告 "
vstate(6) = "电池组过压警告"
vstate(7) = "电池组欠压警告"
vstate(8) = "压差保护 "
vstate(9) = "MTP写EEPORM错误"
vstate(10) = "低压、禁止充电 "
Cstate(0) = "充电状态"
Cstate(1) = "放电状态"
Cstate(2) = "充电过流"
Cstate(3) = "短路保护状态"
Cstate(4) = "放电过流1级保护"
Cstate(5) = "放电过流2级保护"
Cstate(6) = "充电电流警告"
Cstate(7) = "放电电流警告"
Tstate(0) = "充电高温"
Tstate(1) = "充电低温"
Tstate(2) = "放电高温"
Tstate(3) = "放电低温"
Tstate(4) = "环境高温"
Tstate(5) = "环境低温"
Tstate(6) = "功率高温"
Tstate(7) = "功率低温"
Tstate(8) = "电芯高温警告"
Tstate(9) = "电芯低温警告"
Tstate(10) = "环境高温警告"
Tstate(11) = "环境低温警告"
Tstate(12) = "功率高温警告"
Tstate(13) = "功率低温警告"
 
Alarm(0) = "继线保护"
Alarm(1) = "充电FET损坏报警"
Alarm(2) = "FLASH 错误"
Alarm(3) = "AFE通讯错误"
Alarm(4) = "存储备份错误"
Alarm(5) = "预留"
Alarm(6) = "充电学习状态"
Alarm(7) = "放电学习状态"

Fetstate(0) = "放电MOS：关闭"
Fetstate(1) = "充电MOS：关闭"
Fetstate(2) = "放电开关：关闭"
Fetstate(3) = "充电开关：关闭"
Fetstate(4) = "放电MOS损坏"
Fetstate(5) = "充电MOS损坏"
Fetstate(6) = "预充MOS：关闭"
Fetstate(7) = "手动控制关闭"
Batstate(0) = "OV"
Batstate(1) = "UV"
Batstate(2) = "OCD1"
Batstate(3) = "OCD2"
Batstate(4) = "OCC"
Batstate(5) = "SC"
Batstate(6) = "PF"
Batstate(7) = "ONEMISS"
Batstate(8) = "UTC"
Batstate(9) = "OTC"
Batstate(10) = "UTD"
Batstate(11) = "OPT"
 For I = 1 To 16
    Load LabelBitV(I)
    Load LabelBitC(I)
    Load LabelBitT(I)
    Load LabelBitA(I)
    Load LabelBitF(I)
    Load LabelBitG(I)
 Next I
  For I = 0 To 16
    LabelBitV(I).ForeColor = &H8000000C
    LabelBitC(I).ForeColor = &H8000000C
    LabelBitT(I).ForeColor = &H8000000C
    LabelBitA(I).ForeColor = &H8000000C
    LabelBitF(I).ForeColor = &H8000000C
    LabelBitG(I).ForeColor = &H8000000C
 Next I
 For I = 0 To Von - 1
    LabelBitV(I).Visible = True
    LabelBitV(I).Caption = vstate(I)
    LabelBitV(I).Height = 320
    LabelBitV(I).Width = 2400
    LabelBitV(I).Left = 120
    LabelBitV(I).Top = 230 + I * (320 + 40)
 Next I
 For I = 0 To Con - 1
    LabelBitC(I).Visible = True
    LabelBitC(I).Caption = Cstate(I)
    LabelBitC(I).Height = 320
    LabelBitC(I).Width = 2400
    LabelBitC(I).Left = 120
    LabelBitC(I).Top = 230 + I * (320 + 40)
 Next I
 For I = 0 To Ton - 1
    LabelBitT(I).Visible = True
    LabelBitT(I).Caption = Tstate(I)
    LabelBitT(I).Height = 320
    LabelBitT(I).Width = 2400
    LabelBitT(I).Left = 120
    LabelBitT(I).Top = 230 + I * (320 + 40)
 Next I
  For I = 0 To Aon - 1
    LabelBitA(I).Visible = True
    LabelBitA(I).Caption = Alarm(I)
    LabelBitA(I).Height = 320
    LabelBitA(I).Width = 2400
    LabelBitA(I).Left = 120
    LabelBitA(I).Top = 230 + I * (320 + 40)
 Next I
  For I = 0 To Fon - 1
    LabelBitF(I).Visible = True
    LabelBitF(I).Caption = Fetstate(I)
    LabelBitF(I).Height = 320
    LabelBitF(I).Width = 2400
    LabelBitF(I).Left = 120
    LabelBitF(I).Top = 230 + I * (320 + 40)
 Next I
  For I = 0 To Gon - 1
    LabelBitG(I).Visible = True
    LabelBitG(I).Caption = Batstate(I)
    LabelBitG(I).Height = 320
    LabelBitG(I).Width = 2400
    LabelBitG(I).Left = 120
    LabelBitG(I).Top = 230 + I * (320 + 40)
 Next I
End Function

Private Sub 电流校零_Click()
If mode_bit8 = 22 Then
  mode_bit8 = 0
   电流校零.Caption = "√电流校零"
   电流校零.ForeColor = &HC0&
Else
  mode_bit8 = 22
   电流校零.Caption = "×电流校零"
   电流校零.ForeColor = &HE0E0E0
End If
End Sub

Private Sub 记录擦除_Click()

If mode_bit10 = 22 Then
  mode_bit10 = 0
   记录擦除.Caption = "√记录擦除"
    记录擦除.ForeColor = &HC0&
Else
  mode_bit10 = 22
   记录擦除.Caption = "×记录擦除"
    记录擦除.ForeColor = &HE0E0E0
End If
End Sub

Private Sub 解码_Click()
If mode_bit1 = 22 Then
  mode_bit1 = 0
   解码.Caption = "√解码"
   解码.ForeColor = &HC0&
Else
  mode_bit1 = 22
   解码.ForeColor = &HE0E0E0
   解码.Caption = "×解码"
End If
End Sub

Private Sub 连续_Click()
If mode_bit11 = 22 Then
  mode_bit11 = 0
   连续.Caption = "√多台连续配置"
   连续.ForeColor = &HFF00&
Else
  mode_bit11 = 22
   连续.Caption = "√单台配置"
     连续.ForeColor = &HC00000
End If
End Sub

Private Sub 配置下发开始按钮_Click()
    Dim filelocation As String
    Dim strsss, MyBool As String
    Dim I As Integer
    Dim j, s As Integer
    Dim kkk     As Integer
    Dim slen As Integer
   ' Dim bith, bitl  As Integer
   ' Dim rst(8) As Byte
    
    
    If 配置下发开始按钮.Caption = "开始" Then
        配置下发开始按钮.Caption = "暂停"
        
 
        jingdu1 = 0
        自动配置结果.Text = "开始" + vbCrLf
    Else
        配置下发开始按钮.Caption = "开始"
        GoTo outtr_su
    End If

     If 解码.Caption = "√解码" Then
          '每次开始 读下文件 到内存中，暂停
         I = 0
         j = 0
         s = 0
         strsss = ""
         filelocation = App.Path + "\自动配置用参数\系统参数1"
         If filelocation = "" Then
         Else
        ' input files into text1.text
          Open filelocation For Input As #1
            Do Until EOF(1)
                Line Input #1, strsss
                strsss = Replace(strsss, " ", "")
                s = Len(strsss)
                j = InStr(strsss, "=")
                sysCaption(I) = Mid(strsss, 1, j - 1)
                LabelSYS(I).Caption = sysCaption(I)
                Textsys(I) = Mid(strsss, j + 1, s)
                I = I + 1
            Loop
          Close #1
        End If
     End If
     
     If 硬件配置下发.Caption = "√硬件配置下发" Then
         I = 0
         j = 0
         s = 0
         strsss = ""
         filelocation = App.Path + "\自动配置用参数\硬件配置"
         If filelocation = "" Then
         Else
        ' input files into text1.text
            Open filelocation For Input As #1
            For I = 0 To 25
                Line Input #1, strsss
                strsss = Replace(strsss, " ", "")
                RegEERPOM(I) = Mid(strsss, 1, 2)
            Next I
            Call PrintfTheReg
            Close #1
        End If
     End If
     
     If 系统参数1下发.Caption = "√系统参数1下发" Then
          '每次开始 读下文件 到内存中，暂停
         I = 0
         j = 0
         s = 0
         strsss = ""
         filelocation = App.Path + "\自动配置用参数\系统参数1"
         If filelocation = "" Then
         Else
        ' input files into text1.text
          Open filelocation For Input As #1
            Do Until EOF(1)
                Line Input #1, strsss
                strsss = Replace(strsss, " ", "")
                s = Len(strsss)
                j = InStr(strsss, "=")
                sysCaption(I) = Mid(strsss, 1, j - 1)
                LabelSYS(I).Caption = sysCaption(I)
                Textsys(I) = Mid(strsss, j + 1, s)
                I = I + 1
            Loop
          Close #1
        End If
     End If
     
 Call readtxt_sys2to_printf
     

If CAP下发.Caption = "√CAP下发" Then
     I = 0
     j = 0
     s = 0
     strsss = ""
    filelocation = App.Path + "\自动配置用参数\cap"
    ' input files into text1.text
    If filelocation = "" Then
    Else
        Open filelocation For Input As #1
          I = 0
          Do Until EOF(1)
            Line Input #1, strsss
            For j = 0 To 3
              slen = Len(strsss)
              kkk = InStr(strsss, "$")                   ' 找到第一个空格
              capData(I * 4 + j) = Val(Mid(strsss, 1, kkk - 1)) ' 读出空格前数据
              strsss = Mid(strsss, kkk + 1, slen)            ' 清出这个读出数据 继续
            Next j
            If I < 26 Then
                I = I + 1
            End If
           Loop
        Call PrintfThecap
        Close #1
    End If
 End If
 
If OCV下发.Caption = "√OCV下发" Then
    I = 0
     j = 0
     s = 0
     strsss = ""
     filelocation = App.Path + "\自动配置用参数\SOCOCV"
    ' input files into text1.text
     If filelocation = "" Then
    Else
        Open filelocation For Input As #1
          I = 0
          Do Until EOF(1)
            Line Input #1, strsss
            For j = 0 To 5
              slen = Len(strsss)
              kkk = InStr(strsss, "$")                   ' 找到第一个空格
              SOC_OCVData(I, j) = Mid(strsss, 1, kkk - 1) ' 读出空格前数据
              strsss = Mid(strsss, kkk + 1, slen)            ' 清出这个读出数据 继续
            Next j
            If I < 49 Then
                I = I + 1
            End If
           Loop
        Call PrintfTheSOCOCV
        Close #1
    End If
End If

 
outtr_su:
    
End Sub

Private Sub 时间校正_Click()
If mode_bit7 = 22 Then
  mode_bit7 = 0
   时间校正.Caption = "√时间校正"
   时间校正.ForeColor = &HC0&
Else
  mode_bit7 = 22
   时间校正.Caption = "×时间校正"
   时间校正.ForeColor = &HE0E0E0
End If
End Sub

Private Sub 系统参数2下发_Click()
If mode_bit4 = 22 Then
  mode_bit4 = 0
   系统参数2下发.Caption = "√系统参数2下发"
   系统参数2下发.ForeColor = &HC0&
Else
  mode_bit4 = 22
   系统参数2下发.Caption = "×系统参数2下发"
   系统参数2下发.ForeColor = &HE0E0E0
End If
End Sub

Private Sub 系统参数1下发_Click()
If mode_bit3 = 22 Then
  mode_bit3 = 0
   系统参数1下发.Caption = "√系统参数1下发"
   系统参数1下发.ForeColor = &HC0&
Else
  mode_bit3 = 22
   系统参数1下发.Caption = "×系统参数1下发"
   系统参数1下发.ForeColor = &HE0E0E0
End If

End Sub

Private Sub 硬件配置下发_Click()
If mode_bit2 = 22 Then
  mode_bit2 = 0
   硬件配置下发.Caption = "√硬件配置下发"
   硬件配置下发.ForeColor = &HC0&
Else
  mode_bit2 = 22
  硬件配置下发.ForeColor = &HE0E0E0
   硬件配置下发.Caption = "×硬件配置下发"
End If
End Sub
