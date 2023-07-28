VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23400
   LinkTopic       =   "Form1"
   ScaleHeight     =   12675
   ScaleWidth      =   23400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCfgSPort 
      Caption         =   "串口配置"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出程序"
      Height          =   615
      Left            =   60
      TabIndex        =   6
      Top             =   9720
      Width           =   1455
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
      TabIndex        =   5
      Top             =   0
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
      TabIndex        =   4
      Top             =   1080
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
      TabIndex        =   3
      Top             =   2280
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
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
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
      TabIndex        =   1
      Top             =   4560
      Width           =   1575
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
      TabIndex        =   0
      Top             =   8100
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
