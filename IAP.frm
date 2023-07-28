VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form IAP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "QF_BMS升级软件Ver3.5"
   ClientHeight    =   3450
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   17070
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCfgSPort 
      Caption         =   "串口配置"
      Height          =   615
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      TabIndex        =   6
      Top             =   2700
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "下载"
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   16095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   900
         TabIndex        =   2
         Top             =   300
         Width           =   13695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "载入"
         Height          =   495
         Left            =   14700
         TabIndex        =   1
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "文件路径"
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   420
         Width           =   795
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   1080
      TabIndex        =   9
      Top             =   1560
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   1508
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   13320
      TabIndex        =   11
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label_100 
      Caption         =   "0.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   12000
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1M 
      Caption         =   "下载成功"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15060
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "下载成功"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "下载进度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   1500
      Width           =   795
   End
End
Attribute VB_Name = "IAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCfgSPort_Click(Index As Integer)
    chuanshu.Show 1, frmMain
End Sub
Private Sub Command1_Click()
Dim FileName As String
CommonDialog2.ShowOpen
FileName = CommonDialog2.FileName
Open FileName For Binary As #1
    Flen = LOF(1)       '取文本文件字节数
    If Flen > 0 Then    '要长度大于0才行
    
      '  ReDim FileBin(0 To Flen - 1)
        Get #1, , FileBin
        Text1.Text = FileName
        IapCmd = "00"
         jindu100 = 0
         jindu = Flen
         Command2.Enabled = True
    End If
Close #1
End Sub
Private Sub Command2_Click()
    sentIAPflag = 1
    Getringht_sentF = 1
     IAP.Label_100.Caption = "00.0"
    ProgressBar1.Value = 0
    ProgressBar1.Max = Flen / 256 + 4
    IAP_CHONGSHI = 0
    IAP.Label1M.Visible = False
    goto_reset_mcu = 1
    IAP_MCU_START_FLAG = 0
    goto_reset_mcu_into = 0
    Command2.Enabled = False
    manual_time = 5 ' 500ms 发送间隔
   ' Delay_dis_ReadRegAfe = 4  ' 延时处理 回复数据
    NextSentCmd = CMD_ISP_HANDSHAKE
End Sub
 Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
 End Sub

