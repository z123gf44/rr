VERSION 5.00
Begin VB.Form Dialog_bluetooth 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "对话框标题"
   ClientHeight    =   3675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   2760
   End
   Begin VB.TextBox text_input 
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Text            =   "QF20210001"
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "已输入："
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   420
      Left            =   4800
      TabIndex        =   7
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "2.修改主机蓝牙模块名称：a 去掉与BMS蓝牙连接; b 输入名称; c 点确认 。"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "1.修改BMS上蓝牙模块名称：a BMS通讯上;b 输入名称;c 点确认;d 拔掉电脑上的蓝牙;e 等20S; f 重新插上蓝牙模块。"
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "操作说明："
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "输入蓝牙名称："
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Dialog_bluetooth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
    cyInfoTime = 5
     Unload Me
End Sub

 

Private Sub OKButton_Click()
    Dim leddd As Integer
    bluetooth_name = text_input.Text
    leddd = Len(bluetooth_name)
    
    If leddd <= 10 Then
            NextSentCmd = CMD_Blue_name
            manual_time = 5 ' 500ms 发送间隔
            Label5.Visible = True
            Label6.Visible = True
            Label5.Caption = 20
            cyInfoTime = 5000
    End If
            
End Sub

Private Sub text_input_Change()
    Label7.Caption = Len(text_input.Text)
    
     If Len(text_input.Text) > 10 Then
        Label7.BackColor = &HFF&
        OKButton.Enabled = False
     Else
         Label7.BackColor = &HC0C0&
           OKButton.Enabled = True
     End If
   
End Sub

Private Sub Timer1_Timer()
Call text_input_Change
    If Label5.Caption Then
    
        Label5.Caption = Label5.Caption - 1
    Else
        Label5.Visible = False
        Label6.Visible = False
       
    End If
    
    If Label5.Caption = 18 Then
    '    chuanshu.MSComm1.PortOpen = False
          My_msgbox ("请拔掉蓝牙")
    End If
    
End Sub
