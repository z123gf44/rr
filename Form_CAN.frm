VERSION 5.00
Begin VB.Form Form1_can 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "USB_CAN"
   ClientHeight    =   9690
   ClientLeft      =   7155
   ClientTop       =   2280
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "清空列表"
      Height          =   375
      Left            =   6000
      TabIndex        =   30
      Top             =   3360
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form_CAN.frx":0000
      Left            =   3720
      List            =   "Form_CAN.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   430
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   18
      Text            =   "1800F4EE"
      Top             =   2820
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6480
      Top             =   480
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   720
      TabIndex        =   15
      Text            =   "01 02 03 04 05 06 07 08 "
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "Form_CAN.frx":001C
      Left            =   3000
      List            =   "Form_CAN.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2805
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form_CAN.frx":003A
      Left            =   960
      List            =   "Form_CAN.frx":0044
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "复位CAN"
      Height          =   330
      Left            =   6000
      TabIndex        =   12
      Top             =   3000
      Width           =   1005
   End
   Begin VB.CommandButton Command_startCAN 
      Caption         =   "启动CAN"
      Height          =   330
      Left            =   6000
      TabIndex        =   11
      Top             =   2640
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form_CAN.frx":0058
      Left            =   1200
      List            =   "Form_CAN.frx":0066
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   420
      Width           =   1290
   End
   Begin VB.CommandButton Connect_button 
      Caption         =   "连接"
      Height          =   330
      Left            =   5400
      TabIndex        =   1
      Top             =   430
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Caption         =   " 发送数据帧 "
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   5835
      Begin VB.CommandButton CAN升级按 
         Caption         =   "CAN 升级"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox 发送check 
         Caption         =   "发送"
         Height          =   255
         Left            =   2880
         TabIndex        =   36
         Top             =   1200
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox 接收Check 
         Caption         =   "接收"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   1200
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox Check_qianzhui2 
         Caption         =   "前缀"
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox Check_quanxian1 
         Caption         =   "全显"
         Height          =   255
         Left            =   2880
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "发送"
         Height          =   330
         Left            =   4680
         TabIndex        =   6
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label11 
         Caption         =   "0%"
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
         Height          =   255
         Left            =   1200
         TabIndex        =   38
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "发送格式："
         Height          =   240
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label7 
         Caption         =   "数据："
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label6 
         Caption         =   "帧ID："
         Height          =   195
         Left            =   4200
         TabIndex        =   4
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label5 
         Caption         =   "帧格式："
         Height          =   195
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "帧类型："
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "设备参数"
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6855
      Begin VB.Frame Frame3 
         Caption         =   "初始化CAN参数"
         Height          =   1335
         Left            =   0
         TabIndex        =   20
         Top             =   720
         Width           =   6615
         Begin VB.ComboBox Combo_botelv 
            Height          =   300
            ItemData        =   "Form_CAN.frx":007D
            Left            =   1200
            List            =   "Form_CAN.frx":009F
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   840
            Width           =   2535
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            ItemData        =   "Form_CAN.frx":0106
            Left            =   5280
            List            =   "Form_CAN.frx":0113
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            ItemData        =   "Form_CAN.frx":0135
            Left            =   5280
            List            =   "Form_CAN.frx":0142
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3360
            TabIndex        =   24
            Text            =   "FFFFFFFF"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1200
            TabIndex        =   22
            Text            =   "00000000"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "CAN波特率:"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "模式："
            Height          =   255
            Left            =   4680
            TabIndex        =   27
            Top             =   885
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "滤波方式："
            Height          =   255
            Left            =   4320
            TabIndex        =   25
            Top             =   400
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "屏蔽码：0x"
            Height          =   255
            Left            =   2400
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "验收码：0x"
            Height          =   255
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Label10 
         Caption         =   "CAN通道："
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "设备类型："
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "信息"
      Height          =   6135
      Left            =   0
      TabIndex        =   16
      Top             =   4200
      Width           =   13695
      Begin VB.ListBox List1 
         Height          =   5820
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   13335
      End
   End
End
Attribute VB_Name = "Form1_can"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_connect As Byte
Dim m_cannum As Long
Dim sent_candata As String
Dim sent_can8data As String
 Dim cansentbyte_iap(262) As Byte
 Dim sentiap_num As Integer
  Dim ljzhi As Integer
 
Public Function EnableUI(bEnable As Boolean)
    Label1.Enabled = bEnable
    Label2.Enabled = bEnable
    Label12.Enabled = bEnable
    Label8.Enabled = bEnable
 
    Label13.Enabled = bEnable
    Text2.Enabled = bEnable
    Text3.Enabled = bEnable

    Combo1.Enabled = bEnable
    Combo2.Enabled = bEnable
    Combo6.Enabled = bEnable
    Combo7.Enabled = bEnable
End Function


Private Sub CAN升级按_Click()
    Dim aa As String
    aa = InputBox("输入程序升级密码")
    If aa = "666666" Then
        IAP.Show 1, frmMain
    Else
        My_msgbox "密码错误，再来一次！"
    End If
End Sub

Public Sub Command_startCAN_Click()
 
    If m_connect = 0 Then
        MsgBox ("请先打开端口")
        Exit Sub
    End If
    If VCI_StartCAN(m_devtype, 0, m_cannum) <> 1 Then
        MsgBox ("启动CAN错误")
    Else
        List1.AddItem "启动CAN成功", List1.ListCount
        Command1.Enabled = True
        CAN_ONUSB_flag = True
    End If
 
End Sub
 

Public Sub sentcan_bytedata()
Dim I As Integer
    If m_connect = 0 Then
        Exit Sub
    End If
    
    If CAN_ONUSB_flag = False Then
        Exit Sub
    End If

    
   If sentiap_num = 0 Then
        ljzhi = 0
        For I = 0 To 262
            cansentbyte_iap(I) = sentda_iap(I)
        Next
        sentiap_num = 263
        caniap_completeflag = False
   End If
 
     If sentiap_num > 0 Then
        Call Command1_Click
    Else
        caniap_completeflag = True
   End If
 
End Sub

Public Sub sentcandata(outstr As String)
    If m_connect = 0 Then
        Exit Sub
    End If
    
    If CAN_ONUSB_flag = False Then
        Exit Sub
    End If
 
sent_candata = sent_candata + outstr
    If Len(sent_candata) > 8 Then
       sent_can8data = Mid(sent_candata, 1, 8)
      sent_candata = Mid(sent_candata, 9, Len(sent_candata))
      Call Command1_Click
    Else
        If Len(sent_candata) > 0 Then
            sent_can8data = Mid(sent_candata, 1, Len(sent_candata))
            sent_candata = ""
            Call Command1_Click
            
        End If
    End If

End Sub


 
 

Public Sub Command1_Click()
    If m_connect = 0 Then
        MsgBox ("请先打开端口")
        Exit Sub
    End If
    
    Dim SendType, frameformat, frametype As Byte
    Dim ID As Long
    Dim data(7) As Byte
    Dim frameinfo As VCI_CAN_OBJ
    Dim str, tmpstr As String
    Dim wuqianzhui As String
    SendType = 0
    frameformat = Combo5.ListIndex
    frametype = Combo4.ListIndex
    str = "&H"
    str = str + Text1.Text '发送的ID
    ID = Val(str)
    
If sentiap_num > 0 Then
    If sentiap_num > 8 Then
        For I = 0 To 7
            frameinfo.DataLen = frameinfo.DataLen + 1
            frameinfo.data(I) = cansentbyte_iap(263 - sentiap_num)
            sentiap_num = sentiap_num - 1
        Next
   Else
        For I = 0 To sentiap_num - 1
            frameinfo.DataLen = frameinfo.DataLen + 1
            frameinfo.data(I) = cansentbyte_iap(263 - sentiap_num)
            sentiap_num = sentiap_num - 1
        Next
        caniap_completeflag = True
   End If

Else

    If Len(sent_can8data) > 0 Then
      str = sent_can8data
        For I = 0 To 7
          If Len(str) > 0 Then
            frameinfo.DataLen = frameinfo.DataLen + 1
            frameinfo.data(I) = Asc(Mid(str, 1, 1))
            If Len(str) > 0 Then
              str = Mid(str, 2, Len(str))
            Else
              Exit For
            End If
          End If
        Next
        sent_can8data = ""
        Text4.Text = ""
    Else
      str = Text4.Text '发送的数据
        I = 0
        For I = 0 To 7
          If Len(str) > 1 Then
            frameinfo.DataLen = frameinfo.DataLen + 1
            frameinfo.data(I) = Val("&H" + Mid(str, 1, 2))
            If Len(str) > 2 Then
              str = Mid(str, 3, Len(str))
            Else
            
              Exit For
            End If
          End If
        Next
      Text4.Text = ""
    End If
End If
      

  

    If frameinfo.DataLen = 0 Then
 
      Exit Sub
    End If
     
    frameinfo.ExternFlag = frametype
    frameinfo.RemoteFlag = frameformat
    frameinfo.ID = ID
 
    If VCI_Transmit(m_devtype, 0, m_cannum, frameinfo, 1) <> 1 Then
       ' MsgBox ("发送数据失败")
    Else
        str = "发送数据成功:  "
        tmpstr = "时间标识:" + Format(Now, "hh:mm:ss:SSS") + GetTickCount
        str = str + tmpstr
        tmpstr = "  帧ID:0x" + Hex(frameinfo.ID)
        str = str + tmpstr
        str = str + "  帧格式:"
        If frameinfo.RemoteFlag = 0 Then
            tmpstr = "数据帧 "
        Else
            tmpstr = "远程帧 "
        End If
        str = str + tmpstr
        str = str + "  帧类型:"
        If frameinfo.ExternFlag = 0 Then
            tmpstr = "标准帧 "
        Else
            tmpstr = "扩展帧 "
        End If
        str = str + tmpstr
        
If 发送check.Value = 1 Then
ljzhi = ljzhi + 1
str = str + CStr(ljzhi)
        If Check_qianzhui2.Value = 0 Then
        Else
            List1.AddItem str, List1.ListCount
        End If
 End If
                    
        
If frameinfo.RemoteFlag = 0 Then
    str = "  数据:"
    If frameinfo.DataLen > 8 Then
        frameinfo.DataLen = 8
    End If
    For j = 0 To frameinfo.DataLen - 1
       If sentIAPflag = 1 Then
        tmpstr = Hex(frameinfo.data(j)) + " "
         str = str + tmpstr
        wuqianzhui = wuqianzhui + Hex(frameinfo.data(j))
        
       Else
        tmpstr = Hex(frameinfo.data(j)) + " "
        str = str + tmpstr
        wuqianzhui = wuqianzhui + Chr(frameinfo.data(j))
       End If
        

    Next
    If 发送check.Value = 1 Then
            If Check_qianzhui2.Value = 0 Then
                  List1.AddItem wuqianzhui, List1.ListCount
            Else
               List1.AddItem str, List1.ListCount
            End If
    End If
End If
        List1.ListIndex = List1.ListCount - 1
    End If
    List1.ListIndex = List1.ListCount - 1
End Sub



Public Sub Command3_Click()
    If m_connect = 0 Then
        MsgBox ("请先打开端口")
        Exit Sub
    End If
    If VCI_ResetCAN(m_devtype, 0, m_cannum) <> 1 Then
        MsgBox ("复位CAN错误")
   Else
        List1.AddItem "复位CAN成功", List1.ListCount
        Command1.Enabled = False
        CAN_ONUSB_flag = False
    End If

End Sub

Public Sub Command4_Click()
    Dim I As Integer
    For I = 0 To List1.ListCount - 1
        List1.RemoveItem 0
    Next
        
End Sub
 


 

Public Sub Connect_button_Click()

Dim cannum As Long
    Dim code, mask As Long
    Dim Timing0, Timing1, filtertype, Mode As Byte
    Dim InitConfig As VCI_INIT_CONFIG
    
    If m_connect = 1 Then
        m_connect = 0
        Connect_button.Caption = "连接"
        VCI_CloseDevice m_devtype, 0
        EnableUI True
        Exit Sub
    End If
        
    If Combo1.ListIndex <> -1 And Combo2.ListIndex <> -1 Then
        cannum = Combo2.ListIndex
        filtertype = Combo6.ListIndex + 1
        Mode = Combo7.ListIndex
        code = Val("&H" + Text2.Text)
        mask = Val("&H" + Text3.Text)
        If Combo_botelv.ListIndex = 0 Then '10K
            Timing0 = Val("&H" + "31")
            Timing1 = Val("&H" + "1C")
        End If
        
        If Combo_botelv.ListIndex = 1 Then '50K
            Timing0 = Val("&H" + "09")
            Timing1 = Val("&H" + "1C")
        End If
        
        If Combo_botelv.ListIndex = 2 Then '100K
            Timing0 = Val("&H" + "04")
            Timing1 = Val("&H" + "1C")
        End If
        
        If Combo_botelv.ListIndex = 3 Then '125K
            Timing0 = Val("&H" + "03")
            Timing1 = Val("&H" + "1C")
        End If
        
        If Combo_botelv.ListIndex = 4 Then '200K
            Timing0 = Val("&H" + "81")
            Timing1 = Val("&H" + "FA")
        End If
        
        If Combo_botelv.ListIndex = 5 Then '250K
            Timing0 = Val("&H" + "01")
            Timing1 = Val("&H" + "1C")
        End If
        
        If Combo_botelv.ListIndex = 6 Then '400K
            Timing0 = Val("&H" + "80")
            Timing1 = Val("&H" + "FA")
        End If
        
        If Combo_botelv.ListIndex = 7 Then '500K
            Timing0 = Val("&H" + "00")
            Timing1 = Val("&H" + "1C")
        End If
        
        If Combo_botelv.ListIndex = 6 Then '800K
            Timing0 = Val("&H" + "00")
            Timing1 = Val("&H" + "16")
        End If
        
        If Combo_botelv.ListIndex = 7 Then '1000K
            Timing0 = Val("&H" + "00")
            Timing1 = Val("&H" + "14")
        End If
        InitConfig.AccCode = code
        InitConfig.AccMask = mask
        InitConfig.Filter = filtertype
        InitConfig.Mode = Mode
        InitConfig.Timing0 = Timing0
        InitConfig.Timing1 = Timing1
        If VCI_OpenDevice(m_devtype, 0, 0) <> 1 Then
            MsgBox ("打开设备错误/或拔掉USB CAN再试")
        Else
            If VCI_InitCAN(m_devtype, 0, cannum, InitConfig) = 1 Then
                m_connect = 1
                m_cannum = cannum
                Connect_button.Caption = "断开"
               Call Command_startCAN_Click
            Else
                MsgBox ("初始化CAN错误")
            End If
        End If
    End If
    EnableUI False
    
End Sub

 

 Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Form1_can.Visible = True Then
  Cancel = 1
   Form1_can.Visible = False
 Else
  Cancel = 0
  Form1_can.Visible = False
 End If

 End Sub
 Public Function myInitForm_Load()
 
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
   
    m_devtype = 4 'USB_CAN2类型号
    m_connect = 0
    m_cannum = 0
    Combo1.ListIndex = 1
    Combo2.ListIndex = 0
    Combo4.ListIndex = 1
    Combo5.ListIndex = 0
    Combo6.ListIndex = 0
    Combo7.ListIndex = 0
    Combo_botelv.ListIndex = 5
    EnableUI True
    Command1.Enabled = False
    CAN_ONUSB_flag = False
End Function


Public Sub Combo1_Click()
     
    m_devtype = Combo1.ListIndex + 3
    If m_devtype = 3 Then
        Combo2.RemoveItem 1
        m_cannum = 0
        Combo2.ListIndex = 0
    Else
        Combo2.AddItem "通道2"
    End If
  
End Sub

Public Sub Timer1_Timer()
    Timer1.Enabled = False
    Dim ErrInfo As VCI_ERR_INFO
    Dim idstr, wuqianzhui As String
    If m_connect = 0 Then
        Timer1.Enabled = True
        Exit Sub
    End If
    wuqianzhui = ""
   If caniap_completeflag = False Then
    Call sentcan_bytedata
   End If
   
    sentcandata ("") ' 发送 大于8个的数据
    Dim Length As Long
    Dim frameinfo(49) As VCI_CAN_OBJ
    Dim str, str22 As String
    Dim get_canstr As String
    Length = VCI_Receive(m_devtype, 0, m_cannum, frameinfo(0), 50, 10)
    If Length <= 0 Then
        VCI_ReadErrInfo m_devtype, 0, m_cannum, ErrInfo '注意：如果没有读到数据则必须调用此函数来读取出当前的错误码，
                                                               '千万不能省略这一步（即使你可能不想知道错误码是什么）
        Timer1.Enabled = True
        Exit Sub
    End If
    get_canstr = ""
    For I = 0 To Length - 1
       
        str = "接收到数据帧:  "
        If frameinfo(I).TimeFlag = 0 Then
            tmpstr = "时间标识:无  "
        Else
            tmpstr = "时间标识:0x" + Hex(frameinfo(I).TimeStamp)
        End If
        str = str + tmpstr
        tmpstr = "  帧ID:0x" + Hex(frameinfo(I).ID)
        str = str + tmpstr
        str = str + "  帧格式:"
        If frameinfo(I).RemoteFlag = 0 Then
            tmpstr = "数据帧 "
        Else
            tmpstr = "远程帧 "
        End If
        str = str + tmpstr
        str = str + "  帧类型:"
        If frameinfo(I).ExternFlag = 0 Then
            tmpstr = "标准帧 "
        Else
            tmpstr = "扩展帧 "
        End If
        str22 = str + tmpstr
        'List1.AddItem str, List1.ListCount
        If frameinfo(I).RemoteFlag = 0 Then
            str = "  数据:"
            If frameinfo(I).DataLen > 8 Then
                frameinfo(I).DataLen = 8
            End If
               get_canstr = ""  ' 这里 要清除一下 不然 下一帧会多上一帧数据
            For j = 0 To frameinfo(I).DataLen - 1
                get_canstr = get_canstr + Chr(frameinfo(I).data(j))
                tmpstr = Chr(frameinfo(I).data(j))
                str = str + tmpstr
            Next
            str22 = str22 + str
            idstr = Right("0000" + Hex(frameinfo(I).ID Mod 65536), 4)
If 接收Check.Value = 1 Then
        If Check_quanxian1.Value = 0 Then '  idstr = Right("0000" + Hex(frameinfo(i).ID Mod 65536), 4)
          If idstr = "EEF4" Then '*计算的CRCj当前值转换为十六进制
                If Check_qianzhui2.Value = 0 Then
                    List1.AddItem get_canstr, List1.ListCount 'id  1807 50F4   F4 BMS  EE PC软件 Right("0000" + Hex(frameinfo(I).ID / 65536), 4) +
                Else
                    List1.AddItem str22, List1.ListCount 'id  1807 50F4   F4 BMS  EE PC软件 Right("0000" + Hex(frameinfo(I).ID / 65536), 4) +
                End If
          End If
        Else
               If Check_qianzhui2.Value = 0 Then
                    List1.AddItem get_canstr, List1.ListCount 'id  1807 50F4   F4 BMS  EE PC软件 Right("0000" + Hex(frameinfo(I).ID / 65536), 4) +
                Else
                    List1.AddItem str22, List1.ListCount 'id  1807 50F4   F4 BMS  EE PC软件 Right("0000" + Hex(frameinfo(I).ID / 65536), 4) +
                End If
        End If
End If
            If idstr = "EEF4" Then '*计算的CRCj当前值转换为十六进制
                chuanshu.UART_CAN_deal_getdata (get_canstr)
                get_canstr = ""
            End If
        End If

        List1.ListIndex = List1.ListCount - 1
    Next
    Timer1.Enabled = True
    
            If List1.ListCount > 1000 Then
            For I = 0 To List1.ListCount - 1
                List1.RemoveItem 0
            Next
        End If
End Sub
