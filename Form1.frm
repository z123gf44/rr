VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form chuanshu 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "����"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7185
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdsend 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6000
      Top             =   420
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ˢ�¶˿�"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton ClearTx_Click 
      Caption         =   "���"
      Default         =   -1  'True
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   675
      Left            =   0
      TabIndex        =   8
      Top             =   6240
      Width           =   435
   End
   Begin VB.TextBox TextTx 
      Height          =   375
      Left            =   540
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   6435
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   840
      List            =   "Form1.frx":0002
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":0004
      Left            =   840
      List            =   "Form1.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4800
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Cmdportopen 
      Caption         =   "�򿪶˿�"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1980
      TabIndex        =   9
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "����"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H008080FF&
      Height          =   8820
      Left            =   540
      TabIndex        =   7
      Top             =   1920
      Width           =   6075
   End
   Begin VB.Label Label2 
      Caption         =   "�˿�COM"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "������"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "chuanshu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Flag As Boolean
Private Sub Cmdportopen_Click()
 On Error Resume Next
If Flag = True Then
  MSComm1.CommPort = Trim(Combo2.Text)
  MSComm1.PortOpen = True
  If Err = 8005 Then
    Label5.Caption = " ���� COM" & Trim(Combo2.Text) & "�����ڻ�����Ӧ��ռ�ã�"
     Label5.ForeColor = &HFF&
  ElseIf MSComm1.PortOpen = True Then
    MSComm1.CommPort = Trim(Combo2.Text)
    MSComm1.Settings = chuanshu.Combo1.Text + ",N,8,1"
    MSComm1.PortOpen = True
    'If Err Then        '������
     '   msg = My_msgbox(" ���� COM" & PortValue & " ��Ч�� ", vbOKOnly, "����")
    '    Exit Sub
   ' End If
    MSComm1.InputLen = 0
    MSComm1.RThreshold = 1
    MSComm1.InputMode = comInputModeText
    cmdsend.Enabled = True
    Flag = False
    Cmdportopen.Caption = "�رն˿�"
    Combo2.Enabled = False
    Combo1.Enabled = False
    SentPussFlag = 0
    'Command2.Enabled = False
    Command3.Enabled = False
    mscomm_delay = 10 * 2
    Label5.Caption = "���ڴ򿪳ɹ���"
    Label5.ForeColor = &HC000&
    frmMain.Label_strdis.Caption = "COM" & Trim(Combo2.Text) & ":  " & chuanshu.Combo1.Text & ",N,8,1"
    chuanshu.Visible = False
   ElseIf MSComm1.PortOpen = False Then
     Label5.Caption = "���ڴ�ʧ�ܣ�"
     Label5.ForeColor = &HFF&
   End If
Else
    Flag = True
    Label5.Caption = ""
    Cmdportopen.Caption = "�򿪶˿�"
    cmdsend.Enabled = False
    MSComm1.PortOpen = False
    Combo2.Enabled = True
    Combo1.Enabled = True
    'Command2.Enabled = True
    Command3.Enabled = True
End If
End Sub


Private Sub ClearTx_Click_Click() '//�������
    TextTx.Text = ""
End Sub
Private Sub Command1_Click() '//�������
    Label6.Caption = ""
    Getdis = ""
End Sub

Private Sub getport_zgf()  ' ��ȡ�˿�
    Dim I     As Integer
    I = Combo2.ListCount
    If I > 0 Then
    Do
       I = I - 1
       Combo2.RemoveItem I
    Loop Until I = 0
    End If
    On Error GoTo Err1
Err1:
    If Err = 8005 Then
        Combo2.AddItem I
    End If
    Resume Next
    For I = 1 To 16
        MSComm1.CommPort = I
        MSComm1.PortOpen = True
        If MSComm1.PortOpen = True Then
            MSComm1.PortOpen = False
            Combo2.AddItem I  '��������ʾ����
        End If
    Next
  End Sub
 Public Function myInitForm_Load()
    Call getport_zgf
    Combo1.ListIndex = 5
    cmdsend.Enabled = False
    Flag = True
    Label5.Caption = ""
    Cmdportopen.Caption = "�򿪶˿�"
    'chuanshu.MSComm1.PortOpen = False
    Combo2.Enabled = True
    Combo1.Enabled = True
    Cmdportopen.Enabled = True
    SentPussFlag = 1
    Call Command3_Click
 End Function
Private Sub Command3_Click()
    Command3.Enabled = False
    Call getport_zgf
    Command3.Enabled = True
End Sub

 Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If chuanshu.Visible = True Then
  Cancel = 1
   chuanshu.Visible = False
 Else
  Cancel = 0
  chuanshu.Visible = False
 End If

 End Sub
  
  
 Public Sub UART_CAN_deal_getdata(getstr As String)
 Dim intInputLen As Integer
 
 intInputLen = Len(getstr)
 If GetheairF = False Then
        rxStrall = getstr
        I = InStr(1, rxStrall, ":", 1)
        If I > 0 Then
            rx_cnt = intInputLen - I
            If rx_cnt > 0 Then
                rxStrall = Mid(rxStrall, I + 1, rx_cnt)
            Else
                 rxStrall = ""  '���
            End If
            GetheairF = True '������ȷ
            I = InStr(1, rxStrall, "~", 1)
            If I > 0 Then
            '�������� ��ʼ ��������    ��� ����һ֡�ͽ������ Ҳ�������
          
               GetheairF = False
               If sentIAPflag Then
                  Call dealIAP_GETDATA
               Else
                    If rx_cnt >= 14 Then
                       rxStrall = ":" + rxStrall
                       Call deal_getdata
                    End If
               End If
            End If
        Else
          rxStrall = ""  '���
        End If
    Else ' ���յ� ����ͷ
        strData = getstr
        I = InStr(1, strData, "~", 1)
        rx_cnt = rx_cnt + intInputLen
        rxStrall = rxStrall + strData
        If I > 0 Then
        '�������� ��ʼ ��������
   
           GetheairF = False
           If sentIAPflag Then
              Call dealIAP_GETDATA
           Else
                If rx_cnt >= 14 Then
                   rxStrall = ":" + rxStrall
                   Call deal_getdata
                End If
           End If
        End If
        If rx_cnt > 300 Then
            GetheairF = False
            rx_cnt = 0
            rxStrall = ""
        End If
 End If

 End Sub

Private Sub MSComm1_OnComm()  '��������
Dim I, j, k As Integer
Dim strData, gettemp As String

'  On Error GoTo zgfmscommerr  zgf 2020226
If MSComm1.CommEvent = comEvReceive Then
  
    If mscomm_delay = 0 Then ' ��ʱ һ���ٽ���
        gettemp = chuanshu.MSComm1.Input
            MSComm1.InBufferCount = 0 '��ս��ջ�����
        Getdis = Getdis + gettemp
    End If

   UART_CAN_deal_getdata (gettemp)
    
    Getdis = Getdis + vbCrLf
    If Len(Getdis) > 1024 Then
    Getdis = ""
    End If
   Label6.Caption = Getdis
End If 'If MSComm1.CommEvent = comEvReceive Then
 
End Sub


'  ����������յ�һ�� ��������
Private Function deal_getdata()
  Dim crcstr, crcmuch As String
  Dim abc As String
  Dim cmd As Long
  Dim lenth As Long
   
  CRCADD = 0
   lenth = Len(rxStrall)
  crcmuch = CRC16_keycodedata(Mid(rxStrall, 1, lenth - 5))    '//����� crc �� ����
 
    If lenth < 14 Then
        lenth = 14
    End If
    If lenth > 400 Then
        lenth = 14
    End If
    crcstr = Mid(rxStrall, lenth - 4, 4) ' ��85��
 
 
If crcstr = crcmuch Then  ' 1������ �������� �����Ӧ������  ��2 Ȼ�� �����е�����ʾȥ����
 GET_DATA = 1
 rightback_lj = rightback_lj + 1 'У����ȷ������ +1
 
 
     For I = 1 To (lenth - 2) / 2 ' info  ���� ��00AC  14+79*2  CONFIG ��  14+ 112*2      �о��þ÷��� ��������ǿո� ������Ҫת��һ��
        abc = Mid(rxStrall, I * 2, 1)
        OneH = Val("&H" & abc) * 16
        abc = Mid(rxStrall, I * 2 + 1, 1)
        If abc = "" Then  ' ������Ϊ ʱ�� ���� �ս��ջ��ڷ��� �����ִ����·���Ȼ���м���ֶ����ǿ��ַ������ǿո�
        Else
            one = Val("&H" & abc)
            rec(I - 1) = OneH + one
        End If
    Next I
    
    
    crcmuch = Mid(rxStrall, 4, 2)
     cmd = Val("&H" & crcmuch)
             Select Case cmd
                Case CMD_ReadSN  '0x01 ��ȡSN��
                                
                Case CMD_ReadSOCSOP   '0x02 ��ȡSOC,SOP SOP����%�ȣ������ѹ����ʱ��˻�
                               
                Case CMD_ReadVOLTAGE_CURREN   '0x03 ��ȡ��ѹ������
                                 
                Case CMD_ReadInfo   '0x04 ��ȡʵʱ����
                    Call ReadBatInfo_485toRAM
                Case CMD_ReadSysConfig  '0x05 ��ȡ��������
                    Call ReadSysconfig_485toRAM
                    Delay_dis_Readsysconfig = 100    '�� 100 ��ʾ ���ճɹ�
                Case CMD_ReadBalckUp  '0x06 ��ȡ��������    RD_EEPROM      CMD_ReadBalckUp �� CMD_cmd_No �ɽ��մ����� ����
Call ReadBalckUp_485toRAM
                Case CMD_ReadSys2Config  '0x07 ��ȡ��������    RD_MCUSYSTEM
                    Call ReadSys2config_485toRAM
                    Flag_sys2ok = True
                     If jingdu1 = 1 Then
                        Flag_readmcusys2ok = True
                    End If
                       If jingdu1 = 5 Then
                        Flag_readckeckjiemasys2ok = True
                    End If
                    If jingdu1 > 1 Then
                    Delay_dis_Readsys2config = 0
                    Else ' ���� ��Ҫ�ȴ�
                        Delay_dis_Readsys2config = 100    '�� 100 ��ʾ ���ճɹ�
                    End If
                    
                Case CMD_ReadAFEseg  '0x08 ��ȡ�Ĵ�������  RD_MTP
Call ReadRegAfe_485toRAM
Delay_dis_ReadRegAfe = 100
                Case CMD_ReadRTC   '0x09 ��RTC
        
                Case CMD_ReadMcuRAM   '0x0A ��ȡ�ڲ�״̬
                                
                Case CMD_ReadSOC_OCV   '0x0B ��SOC���ò���    CMD_ReadSOC_OCV �� CMD_cmd_No �Ͷ�ȡ�������� ����һ�����ɽ��պ�������
                    Call ReadSOCOCV_485toRAM
                 Case CMD_Readcap  '0x0B ��SOC���ò���    CMD_ReadSOC_OCV �� CMD_cmd_No �Ͷ�ȡ�������� ����һ�����ɽ��պ�������
                     Call Readcap_485toRAM
                Case CMD_WriteAFEseg  '0x20 ���üĴ�������  WR_MTP
                 If rec(5) Then
                    Flag_sys2ok = True
                     Delay_dis_WriteRegAfe = 100
                  Else
                    Delay_dis_WriteRegAfe = 2
                End If
                Case CMD_SetFET   '0x21 �·�FET����
                ' Delay_dis_SetFET = 100
                Case CMD_WriteSysConfig  '0x22 �·����ò���
                NextSentCmd = CMD_ReadInfo
                 If rec(5) Then
                    Flag_sys2ok = True
                     Delay_dis_Writesysconfig = 100
                  Else
                    Delay_dis_Writesysconfig = 2
                End If
                    
                Case CMD_EraseBalckUp  '0x23 �·�������������
                 If rec(5) Then
                    Flag_sys2ok = True
                     Delay_dis_EraseBalckUp = 100
                  Else
                    Delay_dis_EraseBalckUp = 2
                End If
                     
                Case CMD_CALIB_VOLTAGE  '0x24 �·�У���ܵ�ѹ  CALIB_VOLTAGE
                    If rec(5) Then
                         Delay_dis_CALIB_VOLTAGE = 100
                      Else
                        Delay_dis_CALIB_VOLTAGE = 2
                    End If
                Case CMD_CALIB_CURRENT  '0x25 �·�У������    CALIB_CURRENT    00  xH xL
                 If rec(5) Then
                     Flag_sys2ok = True
                     Delay_dis_CALIB_CURRENT = 100
                  Else
                    Delay_dis_CALIB_CURRENT = 2
                End If
                Case CMD_CALIB_TEMPE  '0x26 �·�У���¶�    CALIB_TEMPE
                    If rec(5) Then
                       Delay_dis_CALIB_TEMPE = 100
                    Else
                      Delay_dis_CALIB_TEMPE = 2
                    End If
 
                Case CMD_CALIB_RTC   '0x27 �·�����RTC CALIB_RTC
                  If rec(5) Then
                        Flag_sys2ok = True
                       Delay_dis_CALIB_RTC = 5
                    Else
                      Delay_dis_CALIB_RTC = 2
                    End If
                Case CMD_Enter_Sleep_Mode  '0x28 �·�BMS����ػ� Enter_Sleep_Mode
                    Delay_dis_Enter_Sleep_Mode = 100
                Case V82_SET_POWERON  '0x28 �·�BMS����ػ� Enter_Sleep_Mode
                    Delay_dis_Enter_WORK_Mode = 100
                Case CMD_ReSet_OFFSET  '0x28 �·�BMS����ػ� Enter_Sleep_Mode
                    Delay_dis_CALIB_TEMPE = 100
                Case CMD_ISP_HANDSHAKE  '0x29 �·�����IAP_����    ISP_HANDSHAKE
                          
                Case CMD_WriteSOC_OCV   '0x30 �·�����SOC���ò���
                    Call deal_WriteSOC_ocv_backtoRAM
                Case CMD_Writecap   '0x30 �·�����SOC���ò���
                    Call deal_Writecap_backtoRAM
                Case CMD_WriteSys2Config   '0x31 �·���������
                
                NextSentCmd = CMD_ReadInfo
                
                If jingdu1 = 5 Or jingdu1 = 4 Then
                    Flag_onlysys2ok = True
                        Delay_dis_Writesys2config = 0
                Else
                    If rec(5) Then
                         Delay_dis_Writesys2config = 100
                      Else
                        Delay_dis_Writesys2config = 2
                    End If
                End If
                
                If jingdu1 = 30 Then
                    Flag_sys2ok = True
                        Delay_dis_Writesys2config = 0
 
                End If

             Case CMD_ActiveBms   '0x31 �·���������
                If jingdu1 = 4 Or jingdu1 = 4 Then
                    Flag_onlysys2ok = True
                        Delay_dis_Writesys2config = 0
                Else
                    If rec(5) Then
                         Delay_dis_Writesys2config = 100
                      Else
                        Delay_dis_Writesys2config = 2
                    End If
                End If
             End Select
Else
 GET_DATA = 0
 If NextSentCmd = CMD_ReadBalckUp Then
     Delay_dis_ReadBalckUp = 20  ' У����� ���·�
    NextSentCmd = CMD_ReadBalckUp
    manual_time = 3 ' 500ms ���ͼ��
 End If

 End If
    Strall = ""
    rx_cnt = 0

End Function

' ��ʾ ���յ��� ��������
'�ظ� 01/��һ����02/�м��¼��03/���һ��
'01                              45 36 14 01 11 05 20 04 00 03 00 00 00 00 00 00 00 00 00 00 00 01 0D 42 0D 3A 0D 2E 0D 39 0D 39 0D 28 0D 36 0D 37 0D 3C 0D 27 0D 6B 0D 3C 0D 31 00 00 00 00 00 00 00 00 AB EC 00 00 00 00 0A AB 0A AB 0A AB 18 19 19 19 19 19 AF
Private Function deal_read_backup_todis()
 Dim I As Long
 
   
End Function
' ��ʾ ���յ���AFEREG����
Private Function ReadRegAfe_485toRAM()
 Dim I As Long
 For I = 0 To 25
  RegEERPOM(I) = byte_to_hex(rec(I + 5))
 Next I
 
  Call frmMain.PrintfTheReg
   
End Function

 '  ��ʾ���յ� ��� ʵʱ����
Private Function ReadBatInfo_485toRAM()
 Dim temp As Long
 Dim ii As Integer
 Dim xiaoshu As Single
 Dim strrr As String
': 01 825200900000000000000048F80A0EA90EB30EB60EB40E8C0EB40E450E9E0E9E0E6A0000000002474500000000000000000F00000000000000000000000000002D004800A04E~
' 00 (addr), 02 (cmd), 00 (ver), 000e (len), e8 (crc), ~ (EOI)
': (SOI),
'01 (addr), 82 (cmd), 52 (ver), 0090 (len), 00000000000000(time_t)
myRealV82Info.Time_t = DateSerial(rec(5), rec(6), rec(7)) & "-" & TimeSerial(rec(8), rec(9), rec(10))      ' Byte ': ʱ��,7bit �ֱ����ꡢ�¡��ա� �ܡ�ʱ���֡���'
temp = rec(12)
temp = temp * 256
temp = temp + rec(11)
myRealV82Info.mcu_powerStatu = temp
temp = rec(14)
temp = temp * 256
temp = temp + rec(13)
temp = temp * 2

myRealV82Info.Vbat = temp / 10    ' ��ص�ѹ�����Ϊ�ܵ�ѹ��0.5��'
myRealV82Info.Vcell_num = rec(15)  ' Byte    ': ��ش�����1-16'
myRealV82Info.RealTempNum = rec(16)  ' Byte    ': �¶Ȳ�������'
If myRealV82Info.Vcell_num >= 16 Then
myRealV82Info.Vcell_num = 16
End If
For I = 0 To 15
temp = rec(18 + I * 2)
temp = temp * 256
temp = temp + rec(17 + I * 2)
myRealV82Info.Vcell(I) = Format(temp / 1000, "0.000") ' Integer   '��ÿһ�ڵ�ѹ mV'
Next I
temp = rec(52)
temp = temp * 256
temp = temp + rec(51)

temp = temp * 256 + rec(50) '
temp = temp * 256
temp = temp + rec(49)
xiaoshu = temp
myRealV82Info.Curr = (xiaoshu - 500000) / 1000

'myRealV82Info.Curr(1) = temp '  ': Curr[0]��������Curr[1]�ŵ����'

For I = 0 To 5
 myRealV82Info.temp(I) = Format(rec(53 + I) - 40, "0.0")   ' Byte  ': ÿ���¶ȵ����ݣ�65 ��ʾ 25�棬��ƫ�� 40' '
Next I

 ii = 0
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.vstate = temp ' Integer ' '
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.Cstate = temp  ' Integer  ' '
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.Tstate = temp  ' Integer    '  '���ݽṹ����
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.Alarm = temp   ' Integer '     '���ݽṹ����
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.Fetstate = temp ' Integer  '' ���ݽṹ����
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.NUM_VOV = temp ' Integer  '�����ѹ��Ӧ�ĵ�ص���ţ����� 5 ��ʾ�� 5 �ڸ�ѹ
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.NUM_VUV = temp ' Integer  ' ������Ƿѹ��Ӧ�ĵ�ص����
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.NUM_WARN_VHIGH = temp ' Integer ' �������ѹ�����Ӧ�ĵ�ص����
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.NUM_WARN_VLOW = temp ' Integer ' �������ѹ�����Ӧ�ĵ�ص����
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.BlanceState = temp ' Integer ' �� ����״̬����ʾ��һ�ڵ�ѹ��������
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.DchgNum = temp ' Integer ' ���ŵ����'
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.BatStatus = temp ' Integer ' ��������'
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.SOC = Format(temp / 10, "0.0")  ' Byte '  : ��� soc ���ٷֱ� 0-1000'  ��ǰSOC(0.1)

 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.CapNow = Format(temp / 10, "0.0") ' Integer ' : ��ǰ���� (0.1AH)
 ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.CapFull = Format(temp / 10, "0.0")  ' Integer ' : ��������(0.1AH)
havegetTRightData = 1

ii = ii + 2
temp = rec(ii + 60)
temp = temp * 256
temp = temp + rec(ii + 59)
myRealV82Info.FET_code = temp  ' Integer ' ƫ�����  (0.001A)

myRealV82Info.afe_Temp(1) = Format(rec(ii + 61) - 40, "0.0")   ' Byte  ': ÿ���¶ȵ����ݣ�65 ��ʾ 25�棬��ƫ�� 40' '
myRealV82Info.afe_Temp(2) = Format(rec(ii + 62) - 40, "0.0")   ' Byte  ': ÿ���¶ȵ����ݣ�65 ��ʾ 25�棬��ƫ�� 40' '
myRealV82Info.afe_Temp(3) = Format(rec(ii + 63) - 40, "0.0")   ' Byte  ': ÿ���¶ȵ����ݣ�65 ��ʾ 25�棬��ƫ�� 40' '

havegetTRightData = 1

strrr = myRealV82Info.Time_t
strrr = strrr & "    " & myRealV82Info.Vbat
strrr = strrr & "    " & myRealV82Info.Curr
strrr = strrr & "    " & myRealV82Info.SOC

For I = 0 To 15
strrr = strrr & "    " & myRealV82Info.Vcell(I)
Next I

For X = 0 To 4
    If myRealV82Info.RealTempNum And (2 ^ X) Then
        strrr = strrr & "    " & myRealV82Info.temp(I)
    End If
Next X

strrr = strrr & "    " & myRealV82Info.vstate
strrr = strrr & "    " & myRealV82Info.Cstate
strrr = strrr & "    " & myRealV82Info.Tstate
strrr = strrr & "    " & myRealV82Info.Alarm
strrr = strrr & "    " & myRealV82Info.Fetstate
strrr = strrr & "    " & myRealV82Info.NUM_VOV
strrr = strrr & "    " & myRealV82Info.NUM_VUV
strrr = strrr & "    " & myRealV82Info.NUM_WARN_VHIGH
strrr = strrr & "    " & myRealV82Info.NUM_WARN_VLOW
strrr = strrr & "    " & myRealV82Info.BlanceState
strrr = strrr & "    " & myRealV82Info.DchgNum
strrr = strrr & "    " & myRealV82Info.BatStatus
strrr = strrr & "    " & myRealV82Info.CapNow
strrr = strrr & "    " & myRealV82Info.CapFull
strrr = strrr & "    " & Int_to_hex(myRealV82Info.FET_code)
strrr = strrr & "    " & myRealV82Info.afe_Temp(1)
strrr = strrr & "    " & myRealV82Info.afe_Temp(2)
strrr = strrr & "    " & myRealV82Info.afe_Temp(3)

  RecordTime_add = RecordTime_add + cyInfoTime '//�÷���ʱ�� �ۼ� ��¼ʱ��
  If frmMain.Labeljilu.Caption = "��" Then
       If RecordTime_add >= RecordTime Then
            RecordTime_add = 0
            Open App.Path & jilu_path & ".txt" For Append As #1
            Print #1, strrr
            Close #1
       End If
   End If
End Function
' ��ʾ ��MCU������ ��������
Private Function ReadSysconfig_485toRAM()
  Dim temp As Long
  Dim I As Integer

    I = 5
McuV82SysConfig.EngDesign = rec(I) + rec(I + 1) * 256: I = I + 2 '"     '"�������,(0_1AH)     '"
McuV82SysConfig.BalanceCur = rec(I) + rec(I + 1) * 256: I = I + 2 '"     '"     '"     '"����������С������(mA)     '"     '"     '"ԭ��������ò��������С0_01mR��
McuV82SysConfig.BalanceDelay = rec(I) + rec(I + 1) * 256: I = I + 2 '"���������ʱ(S��ԭ��������òο���ѹmv10
McuV82SysConfig.B_VStart = rec(I) + rec(I + 1) * 256: I = I + 2 '"����������ѹ��mV��
McuV82SysConfig.B_Vdiff = rec(I) + rec(I + 1) * 256: I = I + 2 '"��������ѹ�mV��10
McuV82SysConfig.W_Vcell_H = rec(I) + rec(I + 1) * 256: I = I + 2 '"���ڸ�ѹ����ֵmv
McuV82SysConfig.W_VCell_L = rec(I) + rec(I + 1) * 256: I = I + 2 '" ���ڵ�ѹ����ֵ
McuV82SysConfig.W_VBAT_H = rec(I) + rec(I + 1) * 256: I = I + 2 '"��ظ�ѹ����ֵ
McuV82SysConfig.W_VBAT_L = rec(I) + rec(I + 1) * 256: I = I + 2 '"��ص�ѹ����ֵ26
McuV82SysConfig.W_CURR_C = rec(I) + rec(I + 1) * 256: I = I + 2 '"����������ֵ0_01A20
McuV82SysConfig.W_CURR_D = rec(I) + rec(I + 1) * 256: I = I + 2 '"�ŵ��������ֵ
McuV82SysConfig.W_VDIFF_H = rec(I) + rec(I + 1) * 256: I = I + 2 '"ѹ���ֵ
McuV82SysConfig.W_VDIFF_L = rec(I) + rec(I + 1) * 256: I = I + 2 '"ѹ������ֵ
McuV82SysConfig.OVPVal = rec(I) + rec(I + 1) * 256: I = I + 2 '"��������ѹ
McuV82SysConfig.OVPDly = rec(I) + rec(I + 1) * 256: I = I + 2 '"������䱣����ʱ30
McuV82SysConfig.OVPRel = rec(I) + rec(I + 1) * 256: I = I + 2 '"�������ָ���ѹ
McuV82SysConfig.UVPVal = rec(I) + rec(I + 1) * 256: I = I + 2 '"������ŵ�ѹ
McuV82SysConfig.UVPDly = rec(I) + rec(I + 1) * 256: I = I + 2 '"������ű�����ʱ
McuV82SysConfig.UVPRel = rec(I) + rec(I + 1) * 256: I = I + 2 '"������Żָ���ѹ
McuV82SysConfig.BOVPVal = rec(I) + rec(I + 1) * 256: I = I + 2 '"�����������ѹ40
McuV82SysConfig.BOVPDly = rec(I) + rec(I + 1) * 256: I = I + 2 '"���������䱣����ʱ
McuV82SysConfig.BOVPRel = rec(I) + rec(I + 1) * 256: I = I + 2 '"����������ָ���ѹ
McuV82SysConfig.BUVPVal = rec(I) + rec(I + 1) * 256: I = I + 2 '"��ع��ŵ�ѹ
McuV82SysConfig.BUVPDly = rec(I) + rec(I + 1) * 256: I = I + 2 '"��ع��ű�����ʱ
McuV82SysConfig.BUVPRel = rec(I) + rec(I + 1) * 256: I = I + 2 '"��ع��Żָ���ѹ50
McuV82SysConfig.CC_PRO_VAL = rec(I) + rec(I + 1) * 256: I = I + 2 '"����������ֵ
McuV82SysConfig.CC_PRO_PDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"������������ʱ
McuV82SysConfig.CC_PRO_RDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"�������ָ���ʱ
McuV82SysConfig.CC_PRO_LOCK = rec(I) + rec(I + 1) * 256: I = I + 2 '"��������������
McuV82SysConfig.CD1_PRO_VAL = rec(I) + rec(I + 1) * 256: I = I + 2 '"һ���ŵ籣��ֵ60
McuV82SysConfig.CD1_PRO_PDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"һ���ŵ����������ʱ
McuV82SysConfig.CD1_PRO_RDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"һ���ŵ�����ָ���ʱ
McuV82SysConfig.CD1_PRO_LOCK = rec(I) + rec(I + 1) * 256: I = I + 2 '"һ���ŵ������������
McuV82SysConfig.CD2_PRO_VAL = rec(I) + rec(I + 1) * 256: I = I + 2 '"�����ŵ籣��ֵ
McuV82SysConfig.CD2_PRO_PDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"�����ŵ����������ʱ70
McuV82SysConfig.CD2_PRO_RDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"�����ŵ�����ָ���ʱ
McuV82SysConfig.CD2_PRO_LOCK = rec(I) + rec(I + 1) * 256: I = I + 2 '"�����ŵ������������
McuV82SysConfig.SHORT_RDLY = rec(I) + rec(I + 1) * 256: I = I + 2 '"��·��ʱֵ
McuV82SysConfig.SHORT_LOCK = rec(I) + rec(I + 1) * 256: I = I + 2 '"��·����ֵ
McuV82SysConfig.CTcellHPro = rec(I): I = I + 1 '"��о�����±���
McuV82SysConfig.CTcellHRel = rec(I): I = I + 1 '"��о�����±����ָ�80
McuV82SysConfig.CTcellLPro = rec(I): I = I + 1 '"��о�����±���
McuV82SysConfig.CTcellLRel = rec(I): I = I + 1 '"��о�����±����ָ�
McuV82SysConfig.DTcellHPro = rec(I): I = I + 1 '"��о�ŵ���±���
McuV82SysConfig.DTcellHRel = rec(I): I = I + 1 '"��о�ŵ���±����ָ�
McuV82SysConfig.DTcellLPro = rec(I): I = I + 1 '"��о�ŵ���±���85
McuV82SysConfig.DTcellLRel = rec(I): I = I + 1 '"��о�ŵ���±����ָ�
McuV82SysConfig.TenvHPro = rec(I): I = I + 1 '"��о�������±���
McuV82SysConfig.TenvHRel = rec(I): I = I + 1 '"��о�������±����ָ�
McuV82SysConfig.TenvLPro = rec(I): I = I + 1 '"��о�������±���
McuV82SysConfig.TenvLRel = rec(I): I = I + 1 '"��о�������±����ָ�90
McuV82SysConfig.TfetHPro = rec(I): I = I + 1 '"��о���ʸ��±���
McuV82SysConfig.TfetHRel = rec(I): I = I + 1 '"��о���ʸ��±����ָ�
McuV82SysConfig.TfetLPro = rec(I): I = I + 1 '"��о���ʵ��±���
McuV82SysConfig.TfetLRel = rec(I): I = I + 1 '"��о���ʵ��±����ָ�
McuV82SysConfig.W_Tcell_H = rec(I): I = I + 1 '"��о���¾���ֵ95
McuV82SysConfig.W_Tcell_L = rec(I): I = I + 1 '" ��о���¾���ֵ
McuV82SysConfig.W_Tenv_H = rec(I): I = I + 1 '"�������¾���ֵ
McuV82SysConfig.W_Tenv_L = rec(I): I = I + 1 '"�������¾���ֵ
McuV82SysConfig.W_Tfet_H = rec(I): I = I + 1 '"���ʸ��¾���ֵ
McuV82SysConfig.W_Tfet_L = rec(I): I = I + 1 '"���ʵ��¾���ֵ100
McuV82SysConfig.B_Mode = rec(I): I = I + 1 '"����ģʽ0~2��0 ������1������2���+��̬����
McuV82SysConfig.B_THDIS = rec(I): I = I + 1 '"������½�ֵֹ40��ʾ0��65��ʾ25��
McuV82SysConfig.B_TLDIS = rec(I): I = I + 1 '"������½�ֵֹ
McuV82SysConfig.Addr = rec(I): I = I + 1 '"������RS485��ַ1~255
McuV82SysConfig.CellNum = rec(I): I = I + 1 '"��ؽ���5~16105
McuV82SysConfig.TempsetNum = rec(I): I = I + 1   '"����ʹ��4
McuV82SysConfig.SHORT_VAL = rec(I): I = I + 1 '"��·��ѹ����ֵ
McuV82SysConfig.HEAT_EN = rec(I): I = I + 1 '"���ȹ���ʹ��
McuV82SysConfig.HEAT_TSTART = rec(I): I = I + 1 '"���ȿ����¶�
McuV82SysConfig.HEAT_TEND = rec(I): I = I + 1 '"���ȹر��¶�110


   
End Function
' ��ʾ ��MCU������ ��������
Private Function ReadSys2config_485toRAM()
  Dim temp As Long
  Dim sstr  As String
  Dim I, j As Integer
  
    I = 5:  temp = rec(I): temp = temp + rec(I + 1) * 256:    McuSys2Config.DesignVol = temp          '       //uint32_t ϵͳ�������ѹ(mV)
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.PackConfigMap = temp                 '       // uint32_tFullChargeCapacity ϵͳ��������(mAH)
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256
    McuSys2Config.FCC = temp      '       // uint32_tϵͳ����ѭ���ŵ�����(mAH)
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256
    McuSys2Config.CycleThreshold = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.CycleCount = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.NearFCC = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.DfilterCur = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.LearnLowTemp = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.SWVersion = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.HWVersion = temp
    I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.ShutDownDelay = temp
    I = I + 2: McuSys2Config.SelfDsgRate = rec(I)

 '   i = i + 1: McuSys2Config.ShutDownDelay = rec(i)
 '   i = i + 1: McuSys2Config.IdleDelay = rec(i)
    I = I + 1: McuSys2Config.CommOffDelay = rec(I)
    
         sstr = ""
       For j = 0 To 3
        I = I + 1
        sstr = sstr & byte_to_hex(rec(I))                '       // ManufactureName �������� ManufactureName[16]
       Next j
       McuSys2Config.MNFDate = sstr

        sstr = ""
       For j = 0 To 15
        I = I + 1
        sstr = sstr & Chr(rec(I))           '       // ManufactureName �������� ManufactureName[16]
       Next j
       McuSys2Config.MNFName = sstr
        sstr = ""
       For j = 0 To 15
        I = I + 1
        sstr = sstr & Chr(rec(I))               '       // ManufactureName �������� ManufactureName[16]
       Next j
       McuSys2Config.DeviceName = sstr
        sstr = ""
       For j = 0 To 15
        I = I + 1
        sstr = sstr & Chr(rec(I))                  '       // ManufactureName �������� ManufactureName[16]
       Next j
       McuSys2Config.SN = sstr
        I = I + 1: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.SOH = temp
        I = I + 2:
        sstr = ""
        For j = 0 To 7
            sstr = sstr & byte_to_hex(rec(I))               '       // ManufactureName �������� ManufactureName[16]
            I = I + 1
        Next j
        McuSys2Config.MCU_ID = sstr
 
 
        I = I: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.KEY_CODE = Int_to_Intel_hex(temp)
        I = I + 2: temp = rec(I): temp = temp + rec(I + 1) * 256: McuSys2Config.KEY_CODE = McuSys2Config.KEY_CODE & Int_to_Intel_hex(temp)
 Call frmMain.Printf_McuSys2Config
End Function
Private Function CRC_ADD(getStrall As String) As Byte
  Dim one As Long
  Dim OneH As Long
 
  Dim abc As String
  Dim temp As Long
  Dim CRCADD As Byte
  Dim crcmuch As Long
  Dim Cmdback As Byte
  Dim lenth As Long
  CRCADD = 0
  crcmuch = 0
  lenth = Len(getStrall)
  If lenth > 250 Then
     lenth = 250
  End If
  lenth = lenth - 1
  lenth = lenth / 2

    For I = 1 To lenth  ' info  ���� ��00AC  14+79*2  CONFIG ��  14+ 112*2      �о��þ÷��� ��������ǿո� ������Ҫת��һ��
        abc = Mid(getStrall, I * 2, 1)
        OneH = Val("&H" & abc) * 16
        crcmuch = crcmuch + Asc(abc)
        abc = Mid(getStrall, I * 2 + 1, 1)
        If abc = "" Then  ' ������Ϊ ʱ�� ���� �ս��ջ��ڷ��� �����ִ����·���Ȼ���м���ֶ����ǿ��ַ������ǿո�
        Else
            crcmuch = crcmuch + Asc(abc)
            one = Val("&H" & abc)
            rec(I - 1) = OneH + one
        End If
    Next I
    If crcmuch > 0 Then
      CRC_ADD = crcmuch Mod 256
        CRC_ADD = Not CRC_ADD
    End If

End Function

'  ���� �ַ���������� ��λУ����
Public Function ADDCRC(data As String) As Byte
 Dim lenss, I As Long
 Dim bbdata As Byte
 Dim sstrff As String
   Dim flage As Boolean
     sstrff = ""
     CRC16Lo = &HFF
     CRC16Hi = &HFF
    
     For I = 1 To 261
        sstrff = Mid(data, I * 2 - 1, 2)
        bbdata = Val("&H" & sstrff)
        CRC_oneByte (bbdata)
        sentda_iap(I - 1) = bbdata
     Next I
      I = CRC16Hi
      sentda_iap(261) = Val("&H" & byte_to_hex(I))
      I = CRC16Lo
      sentda_iap(262) = Val("&H" & byte_to_hex(I))
      
       
    flage = False


    If MSComm1.PortOpen = True Then
       If cmdsend.Enabled = True Then
        If MSComm1.CommEvent = 1010 Then
         Else
         flage = True
          MSComm1.Output = sentda_iap ' ���ﴫ����BYTE ����
        End If
       End If
    End If
    If flage = False Then 'ֻ�д��ڲ��� �ŷ�CAN
        If CAN_ONUSB_flag = True Then  ' û�д���ʱ ����CAN
          Form1_can.sentcan_bytedata
       End If
    End If
    
End Function
 

Private Function dealIAP_GETDATA()
   Onesecond = 55
    If IAP.ProgressBar1.Value < IAP.ProgressBar1.Max Then
   
        IAP.ProgressBar1.Value = IapCmd + 1
 
     IAP.Label_100.Caption = Format((IapCmd / IAP.ProgressBar1.Max) * 100 + 2, "#0.0")
    
      
    End If
    
    If Mid(rxStrall, 1, 2) = "6A" Then ' ��һ֡
        goto_reset_mcu_into = 1
        IapCmd = IapCmd + 1
        IAP_MCU_START_FLAG = 1 '  �����ٷ� 9528��
    End If
    If Mid(rxStrall, 1, 2) = "5B" Then '  �ط���֡
       If IapCmd > 4 Then
         IapCmd = IapCmd - 4
         Else
         IapCmd = 0
       End If
        IapCmd = IapCmd + 0
            goto_reset_mcu_into = 1
        IAP_CHONGSHI = IAP_CHONGSHI + 1
        If IAP_CHONGSHI > 20 Then
             IAP.Label1M.Caption = "����ʧ�ܣ������ϵ�����"
             IAP.Label1M.BackColor = &HFF&
             IAP.Label1M.Visible = True
        End If
    End If
    If Mid(rxStrall, 1, 2) = "6C" Then ' ��һ֡
        goto_reset_mcu_into = 1
       If IapCmd Then
          IapCmd = IapCmd - 1
       End If
    End If
    If Mid(rxStrall, 1, 2) = "5D" Then ' ���¿�ʼ
        goto_reset_mcu_into = 1
         IAP.Label_100.Caption = "100"
         IapCmd = 0
    End If
    If Mid(rxStrall, 1, 2) = "5E" Then
       
        sentIAPflag = 0 '  ���
        IapCmd = 0
        Getringht_sentF = 0
        IAP.Label1M.Caption = "�������"
        IAP.Label1M.BackColor = &HFF00&
        IAP.Label1M.Visible = True
    End If
    manual_time = 0
    Getringht_sentF = 1
 End Function

Private Function can_and_uart_sentiap()
Dim sstr As String
Dim lenbyte As Long
Dim oteher As Long
Dim I, lenn, lllbty As Long
    Getringht_sentF = 1 '  һֱ��
     If Getringht_sentF Then '4
     Getringht_sentF = 0
         If IapCmd = 0 Then '5
           jindu100 = 0
           jindu = 0   '  Flen
           If Flen < 256 Then '6
              sstr = "0000" & Int_to_hex(Flen)
           Else
               If Flen < 65536 Then '7
                   sstr = "00" & Int_to_hex(Flen)
               Else
                   If Flen < 16777210 Then '8
                        sstr = byte_to_hex(Fix(Flen / 256 / 256)) & Int_to_hex(Flen Mod 65536)
                   End If '8
               End If '7
           End If ''6
             '   �ļ�ͷ  CMD  �������ܳ�  ����  CRC
           sstr = "BE0000" & "0106" & sstr
           Call ADDCRC(sstr)
         Else '5
           sstr = ""
           lenbyte = IapCmd
           lenbyte = lenbyte * 256
           oteher = IapCmd - 1
           oteher = oteher * 256
           If lenbyte < Flen Then '6
                For I = 0 To 256 - 1
                   lllbty = I + oteher
                   sstr = sstr & byte_to_hex((FileBin(lllbty) Mod 256))
                Next I
                sstr = "BE" & Int_to_hex(IapCmd Mod 256) & "0106" & sstr
                Call ADDCRC(sstr)
            Else '6
                lenn = Flen + 256 - lenbyte
                For I = 0 To lenn - 1
                lllbty = I + oteher
                   sstr = sstr & byte_to_hex(FileBin(lllbty) Mod 256)
                Next I
                sstr = "BE" & Int_to_hex(IapCmd Mod 256) & "0106" & sstr
               Call ADDCRC(sstr)
           End If '6
         End If '5
    End If '4
    '   TextTx.Text = TextTx.Text & sent_out_UART_AND_CAN & vbCrLf
End Function
Private Function sent_IAP()

    If MSComm1.PortOpen = True Then '1
       If cmdsend.Enabled = True Then '2
           If MSComm1.CommEvent = 1010 Then '3
           Else
           
           End If '3
        End If '2
    Else '1
       If SentPussFlag = 0 Then '2
       End If '2
    End If '1
     
    Call can_and_uart_sentiap
    
End Function

 Private Function ReadBalckUp_485toRAM() As String   'PC  �·�   == 0x00 ) //��һ����¼������һ����   == 0x01 ) //��һ����¼    == 0x02 ) //�ط���ǰ��¼
 Dim eestr As String                                     'MCU �ظ�  01 ��һ��  02 �м��¼ 03 ���һ��
 Dim temp As Long
 Dim ress As Integer
 Dim aa, bb, cc As Long
 Dim ee, ff, gg As Long
 
 On Error GoTo Err1
 


  manual_time = 5 '��ȷ�յ� �������� ����ʱ ���ٷ���
 Delay_dis_ReadBalckUp = 6 ' ��ȷ�յ� ֱ����ʾ��ֱ������
 If rec(5) = 1 Then
   Record_Num = 0
 End If
 If rec(5) = 2 Then
   Record_Num = Record_Num + 1
 End If
If rec(5) = 3 Then ' ���� ��
     Record_Num = Record_Num + 1
     Delay_dis_ReadBalckUp = 100 '��ʾ �ɹ�
      manual_time = 0
      NextSentCmd = CMD_ReadInfo  '������ȷ �����������ݣ���Ȼ����� �ٴζ����һ����¼
 End If
 
 If rec(5) = 0 Then
     Record_Num = 0
     Delay_dis_ReadBalckUp = 200 '��ʾ�޼�¼
    manual_time = 0
    NextSentCmd = CMD_ReadInfo  '������ȷ �����������ݣ���Ȼ����� �ٴζ����һ����¼
       GoTo jumppp
 End If
  CMD_cmd_No = 1 '������ ֻҪ ������ȷ�ͷ�=1 ��һ�� ���ղ���ȷ �ڷ��ͺ���2 ����
  ress = 1
  
  If puse_blackup_button Then
   Delay_dis_ReadBalckUp = 0 ' ���ν��ղ���ʾ
   puse_blackup_button = 0
   Else
  End If
  ' Byte ': ʱ��,7bit �ֱ����ꡢ�¡��ա� �ܡ�ʱ���֡���'
  
  '  tdsRtcVal Rtc;              //uint8 * 7  ��¼ʱ��
  ress = 0
 
     
     aa = Hex(rec(24 + ress)) Mod 24
     bb = Hex(rec(23 + ress)) Mod 60
     cc = (rec(22 + ress)) Mod 60
     
     ee = Hex(rec((27 + ress)))
     ff = Hex(rec(26 + ress)) Mod 13
     gg = Hex(rec(25 + ress)) Mod 32
  eestr = DateSerial(ee, ff, gg) & " " & TimeSerial(aa, bb, cc)

  frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 1) = eestr
  
    Select Case rec(29 + ress)
        Case 1: eestr = "����"
        Case 2: eestr = "����"
        Case 3: eestr = "�ػ�"
        Case 4: eestr = "��ձ���"
        Case 5: eestr = "����"
        Case 6: eestr = "�յ�"
        Case 7: eestr = "��ѹ����"
        Case 8: eestr = "��ѹ�ָ�"
        Case 9: eestr = "Ƿѹ����"
        Case 10: eestr = "Ƿѹ�ָ�"
        Case 11: eestr = "��·����"
        Case 12: eestr = "��·�ָ�"
        Case 13: eestr = "������"
        Case 14: eestr = "�������ָ�"
        Case 15: eestr = "�ŵ����"
        Case 16: eestr = "�ŵ�����ָ�"
        Case 17: eestr = "�ŵ����2"
        Case 18: eestr = "�ŵ����2�ָ�"
        Case 19: eestr = "������"
        Case 20: eestr = "�����»ָ�"
        Case 21: eestr = "�ŵ����"
        Case 22: eestr = "�ŵ���»ָ�"
        Case 23: eestr = "������"
        Case 24: eestr = "�����»ָ�"
        Case 25: eestr = "�ŵ����"
        Case 26: eestr = "�ŵ���»ָ�"
        Case 27: eestr = "���θ�ѹ����"
        Case 28: eestr = "���ʸ��±���"
        Case 29: eestr = "���ʸ��»ָ�"
        Case 30: eestr = "�����λ"
        Case 31: eestr = "Ӳ����λ"
        Case 32: eestr = "��翪ʼ"
        Case 33: eestr = "������"
        Case 34: eestr = "�ŵ翪ʼ"
        Case 35: eestr = "�ŵ����"
        Case 36: eestr = "��ʱ��¼"
        Case 37: eestr = "������¼"
        Case 38: eestr = "SOC��̬����1"
        Case 39: eestr = "SOC��̬����1"
   End Select

  
  frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 2) = eestr                        '    uint8 * 1  ��¼����
 
    eestr = ""
    temp = rec(67 + ress)
    temp = temp * 256
    temp = temp + rec(66 + ress)
    If temp And 1 Then
        eestr = eestr & "�����ѹ "
    End If
    If temp And 2 Then
        eestr = eestr & "����Ƿѹ "
    End If
    If temp And 4 Then
        eestr = eestr & "������ѹ "
    End If
    If temp And 8 Then
        eestr = eestr & "�����Ƿѹ "
    End If
    If temp And 16 Then
        eestr = eestr & "����Ƿѹ "
    End If
    If temp And 32 Then
        eestr = eestr & "������ "
    End If
    If temp And 64 Then
        eestr = eestr & "������ "
    End If
    If temp And 128 Then
        eestr = eestr & "�ŵ���� "
    End If
    If temp And 256 Then
        eestr = eestr & "�ŵ���� "
    End If
    If temp And 1024 Then
        eestr = eestr & "���ʸ��� "
    End If
    If temp And 2048 Then
        eestr = eestr & "���ʵ��� "
    End If
    If temp And 4096 Then
        eestr = eestr & "������ "
    End If
    If temp And 8192 Then
        eestr = eestr & "�ŵ����1 "
    End If
    If temp And 16384 Then
        eestr = eestr & "�ŵ����2 "
    End If
    If temp = 0 Then
        eestr = eestr & "��������� "
    End If
 
   
  frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 3) = eestr
                                            temp = rec(31 + ress)
                                            temp = temp * 256
                                            temp = temp + rec(30 + ress)
  frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 4) = Int_to_hex(temp)                             '    uint16 ϵͳ״̬ "PackStatus"
                                            temp = rec(33 + ress)
                                            temp = temp * 256
                                            temp = temp + rec(32 + ress)
  frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 5) = Int_to_hex(temp)                           '    uint16 ���״̬
                                            temp = rec(9 + ress) And &H7F
                                            temp = temp * 256 + rec(8 + ress)
                                            temp = temp * 256 + rec(7 + ress)
                                            temp = temp * 256 + rec(6 + ress)
 frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 6) = temp                                '    uint32//ϵͳ��������(mAH)
                                            temp = rec(13 + ress) And &H7F
                                            temp = temp * 256 + rec(12 + ress)
                                            temp = temp * 256 + rec(11 + ress)
                                            temp = temp * 256 + rec(10 + ress)
  frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 7) = temp                                 '    uint32��ذ���ǰʣ�����(mAh)
  
                                       
  
                                        
                                                  temp = rec(79 + ress)
                                            temp = temp * 256 + rec(78 + ress)
                                        
 frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 8) = temp           '    uint16_t��ذ���ʣ������ٷֱ�(%)
  
  
                                            temp = rec(17 + ress) And &H7F
                                            temp = temp * 256 + rec(16 + ress)
                                            temp = temp * 256 + rec(15 + ress)
                                            temp = temp * 256 + rec(14 + ress)
    frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 9) = temp                          '    uint32_t��ذ��ܵ�ѹֵ(mV)
    
                                                temp = rec(21 + ress) And &H7F
                                                temp = temp * 256 + rec(20 + ress)
                                                temp = temp * 256 + rec(19 + ress)
                                                temp = temp * 256 + rec(18 + ress)
 frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 10) = (temp - 5000) / 10                   '    int32_tʵʱ����ֵ(mA)
 
    For I = 0 To 15
                                                  temp = rec(35 + I * 2 + ress)
                                                  temp = temp * 256 + rec(33 + I * 2 + 1 + ress)
    frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 11 + I) = temp                      '    uint16_t��о 1~16 �ĵ�ѹ(mV)
    Next I
    
  
    For I = 0 To 8
        frmMain.MSFlexGrid1.TextMatrix(Record_Num + 1, 27 + I) = rec(68 + I + ress) - 40
    Next I
    
    frmMain.MSFlexGrid1.ScrollTrack = True
    If Record_Num > 40 Then
        frmMain.MSFlexGrid1.TopRow = Record_Num - 40

   '  frmMain.MSFlexGrid1.TopRow = 0
    End If
     
    GoTo jumppp
Err1:
  ' My_msgbox "��⵽�������� ������ "
jumppp:
 End Function
 '     0,      5,  10,   15,   20, 25, 30,  35,    40, 45, 50,  55,   60,  65,  70,  75,  80,  85,  90,  95, 100, // �¶�
'    3433,  3461,3500,3544,3578,3594,3606,3617,3630,3645,3668,3716,3751,3789,3834,3882,3931,3981,4036,4091,4147, //-20c
'    3429,  3458,3498,3543,3580,3594,3606,3618,3631,3646,3669,3715,3750,3789,3834,3882,3931,3981,4036,4091,4147, //  0c
'    3404,  3452,3493,3538,3573,3593,3607,3620,3634,3651,3675,3720,3753,3791,3836,3884,3934,3985,4040,4097,4157, // 25c
'    3417,  3454,3493,3539,3574,3596,3611,3623,3636,3651,3672,3712,3749,3788,3832,3880,3928,3978,4032,4086,4140, // 45c
'    3416,  3454,3493,3538,3573,3596,3611,3623,3636,3651,3671,3710,3748,3787,3831,3879,3927,3976,4029,4083,4137, // 55c
'    -20,0,25,45,55,100
 Private Function ReadSOCOCV_485toRAM() As String  ''PC  �·�   == 0x00 ) //��һ����¼������һ����   == 0x01 ) //��һ����¼    == 0x02 ) //�ط���ǰ��¼
 Dim eestr As String                                     'MCU �ظ�  01 ��һ��  02 �м��¼ 03 ���һ��
 
 Dim temp As Long
 Dim I   As Integer
 Dim ss  As Single

 manual_time = 5 '��ȷ�յ� �������� ����ʱ ���ٷ���
 Delay_dis_ReadSOC_OCV = 6 ' ��ȷ�յ� ֱ����ʾ��ֱ������

    If rec(5) = 1 Then
      Record_Num = 1
    End If
    If rec(5) = 2 Then
      Record_Num = Record_Num + 1
     
    End If
     If rec(5) = 3 Then ' ���� ��
        Record_Num = Record_Num + 1
        Delay_dis_ReadSOC_OCV = 100 '��ʾ �ɹ�
        manual_time = 0
        NextSentCmd = CMD_ReadInfo  '������ȷ �����������ݣ���Ȼ����� �ٴζ����һ����¼
    End If
If Record_Num = 0 Then
 CMD_cmd_No = 0   'û���յ���һ�� ��һֱ��0
 Else
  CMD_cmd_No = 1
End If
     '������ ֻҪ ������ȷ�ͷ�=1 ��һ�� ���ղ���ȷ �ڷ��ͺ���2 ����
 
  
  If Record_Num = 1 Then '��һ���������в�ͬ
    SOC_OCVData(0, 0) = " "       'һ����ʾһ������     0-20��%
    For I = 0 To 20
       ss = rec(8 + I * 2) + rec(9 + I * 2) * 256 'һ����ʾһ������     0-20��%
       ss = ss
       SOC_OCVData(I + 1, 0) = str(ss) & "%SOC"
    Next I
  Else
    ss = rec(6) + rec(7) * 256
    ss = ss - 40
    SOC_OCVData(0, Record_Num - 1) = str(ss) & "��" 'һ����ʾһ������     0-20��%
    For I = 0 To 20
        ss = rec(8 + I * 2) + rec(9 + I * 2) * 256 'һ����ʾһ������     0-20��%
        ss = ss / 1000
        SOC_OCVData(I + 1, Record_Num - 1) = str(ss)
    Next I
  End If
 End Function
 '     0,      5,  10,   15,   20, 25, 30,  35,    40, 45, 50,  55,   60,  65,  70,  75,  80,  85,  90,  95, 100, // �¶�
'    3433,  3461,3500,3544,3578,3594,3606,3617,3630,3645,3668,3716,3751,3789,3834,3882,3931,3981,4036,4091,4147, //-20c
'    3429,  3458,3498,3543,3580,3594,3606,3618,3631,3646,3669,3715,3750,3789,3834,3882,3931,3981,4036,4091,4147, //  0c
'    3404,  3452,3493,3538,3573,3593,3607,3620,3634,3651,3675,3720,3753,3791,3836,3884,3934,3985,4040,4097,4157, // 25c
'    3417,  3454,3493,3539,3574,3596,3611,3623,3636,3651,3672,3712,3749,3788,3832,3880,3928,3978,4032,4086,4140, // 45c
'    3416,  3454,3493,3538,3573,3596,3611,3623,3636,3651,3671,3710,3748,3787,3831,3879,3927,3976,4029,4083,4137, // 55c
'    -20,0,25,45,55,100
 Private Function Readcap_485toRAM() As String  ''PC  �·�   == 0x00 ) //��һ����¼������һ����   == 0x01 ) //��һ����¼    == 0x02 ) //�ط���ǰ��¼
 Dim eestr As String                                     'MCU �ظ�  01 ��һ��  02 �м��¼ 03 ���һ��
 
 Dim temp As Long
 Dim I   As Integer
 Dim ss  As Single

 manual_time = 5 '��ȷ�յ� �������� ����ʱ ���ٷ���
 Delay_dis_Readcap = 100 ' ��ȷ�յ� ֱ����ʾ��ֱ������
     '������ ֻҪ ������ȷ�ͷ�=1 ��һ�� ���ղ���ȷ �ڷ��ͺ���2 ����
        For I = 0 To 107
          'һ����ʾһ������     0-20��%
        capData(I) = rec(5 + I)
        Next I
   capReal_inMAXV = rec(5 + 107 + 1) * 256 + rec(5 + 107 + 2)
   capReal_inMAXV = capReal_inMAXV / 100
   capReal_inMAXA = rec(5 + 107 + 3) * 256 + rec(5 + 107 + 4)
   capReal_inMAXA = capReal_inMAXA / 100
   capReal_OutMAXA = rec(5 + 107 + 5) * 256 + rec(5 + 107 + 6)
   capReal_OutMAXA = capReal_OutMAXA / 100
   capReal_charge_onoff = rec(5 + 107 + 7)
   capReal_discharge_onoff = rec(5 + 107 + 8)

  frmMain.get_real_cap.Caption = "InMAXV" & "=" & capReal_inMAXV & "V" & "//" & "InMAXA" & "=" & capReal_inMAXA & "A" & "//" & "OutMAXA" & "=" & capReal_OutMAXA & "A" & "//" & "Csg" & "=" & capReal_charge_onoff & "//""Dsg" & "=" & capReal_discharge_onoff & "//"

        
 End Function
 
 Private Function deal_WriteSOC_ocv_backtoRAM() As String  ''PC  �·�  1  ��һ��  2 �м���  3 ������
 Dim eestr As String                                        'MCU �ظ�  1��һ��      2 �� û�лظ�  �ظ���ǰ��      4 ���´�������  0 �ɹ�
 
 Dim temp As Long
 Dim I   As Integer
 Dim ss  As Single
 
 manual_time = 5 '��ȷ�յ� �������� ����ʱ ���ٷ���
 Delay_dis_WriteSOC_OCV = 6 ' ��ȷ�յ� ֱ����ʾ��ֱ������
    
    If rec(5) = 4 Then
            manual_time = 0
    End If
    
    If rec(5) = 2 Then
        If Record_Num = 0 Then
          CMD_cmd_No = 1
        End If
        If Record_Num < 5 Then
          CMD_cmd_No = 2
         Else
          CMD_cmd_No = 3
        End If
    End If
    
    If rec(5) = 1 Then ' ���� ��
       Record_Num = Record_Num + 1
       
        If Record_Num < 5 Then
          CMD_cmd_No = 2
        Else
            CMD_cmd_No = 3
        End If
        If Record_Num > 5 Then
            NextSentCmd = CMD_ReadInfo
            manual_time = 0
        End If

    End If
    
    If rec(5) = 0 Then
    Flag_sys2ok = True
         NextSentCmd = CMD_ReadInfo
        manual_time = 0
         Delay_dis_WriteSOC_OCV = 100  'û���յ���һ�� ��һֱ��0
    End If
 
 
 End Function
 
  Private Function deal_Writecap_backtoRAM() As String  ''PC  �·�  1  ��һ��  2 �м���  3 ������
 Dim eestr As String                                        'MCU �ظ�  1��һ��      2 �� û�лظ�  �ظ���ǰ��      4 ���´�������  0 �ɹ�
 
 Dim temp As Long
 Dim I   As Integer
 Dim ss  As Single
 
 manual_time = 5 '��ȷ�յ� �������� ����ʱ ���ٷ���
 Delay_dis_Writecap = 6 ' ��ȷ�յ� ֱ����ʾ��ֱ������
    
    If rec(5) = 4 Then
            manual_time = 0
    End If
    
    If rec(5) = 2 Then
        If Record_Num = 0 Then
          CMD_cmd_No = 1
        End If
        If Record_Num < 5 Then
          CMD_cmd_No = 2
         Else
          CMD_cmd_No = 3
        End If
    End If
    
    If rec(5) = 1 Then ' ���� ��
       Record_Num = Record_Num + 1
       
        If Record_Num < 5 Then
          CMD_cmd_No = 2
        Else
            CMD_cmd_No = 3
        End If
        If Record_Num > 5 Then
            NextSentCmd = CMD_ReadInfo
            manual_time = 0
        End If

    End If
    
    If rec(5) = 0 Then
        Flag_sys2ok = True
         NextSentCmd = CMD_ReadInfo
        manual_time = 0
         Delay_dis_Writecap = 100  'û���յ���һ�� ��һֱ��0
    End If
 
 
 End Function
 
'     ���� SOC-OCV    PC  �����·�          1  ��һ��  2 �м���  3 ������
 '               MCY  �ظ�                  1  ��һ��            2 �� û�лظ�  �ظ���ǰ��      4 ���´�������  0 �ɹ�
Private Function deal_WriteSOC_OCV_RamtoStr() As String  ' �� SOC ���� ת�� �ַ���
Dim eestr As String
Dim I, j, k As Integer
 

  If Record_Num = 0 Then '��һ���������в�ͬ
    eestr = byte_to_hex(CMD_cmd_No) & "0000"       'һ����ʾһ������     0-20��%
    For I = 0 To 20
       ss = SOC_OCVData(I + 1, 0) 'һ����ʾһ������     0-20��%
       j = Len(ss)
       If (j > 4) Then
           eestr = eestr & Int_to_Intel_hex(Val(Mid(ss, 1, j - 4)))
       Else
       
       End If
    Next I
  Else
  If Record_Num > 5 Then
    Record_Num = 5 '̫������ ���ܵ�6 ����
  End If
  
       ss = SOC_OCVData(0, Record_Num) 'һ����ʾһ������     0-20��%
       j = Len(ss)
       ss = Mid(ss, 1, j - 1)
       eestr = byte_to_hex(CMD_cmd_No) & Int_to_Intel_hex(Val(ss) + 40)
 
    For I = 0 To 20
       ss = SOC_OCVData(I + 1, Record_Num) 'һ����ʾһ������     0-20��%
       eestr = eestr & Int_to_Intel_hex(Val(ss) * 1000)
    Next I
  End If
  deal_WriteSOC_OCV_RamtoStr = eestr
End Function

'     ���� SOC-OCV    PC  �����·�          1  ��һ��  2 �м���  3 ������
 '               MCY  �ظ�                  1  ��һ��            2 �� û�лظ�  �ظ���ǰ��      4 ���´�������  0 �ɹ�
Private Function deal_Writecap_RamtoStr() As String  ' �� SOC ���� ת�� �ַ���
Dim eestr As String
Dim I, j, k As Integer
    eestr = ""       'һ����ʾһ������     0-20��%
    For I = 0 To 107
           eestr = eestr & byte_to_hex(capData(I))
    Next I
  deal_Writecap_RamtoStr = eestr
End Function

Private Function deal_SysConfig_RamtoStr() As String   ' ��SYSCONFIG ����ת�� �ַ���
Dim eestr As String

Call frmMain.CmdSYSWrite_Click  '//�������� �·�
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.EngDesign)                 '   //�������,�� McuV82SysConfig.0_1AH)
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BalanceCur)                 '  //"����������С�������� McuV82SysConfig.mA)"  ԭ���������    ���������С    0_01mR��
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BalanceDelay)                 '    //  ���������ʱ�� McuV82SysConfig.S��  ԭ���������    �ο���ѹ    mv  10
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.B_VStart)                 '    //����������ѹ�� McuV82SysConfig.mV��
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.B_Vdiff)                 ' //��������ѹ� McuV82SysConfig.mV��10
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_Vcell_H)                 '   //���ڸ�ѹ����ֵmv
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_VCell_L)                    '    //���ڵ�ѹ����ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_VBAT_H)                 '    //��ظ�ѹ����ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_VBAT_L)                 '    //��ص�ѹ����ֵ    26
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_CURR_C)                 '    //����������ֵ0_01A   20
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_CURR_D)                 '    //�ŵ��������ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_VDIFF_H)                 '   //ѹ���ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.W_VDIFF_L)                 '   //ѹ������ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.OVPVal)                 '  //��������ѹ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.OVPDly)                 '  //������䱣����ʱ  30
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.OVPRel)                 '  //�������ָ���ѹ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.UVPVal)                 '  //������ŵ�ѹ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.UVPDly)                 '  //������ű�����ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.UVPRel)                 '  //������Żָ���ѹ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BOVPVal)                 ' //�����������ѹ  40
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BOVPDly)                 ' //���������䱣����ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BOVPRel)                 ' //����������ָ���ѹ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BUVPVal)                 ' //��ع��ŵ�ѹ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BUVPDly)                 ' //��ع��ű�����ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.BUVPRel)                 ' //��ع��Żָ���ѹ  50
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CC_PRO_VAL)                 '  //����������ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CC_PRO_PDLY)                 ' //������������ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CC_PRO_RDLY)                 ' //�������ָ���ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CC_PRO_LOCK)                 ' //��������������
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD1_PRO_VAL)                 ' //  һ���ŵ籣��ֵ  60
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD1_PRO_PDLY)                 '    //һ���ŵ����������ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD1_PRO_RDLY)                 '    //һ���ŵ�����ָ���ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD1_PRO_LOCK)                 '    //һ���ŵ������������
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD2_PRO_VAL)                 ' //  �����ŵ籣��ֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD2_PRO_PDLY)                 '    //�����ŵ����������ʱ  70
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD2_PRO_RDLY)                 '    //�����ŵ�����ָ���ʱ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.CD2_PRO_LOCK)                 '    //�����ŵ������������
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.SHORT_RDLY)                 '  //��·��ʱֵ
eestr = eestr & Int_to_Intel_hex(McuV82SysConfig.SHORT_LOCK)                 '  //��·����ֵ
eestr = eestr & byte_to_hex(McuV82SysConfig.CTcellHPro)                  '  //��о�����±���
eestr = eestr & byte_to_hex(McuV82SysConfig.CTcellHRel)                  '  //��о�����±����ָ�80
eestr = eestr & byte_to_hex(McuV82SysConfig.CTcellLPro)                  '  //��о�����±���
eestr = eestr & byte_to_hex(McuV82SysConfig.CTcellLRel)                  '  //��о�����±����ָ�
eestr = eestr & byte_to_hex(McuV82SysConfig.DTcellHPro)                  '  //��о�ŵ���±���
eestr = eestr & byte_to_hex(McuV82SysConfig.DTcellHRel)                  '  //��о�ŵ���±����ָ�
eestr = eestr & byte_to_hex(McuV82SysConfig.DTcellLPro)                  '  //��о�ŵ���±���85
eestr = eestr & byte_to_hex(McuV82SysConfig.DTcellLRel)                  '  //��о�ŵ���±����ָ�
eestr = eestr & byte_to_hex(McuV82SysConfig.TenvHPro)                  '    //��о�������±���
eestr = eestr & byte_to_hex(McuV82SysConfig.TenvHRel)                  '    //��о�������±����ָ�
eestr = eestr & byte_to_hex(McuV82SysConfig.TenvLPro)                  '    //��о�������±���
eestr = eestr & byte_to_hex(McuV82SysConfig.TenvLRel)                  '    //��о�������±����ָ�90
eestr = eestr & byte_to_hex(McuV82SysConfig.TfetHPro)                  '    //��о���ʸ��±���
eestr = eestr & byte_to_hex(McuV82SysConfig.TfetHRel)                  '    //��о���ʸ��±����ָ�
eestr = eestr & byte_to_hex(McuV82SysConfig.TfetLPro)                  '    //��о���ʵ��±���
eestr = eestr & byte_to_hex(McuV82SysConfig.TfetLRel)                  '    //��о���ʵ��±����ָ�
eestr = eestr & byte_to_hex(McuV82SysConfig.W_Tcell_H)                  '   //��о���¾���ֵ95
eestr = eestr & byte_to_hex(McuV82SysConfig.W_Tcell_L)                     '    //��о���¾���ֵ
eestr = eestr & byte_to_hex(McuV82SysConfig.W_Tenv_H)                  '    //�������¾���ֵ
eestr = eestr & byte_to_hex(McuV82SysConfig.W_Tenv_L)                  '    //�������¾���ֵ
eestr = eestr & byte_to_hex(McuV82SysConfig.W_Tfet_H)                  '    //���ʸ��¾���ֵ
eestr = eestr & byte_to_hex(McuV82SysConfig.W_Tfet_L)                  '    //���ʵ��¾���ֵ    100
eestr = eestr & byte_to_hex(McuV82SysConfig.B_Mode)                  '  //����ģʽ  0~2��0)              '  //������    1)              '   //������  2   ���+��̬����
eestr = eestr & byte_to_hex(McuV82SysConfig.B_THDIS)                  ' //������½�ֵֹ    40  ��ʾ0�� 65  ��ʾ25��
eestr = eestr & byte_to_hex(McuV82SysConfig.B_TLDIS)                  ' //������½�ֵֹ
eestr = eestr & byte_to_hex(McuV82SysConfig.Addr)                  '    //������    RS485   ��ַ    1~255
eestr = eestr & byte_to_hex(McuV82SysConfig.CellNum)                  ' //��ؽ���  5~16    105
eestr = eestr & byte_to_hex(McuV82SysConfig.TempsetNum)                  '  //����ʹ��  4���¶� ���ò���
eestr = eestr & byte_to_hex(McuV82SysConfig.SHORT_VAL)                  '   //  ��·��ѹ����ֵ
eestr = eestr & byte_to_hex(McuV82SysConfig.HEAT_EN)                  ' //���ȹ���ʹ��
eestr = eestr & byte_to_hex(McuV82SysConfig.HEAT_TSTART)                  ' //  ���ȿ����¶�
eestr = eestr & byte_to_hex(McuV82SysConfig.HEAT_TEND)                  '   //  ���ȹر��¶�    110


deal_SysConfig_RamtoStr = eestr
End Function

Private Function deal_str_to_asciiStr(strrrr As String) As String    ' ��SYSCONFIG ����ת�� �ַ���
Dim I As Integer
Dim eestr As String

For I = 1 To Len(strrrr)
    ccchar = Mid(strrrr, I, 1)
    If ccchar = "" Then
           eestr = eestr & byte_to_hex(20)
    Else:
      eestr = eestr & byte_to_hex(Val(Asc(ccchar)))
    End If
Next I
deal_str_to_asciiStr = eestr
End Function
Private Function deal_WriteSys2Config_RamtoStr() As String   ' ��SYSCONFIG ����ת�� �ַ���
Dim eestr, strrrr, ccchar, ssrrtt As String
Dim I As Integer

Call frmMain.Command_Sys2Write_Click '//�������� �·�
eestr = ""
eestr = eestr & Int_to_Intel_hex(McuSys2Config.DesignVol)
eestr = eestr & Int_to_Intel_hex(Val("&H" & McuSys2Config.PackConfigMap))  '       // MCU ϵͳ���ò���
eestr = eestr & Int_to_Intel_hex(McuSys2Config.FCC)
eestr = eestr & Int_to_Intel_hex(McuSys2Config.CycleThreshold)
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.CycleCount)) '       // MCU ϵͳ���ò���
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.NearFCC)) '       // MCU ϵͳ���ò���
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.DfilterCur)) '       // MCU ϵͳ���ò���
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.LearnLowTemp)) '       // MCU ϵͳ���ò���
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.SWVersion)) '       // MCU ϵͳ���ò���
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.HWVersion)) '       // MCU ϵͳ���ò���

eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.ShutDownDelay)) 'eestr = eestr & byte_to_hex(Int(McuSys2Config.IdleDelay)) '
eestr = eestr & byte_to_hex(Int(McuSys2Config.SelfDsgRate)) '// �Էŵ���(0.01%)
eestr = eestr & byte_to_hex(Int(McuSys2Config.CommOffDelay)) '
eestr = eestr & Mid(McuSys2Config.MNFDate, 1, 8) '
strrrr = Mid(McuSys2Config.MNFName, 1, 16) '

For I = 1 To 16
    ccchar = Mid(strrrr, I, 1)
    If ccchar = "" Then
           eestr = eestr & byte_to_hex(20)
    Else:
      eestr = eestr & byte_to_hex(Val(Asc(ccchar)))
  
    End If
Next I
strrrr = Mid(McuSys2Config.DeviceName, 1, 16) '
For I = 1 To 16
    ccchar = Mid(strrrr, I, 1)
    If ccchar = "" Then
           eestr = eestr & byte_to_hex(20)
    Else:
      eestr = eestr & byte_to_hex(Val(Asc(ccchar)))
  
    End If
Next I
ssrrtt = str(Val(Mid(McuSys2Config.SN, 4, 8)) + AUTO_SNUM)
 ssrrtt = Replace(ssrrtt, " ", "")
McuSys2Config.SN = Mid(McuSys2Config.SN, 1, 4)
McuSys2Config.SN = McuSys2Config.SN + ssrrtt '�Զ���д �Զ���
 strrrr = Mid(McuSys2Config.SN, 1, 16)
For I = 1 To 16
    ccchar = Mid(strrrr, I, 1)
    If ccchar = "" Then
           eestr = eestr & byte_to_hex(20)
    Else:
      eestr = eestr & byte_to_hex(Val(Asc(ccchar)))
  
    End If
Next I
eestr = eestr & Int_to_Intel_hex(Int(McuSys2Config.SOH)) '       // MCU SOH
strrrr = Mid(McuSys2Config.MCU_ID, 1, 16) '
For I = 1 To 8
    ccchar = Mid(strrrr, I, 1)
    If ccchar = "" Then
           eestr = eestr & byte_to_hex(20)
    Else:
            eestr = eestr & byte_to_hex(Val(Asc(ccchar)))
    End If
Next I


If BMS_active_mode = 22 Or jingdu1 = 4 Then
    eestr = McuSys2Config.KEY_CODE      '       // MCU KEY_CODE
Else
    eestr = eestr & McuSys2Config.KEY_CODE    '       // MCU KEY_CODE
End If
deal_WriteSys2Config_RamtoStr = eestr

End Function
Private Function deal_RegEERPOM_RamtoStr() As String
Dim eestr As String
Dim I As Integer
Call frmMain.ReadTheRegchang
RegEERPOM(25) = "00"
 For I = 0 To 25
   eestr = eestr & RegEERPOM(I)
 Next I
deal_RegEERPOM_RamtoStr = eestr
End Function
 
'  ����һ֡ ����  ����CRC �� ����β��
Private Function deal_setTheCMD(ss As Long) As String
Dim eestr As String
Dim I As Integer
    eestr = ":" & byte_to_hex(PC_ADDR) & byte_to_hex(ss) & byte_to_hex(PC_VER) & "000E" ' SOI  Addr  Cmd  Ver  Len  Info  CRC  EOI
    eestr = eestr & CRC16_keycodedata(eestr) & "~"
deal_setTheCMD = eestr
End Function
Private Function deal_sent_more_TheCMD(cmd As Long, sstr As String) As String
Dim eestr As String
Dim I As Long
    I = Len(sstr)
    eestr = ":" & byte_to_hex(PC_ADDR) & byte_to_hex(cmd) & byte_to_hex(PC_VER) & "00" & byte_to_hex(14 + I)    ' SOI  Addr  Cmd  Ver  Len  Info  CRC  EOI
    eestr = eestr & sstr
    eestr = eestr & CRC16_keycodedata(eestr) & "~"
deal_sent_more_TheCMD = eestr
End Function
Public Sub sent_CMD_withcan(cmd As Byte)
             Select Case cmd
                Case CMD_ReadSN  '0x01 ��ȡSN��
                           sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadSN))
                Case CMD_ReadSOCSOP   '0x02 ��ȡSOC,SOP SOP����%�ȣ������ѹ����ʱ��˻�
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadSOCSOP))
                Case CMD_ReadVOLTAGE_CURREN   '0x03 ��ȡ��ѹ������
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadVOLTAGE_CURREN))
                Case CMD_ReadInfo   '0x04 ��ȡʵʱ����
                If jingdu1 = 0 Then
                  sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadInfo))
                End If
                              
                Case CMD_ReadSysConfig  '0x05 ��ȡ��������
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadSysConfig))
                Case CMD_ReadBalckUp  '0x06 ��ȡ��������    RD_EEPROM      CMD_ReadBalckUp �� CMD_cmd_No �ɽ��մ����� ����
                 '�·�   == 0x00 ) //��һ����¼������һ����              == 0x01 ) //��һ����¼                       == 0x02 ) //�ط���ǰ��¼
                                 sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_ReadBalckUp, byte_to_hex(CMD_cmd_No)))
                                 NextSentCmd = CMD_ReadBalckUp ' û�лظ�һֱ��
                                 CMD_cmd_No = 2 ' ֻҪ���ղ���ȷ ���ط�
                                 manual_time = 40
                Case CMD_ReadSys2Config  '0x07 ��ȡ��������    RD_MCUSYSTEM
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadSys2Config))
                Case CMD_ReadAFEseg  '0x08 ��ȡ�Ĵ�������  RD_MTP
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadAFEseg))
                Case CMD_ReadRTC   '0x09 ��RTC
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReadRTC))
             '   Case CMD_ReadMcuRAM   '0x0A ��ȡ�ڲ�״̬
                              '  sent_out_UART_AND_CAN( deal_setTheCMD(CMD_ReadMcuRAM)
                Case CMD_ReadSOC_OCV   '0x0B ��SOC���ò���    CMD_ReadSOC_OCV �� CMD_cmd_No �Ͷ�ȡ�������� ����һ�����ɽ��պ�������
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_ReadSOC_OCV, byte_to_hex(CMD_cmd_No)))
                               manual_time = 5
                               NextSentCmd = CMD_ReadSOC_OCV ' û�лظ�һֱ��
                               CMD_cmd_No = 2 ' ֻҪ���ղ���ȷ ���ط�
                 Case CMD_Readcap   '0x0B ��SOC���ò���    CMD_ReadSOC_OCV �� CMD_cmd_No �Ͷ�ȡ�������� ����һ�����ɽ��պ�������
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_Readcap, byte_to_hex(CMD_cmd_No)))
                               manual_time = 10
                               
                               CMD_cmd_No = 2 ' ֻҪ���ղ���ȷ ���ط�
                Case CMD_WriteAFEseg  '0x20 ���üĴ�������  WR_MTP
                                sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_WriteAFEseg, deal_RegEERPOM_RamtoStr))   ' �����ñ���ֵ ���� �ٷ���
                                manual_time = 40 ' 200ms  ���һʵʱ״̬
                Case CMD_SetFET   '0x21 �·�FET����
                                sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_SetFET, byte_to_hex(sent_result)))
                Case CMD_WriteSysConfig  '0x22 �·����ò���
                                sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_WriteSysConfig, deal_SysConfig_RamtoStr))  ' �����ñ���ֵ ���� �ٷ���
                Case CMD_EraseBalckUp  '0x23 �·�������������
                                sent_out_UART_AND_CAN (deal_setTheCMD(CMD_EraseBalckUp))
                Case CMD_CALIB_VOLTAGE  '0x24 �·�У���ܵ�ѹ  CALIB_VOLTAGE
                                manual_time = 10 ' 200ms  ���һʵʱ״̬
                                sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_CALIB_VOLTAGE, Int_to_Intel_hex(Fix(sent_result Mod 65536)) & Int_to_Intel_hex(Fix(sent_result / 65536))))  '�·��ܵ�ѹʱ��Mv����
                Case CMD_CALIB_CURRENT  '0x25 �·�У������    CALIB_CURRENT    00  xH xL
                                manual_time = 10 ' 200ms  ���һʵʱ״̬  ���� ԭ����2 ����ʧ�����ڸĳ�5�ɹ�
                                If CMD_cmd_No = 1 Then
                                    strm = frmMain.RealCurrentkText.Text
                                    If (InStr(1, strm, "-", 1)) > 0 Then
                                         strm = Replace(strm, "-", "")
                                         sent_result = strm * 1000
                                    ' RealCurrentkText.Text = -RealCurrentkText.Text
                                        strm = Int_to_Intel_hex(65536 - sent_result)    ' ���ָ���
                                     Else
                                         sent_result = strm * 1000
                                         strm = Int_to_Intel_hex(sent_result)
                                     End If
                                End If
                                sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_CALIB_CURRENT, byte_to_hex(CMD_cmd_No) & strm)) '�·��ܵ���ʱ��ma����
                Case CMD_CALIB_TEMPE  '0x26 �·�У���¶�    CALIB_TEMPE
                                 manual_time = 10 ' 200ms ����һ�� ���8��
                                 For I = CMD_cmd_No To 8
                                    If Claib_temp(CMD_cmd_No) Then
                                          sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_CALIB_TEMPE, byte_to_hex(CMD_cmd_No) & Int_to_Intel_hex(Fix(frmMain.RealTemptext(CMD_cmd_No + 1).Text + 40))))  '�·��ܵ���ʱ��ma����
                                          I = 8
                                    End If
                                          CMD_cmd_No = CMD_cmd_No + 1
                                    If CMD_cmd_No < 9 Then
                                           Delay_dis_CALIB_Temp = 4
                                          NextSentCmd = CMD_CALIB_TEMPE
                                    Else
                                           Delay_dis_CALIB_Temp = 100
                                    End If
                                 Next I
                                 
                Case CMD_CALIB_RTC   '0x27 �·�����RTC CALIB_RTC  2020/5/6 15:32:01 Time��Date��Now��
                               manual_time = 10 ' 200ms  ���һʵʱ״̬
                   texxx = Format(Now, "yyyy-mm-dd hh:mm:ss") ' ͳһ ʱ���ʽ
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_CALIB_RTC, Mid(texxx, 3, 2) & Mid(texxx, 6, 2) & Mid(texxx, 9, 2) & Mid(texxx, 12, 2) & Mid(texxx, 15, 2) & Mid(texxx, 18, 2)))       '
                Case CMD_Enter_Sleep_Mode  '0x28 �·�BMS����ػ� Enter_Sleep_Mode
                               sent_out_UART_AND_CAN (deal_setTheCMD(CMD_Enter_Sleep_Mode))
                Case V82_SET_POWERON  '0x28 �·�BMS����ػ� Enter_Sleep_Mode
                               sent_out_UART_AND_CAN (deal_setTheCMD(V82_SET_POWERON))
                Case CMD_ReSet_OFFSET  ' 0x35
                               sent_out_UART_AND_CAN (deal_setTheCMD(CMD_ReSet_OFFSET))
                Case CMD_ISP_HANDSHAKE  '0x29 �·�����IAP_����    ISP_HANDSHAKE
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_ISP_HANDSHAKE, "09050208"))
                Case CMD_WriteSOC_OCV   '0x30 �·�����SOC���ò���
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_WriteSOC_OCV, deal_WriteSOC_OCV_RamtoStr))  ' �����ñ���ֵ ���� �ٷ���
                               manual_time = 5
                               NextSentCmd = CMD_WriteSOC_OCV
                Case CMD_Writecap   '0x30 �·�����SOC���ò���
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_Writecap, deal_Writecap_RamtoStr))  ' �����ñ���ֵ ���� �ٷ���
                               manual_time = 5
                            
                Case CMD_WriteSys2Config   '0x31 �·���������
                    If BMS_active_mode = 22 Or jingdu1 = 4 Then
                        sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_ActiveBms, deal_WriteSys2Config_RamtoStr))  ' �����ñ���ֵ ���� �ٷ���
                    Else
                        sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_WriteSys2Config, deal_WriteSys2Config_RamtoStr))  ' �����ñ���ֵ ���� �ٷ���
                    End If
                    manual_time = 20
                     NextSentCmd = CMD_ReadInfo '2023.2.4
                Case CMD_ReSet_MCU
                               sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_ReSet_MCU, "09050202"))
                Case CMD_Blue_name
                               ' If GetheairF = True Then
                                    sent_out_UART_AND_CAN (deal_sent_more_TheCMD(CMD_Blue_name, deal_str_to_asciiStr(bluetooth_name)))
                              '  Else
                                '    sent_out_UART_AND_CAN( "AT+NAME" & bluetooth_name & vbCr & vbLf
                             '   End If
 
                                
             End Select
End Sub

Public Sub sent_CMD(cmd As Byte)
  Dim texxx As String
    Dim strm As String
  PC_ADDR = 0 ' �趨PC��ַΪ0
    If MSComm1.PortOpen = True Then
       If cmdsend.Enabled = True Then
       If MSComm1.CommEvent = 1010 Then
           frmMain.Label_dis1.Caption = "���ڴ���"
           frmMain.Label_dis1.ForeColor = &HFF&
        Else
              frmMain.Label_dis1.Caption = "����������"
              frmMain.Label_dis1.ForeColor = &HFFF580
        Call sent_CMD_withcan(cmd)
        frmMain.Label_strdis.ForeColor = &H800000
       End If
       End If
       
       
       '  TextTx.Text = TextTx.Text & texxx & vbCrLf
  
    Else
       If CAN_ONUSB_flag = True Then   ' û�д���ʱ ����CAN
         Call sent_CMD_withcan(cmd)
       End If
       If SentPussFlag = 0 Then
        frmMain.Label_dis1.Caption = "����δ��"
        frmMain.Label_dis1.ForeColor = &HFF&
       End If
    End If
End Sub

Public Sub sent_out_UART_AND_CAN(outstr As String)

  Dim flage As Boolean
       
    flage = False

    If MSComm1.PortOpen = True Then
       If cmdsend.Enabled = True Then
        If MSComm1.CommEvent = 1010 Then
         Else
         flage = True
          MSComm1.Output = outstr
        End If
       End If
    End If
    If flage = False Then 'ֻ�д��ڲ��� �ŷ�CAN
        If CAN_ONUSB_flag = True Then  ' û�д���ʱ ����CAN
         Form1_can.sentcandata (outstr)
       End If
    End If

End Sub
' �����·� ���� ��˼   �ɼ�PACK ʱ ���趨ֵ
' У׼ʱ    �ɼ�PACK  ��ʱѭ������
' ���� AFE  ϵͳ����1 ϵͳ����2 ��ʱ����
' ��ȡ ��������ʱ SOC����   ��ʱ����ѭ������
' IAP����ʱ  �����ö�ʱѭ������
' ��ʱ����   ÿ100ms ����һ�� ������MCU ��������
Private Sub Timer2_Timer()
 Dim the_cytime As Long
 
 
    Call frmMain.deal_auto   ' �Զ���д�ж�
    
 Onesecond = Onesecond + 1
 '                                    �ܷ��� -������ȷ                         ������ȷ                      У���ʶ
 frmMain.tongxunerror.Caption = CStr(Sent_data_lj - rightback_lj) + "/" + CStr(rightback_lj) + "/" + CStr(backcrc_error_lj)
 
 If GET_DATA Then  ' ��������ʾ ����������ȷ ��ǰ��һ��
    jingdutiao = jingdutiao + 1 '
    If jingdutiao > 32760 Then
        jingdutiao = 0
    End If
    
    If jingdutiao > 5 Then
      jingdutiao = 0
    End If
    frmMain.ProgressBar1.Value = jingdutiao
 End If

If manual_time Then
 the_cytime = manual_time
Else
 the_cytime = cyInfoTime
End If

   GET_DATA = 0 '������ȷ ��� Ŀǰֻ���� ��������ʾ��
 If Delay_waite_muc_back_cmd Then
  Delay_waite_muc_back_cmd = Delay_waite_muc_back_cmd - 1
  Onesecond = Onesecond - 1
 End If
If Onesecond >= (the_cytime) Then    ' ÿ manual_time*100ms ִ��һ��
   manual_time = 0 ' ���ΰ���  ���͵���ʱ��ֻ��һ��
   Onesecond = 0

 
    If sentIAPflag Then ' =1 Ϊ����ģʽ
         manual_time = 20
        If goto_reset_mcu Then
           goto_reset_mcu = 0
           If IAP_MCU_START_FLAG = 0 Then   '  �����ٷ� 9528��
            Call sent_CMD(CMD_ISP_HANDSHAKE)  ' ���� V82��λ����
           End If
           NextSentCmd = CMD_ReadInfo
        Else
           If goto_reset_mcu_into = 0 Then
                 goto_reset_mcu = 1          ' ��δ ����IAP ʱ��=1����һ�� V82��λ����
           End If
           Call sent_IAP  '  ���� IAP ͷ֡���ݼ�����
        End If
    Else
        SentCmd = NextSentCmd
        NextSentCmd = CMD_ReadInfo ' ֻ��һ�� �������� ����
        Call sent_CMD(SentCmd)
        Sent_data_lj = Sent_data_lj + 1
    End If
End If


If mscomm_delay > 0 Then  ' �մ򿪴���ʱ ��ʼ��ʱʹ�ã����ܻ���ֲ��ܽ��յ�����
    mscomm_delay = mscomm_delay - 1
End If
End Sub





