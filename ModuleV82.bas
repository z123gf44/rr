Attribute VB_Name = "Module2"
Option Explicit
Public auto_500ms  As Long
Public m_devtype As Long  'CAN ͨ�����ַ��� �ر�
Public caniap_completeflag As Boolean  '  =1  ������
 Public sentda_iap(262) As Byte
Public CAN_ONUSB_flag   As Boolean  '  =1  ������
Public LAST_MCU_ID   As String  '  ����ͣ=1
Public BMS_admin_mode  As Long
'=0 �����û� ֻ�鿴 һ������
'=1 �м��û�Ӳֻ�鿴 Ӳ������   16874162
'=2�м��û����ò���1            25649813
'=3�м��û����ò���2            36546123
'=4�м�ȫ��鿴���û�+�鿴���޸ļ�¼ + ��3��������       44654658
'=5�м�ȫ��鿴���û�+��¼ + 3��������    52342342
'=6�����������úͲ鿴           66546546
'=7 �����û� ȫ������ +��¼ + 3��������+У��        75212354
Public Flag_sys2ok  As Boolean  '�Զ����ͽ���
Public Flag_onlysys2ok  As Boolean  '�Զ����ͽ���
Public Flag_readmcusys2ok  As Boolean  '�Զ����ͽ���
Public Flag_readckeckjiemasys2ok  As Boolean  '�Զ����ͽ���
Public jingdu1, AUTO_SNUM As Long '�Զ����ͽ���
Public mode_bit1, mode_bit2, mode_bit3, mode_bit4, mode_bit5, mode_bit6, mode_bit7, mode_bit8, mode_bit9, mode_bit10, mode_bit11, mode_bit12, mode_bit13    As Long '
Public BMS_active_mode  As Long  ' =1 ��ʾ��ɫ  =0 ����  ��������Ǻ�ɫ
Public havegetTRightData As Byte  '������ȷ���ݱ�־
Public GET_DATA         As Byte           '���յ�У����ȷ������ �·�����ʱ=0 ��������ȷʱ=1 ��=1ʱ������++
Public Sent_data_lj            As Long
Public rightback_lj         As Long
Public backcrc_error_lj         As Long

Public myRealV82Info    As MyV82Type
Public McuV82SysConfig  As MyV82SysConfig
Public McuSys2Config    As MySys2Config
Public Record_Num       As Long
Public jilu_path        As String    '��¼�ļ� ����
Public jiema_jilu_path        As String    '��¼�ļ� ����
Public LOAD_CELLmun     As Long    ' ��ؽ����仯 ���¼���һ�� ����λ��
Public LOAD_Tempmun     As Long    ' ��ؽ����仯 ���¼���һ�� ����λ��
Public Const Von = 11  ' ��ѹ�������ϱ�״̬
Public Const Con = 8
Public Const Ton = 14
Public Const Aon = 8
Public Const Fon = 9
Public Const Gon = 12
Public bluetooth_name As String  ' 1S ʱ�� �ⷢ��ʱ ��ѯ
Public OnecyTimes As Integer  ' 1S ʱ�� �ⷢ��ʱ ��ѯ
Public Onesecond As Integer  ' 1S ʱ�� �ⷢ��ʱ ��ѯ
Public SentCmd   As Byte ' ͨѶ ���� 01����������  02 ��ʵʱ����  05 ���ñ������� 06 FET ���� 09 ��ȡ�汾
                            ' 51 ��ȡ�ڲ�״̬  52 ��
                            ' ˼· ������ ÿ��һ��ȡ���� ������ ��ʱ�ⷢ ������ �Ӻ�
                            ' �ⷢ���� SentCmd   �������� ���� �ⷢ���� �жϣ�
Public rec(300)                     As Long
Public Backmessage                  As Byte  ' �������� =0 δ�յ�   =8a ���ճɹ�    =8bʧ��
Public sysCaption(75)               As String
Public sys2Caption(75)              As String
Public NextSentCmd                  As Byte
Public SOC_OCVData(56, 26)          As String
Public capData(300)               As Long
Public capReal_inMAXV               As Single
Public capReal_inMAXA              As Single
Public capReal_OutMAXA              As Single
Public capReal_charge_onoff              As Long
Public capReal_discharge_onoff              As Long
Public RegEERPOM(26)                As String
Public time_regScan                 As Integer
Public jingdutiao                   As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'��ʱ���� ��ȡ�Ĵ����ظ����� =100 ��ʾ�ɹ����յ� ������Ϣ��
' ÿ������ ���ֶ��·�ʱ ��ʼ���� 4�κ�û�н��գ����� ʧ����Ϣ���յ������ɹ�����
Public Delay_dis_Readsysconfig      As Long
Public Delay_dis_Writesysconfig     As Long
Public Delay_dis_ReadRegAfe         As Long
Public Delay_dis_WriteRegAfe        As Long
Public Delay_dis_Readsys2config     As Long
Public Delay_dis_Writesys2config    As Long
Public Delay_dis_EraseBalckUp       As Long
Public Delay_dis_Enter_Sleep_Mode   As Long
Public Delay_dis_Enter_WORK_Mode   As Long
Public Delay_dis_SetFET             As Long
Public Delay_dis_ReadSOC_OCV        As Long
Public Delay_dis_WriteSOC_OCV       As Long  ' �е���
Public Delay_dis_Readcap       As Long
Public Delay_dis_Writecap       As Long  ' �е���
Public Delay_dis_CALIB_RTC          As Long
Public Delay_dis_CALIB_VOLTAGE      As Long
Public Delay_dis_CALIB_CURRENT      As Long
Public Delay_dis_CALIB_Temp         As Long
Public Delay_dis_ReadBalckUp        As Long
Public Delay_dis_ReadMcuRAM         As Long
Public Delay_waite_muc_back_cmd     As Long ' �ȴ� n*100ms ʱ��
Public manual_time                  As Long
Public RecordTime_add               As Long         '��¼ʱ�� �ж��ж���CYINFOTIME
Public RecordTime                   As Long         '��¼ʱ�� ��� 500 ms
Public cyInfoTime                   As Long         '��¼ʱ�� ��� 500 ms
Public FileBin(131072)              As Byte         '�������ļ�
Public Flen                         As Long         '�ļ�����
Public IapCmd                       As Byte          '�·� ���� 00 �ļ�ͷ 01-FF �ļ� ���һ֡ ����У��� У��
Public jindu100                     As Integer           ' �·�����%
Public jindu                        As Long          ' �·�����
Public IAP_MCU_START_FLAG           As Long         ' ���ֹ� ����IAP��ɺ��ַ���һ�� 9528ȥ�ˣ�Ȼ����Ҫ����IAP��
Public sentIAPflag                  As Byte          ' �·���־ =1 �ڷ��� =0 û��
Public Getringht_sentF              As Byte          '  =1 �յ��� �ٷ� =0�ȴ�
Public IAP_CHONGSHI                 As Long          ' 100������ʧ��
Public Claib_temp(10)               As Byte          ' �·���MCU ���õ�ֵ
Public sent_result                  As Long          ' �·���MCU ���õ�ֵ
Public mscomm_delay                 As Long          ' ������ʱ2s �ٿ�����
Public goto_reset_mcu               As Long               ' =1 ���� ��9527 =0 ���� IAPͷ�ļ�
Public goto_reset_mcu_into          As Long          ' =1 9527 �Ѿ�=0
Public PC_ADDR                      As Long          ' PC ��ַ
Public PC_VER                       As Long          ' ����汾
Public CMD_cmd_No                   As Long          ' ͬһ���������˳�� ��У����������� �� ����
Public puse_blackup_button          As Integer           '  ����ͣ=1
Public Const CMD_ReadSN = &H1               '0x01 ��ȡSN��
Public Const CMD_ReadSOCSOP = &H2           '0x02 ��ȡSOC,SOP SOP����%�ȣ������ѹ����ʱ��˻�
Public Const CMD_ReadVOLTAGE_CURREN = &H3   '0x03 ��ȡ��ѹ������
Public Const CMD_ReadInfo = &H4             '0x04 ��ȡʵʱ����
Public Const CMD_ReadSysConfig = &H5        '0x05 ��ȡ��������
Public Const CMD_ReadBalckUp = &H6          '0x06 ��ȡ��������    RD_EEPROM
Public Const CMD_ReadSys2Config = &H7       '0x07 ��ȡ��������    RD_MCUSYSTEM
Public Const CMD_ReadAFEseg = &H8           '0x08 ��ȡ�Ĵ�������  RD_MTP
Public Const CMD_ReadRTC = &H9              '0x09 ��RTC
'//Public Const CMD_ReadMcuRAM = &HB           '0x0A ��ȡ�ڲ�״̬
Public Const CMD_ReadSOC_OCV = &HA          '0x0B ��SOC���ò���
Public Const CMD_Readcap = &HB        '0x0B ��SOC���ò���
Public Const CMD_WriteAFEseg = &H20         '0x20 ���üĴ�������  WR_MTP
Public Const CMD_SetFET = &H21              '0x21 �·�FET����
Public Const CMD_WriteSysConfig = &H22      '0x22 �·����ò���
Public Const CMD_EraseBalckUp = &H23        '0x23 �·�������������
Public Const CMD_CALIB_VOLTAGE = &H24       '0x24 �·�У���ܵ�ѹ  CALIB_VOLTAGE
Public Const CMD_CALIB_CURRENT = &H25       '0x25 �·�У������    CALIB_CURRENT
Public Const CMD_CALIB_TEMPE = &H26         '0x26 �·�У���¶�    CALIB_TEMPE
Public Const CMD_CALIB_RTC = &H27           '0x27 �·�����RTC CALIB_RTC
Public Const CMD_Enter_Sleep_Mode = &H28    '0x28 �·�BMS����ػ� Enter_Sleep_Mode
Public Const CMD_ISP_HANDSHAKE = &H29       '0x29 �·�����IAP_����    ISP_HANDSHAKE
Public Const CMD_WriteSOC_OCV = &H30        '0x30 �·�����SOC���ò���
Public Const CMD_Writecap = &H33        '0x30 �·�����SOC���ò���
Public Const CMD_WriteSys2Config = &H31     '0x31 �·�ϵͳ����2����
Public Const CMD_ActiveBms = &H91     '0x31 �·�ϵͳ����2����
Public Const CMD_ReSet_MCU = &H32           '0x31 �·�BMS��λ
Public Const CMD_Blue_name = &H41           '0x41 �·�BMS�޸�����BMS ����
Public Const V82_SET_POWERON = &H34           '0x34 �·�BMS��������
Public Const CMD_ReSet_OFFSET = &H35           '0x35 �·���λ�¶� ���� ��У��ֵ
 'ͨѶ���
 ' ����������ʱ  �� У������ʱ ÿ0.5S���� ��ȡʵʱ����
 ' ���������� �����  ���ͺ� ÿ0.5S ����
 ' ��������  ������ ������
 ' ���з��� ��20ms ��ʱ�� ��������
