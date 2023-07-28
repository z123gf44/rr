Attribute VB_Name = "Module2"
Option Explicit
Public auto_500ms  As Long
Public m_devtype As Long  'CAN 通道名字方便 关闭
Public caniap_completeflag As Boolean  '  =1  发送中
 Public sentda_iap(262) As Byte
Public CAN_ONUSB_flag   As Boolean  '  =1  可以用
Public LAST_MCU_ID   As String  '  按暂停=1
Public BMS_admin_mode  As Long
'=0 初级用户 只查看 一个界面
'=1 中级用户硬只查看 硬件配置   16874162
'=2中级用户设置参数1            25649813
'=3中级用户设置参数2            36546123
'=4中级全面查看用用户+查看不修改记录 + 无3参数配置       44654658
'=5中级全面查看用用户+记录 + 3参数配置    52342342
'=6三个参数设置和查看           66546546
'=7 超级用户 全部功能 +记录 + 3参数配置+校正        75212354
Public Flag_sys2ok  As Boolean  '自动发送进度
Public Flag_onlysys2ok  As Boolean  '自动发送进度
Public Flag_readmcusys2ok  As Boolean  '自动发送进度
Public Flag_readckeckjiemasys2ok  As Boolean  '自动发送进度
Public jingdu1, AUTO_SNUM As Long '自动发送进度
Public mode_bit1, mode_bit2, mode_bit3, mode_bit4, mode_bit5, mode_bit6, mode_bit7, mode_bit8, mode_bit9, mode_bit10, mode_bit11, mode_bit12, mode_bit13    As Long '
Public BMS_active_mode  As Long  ' =1 显示黄色  =0 正常  激活错误是红色
Public havegetTRightData As Byte  '接收正确数据标志
Public GET_DATA         As Byte           '接收到校验正确的数据 下发命令时=0 ，接收正确时=1 ，=1时进度条++
Public Sent_data_lj            As Long
Public rightback_lj         As Long
Public backcrc_error_lj         As Long

Public myRealV82Info    As MyV82Type
Public McuV82SysConfig  As MyV82SysConfig
Public McuSys2Config    As MySys2Config
Public Record_Num       As Long
Public jilu_path        As String    '记录文件 名字
Public jiema_jilu_path        As String    '记录文件 名字
Public LOAD_CELLmun     As Long    ' 电池节数变化 重新加载一个 窗口位置
Public LOAD_Tempmun     As Long    ' 电池节数变化 重新加载一个 窗口位置
Public Const Von = 11  ' 电压方面总上报状态
Public Const Con = 8
Public Const Ton = 14
Public Const Aon = 8
Public Const Fon = 9
Public Const Gon = 12
Public bluetooth_name As String  ' 1S 时间 外发定时 查询
Public OnecyTimes As Integer  ' 1S 时间 外发定时 查询
Public Onesecond As Integer  ' 1S 时间 外发定时 查询
Public SentCmd   As Byte ' 通讯 命令 01读保护参数  02 读实时参数  05 设置保护参数 06 FET 操作 09 读取版本
                            ' 51 读取内部状态  52 读
                            ' 思路 无命令 每秒一次取数， 有命令 及时外发 ，读数 延后
                            ' 外发命令 SentCmd   接收命令 根据 外发命令 判断，
Public rec(300)                     As Long
Public Backmessage                  As Byte  ' 发下命令 =0 未收到   =8a 接收成功    =8b失败
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
'延时处理 读取寄存器回复数据 =100 表示成功接收到 弹出消息框
' 每个变量 在手动下发时 开始计数 4次后没有接收，弹出 失败消息，收到弹出成功内容
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
Public Delay_dis_WriteSOC_OCV       As Long  ' 有弹窗
Public Delay_dis_Readcap       As Long
Public Delay_dis_Writecap       As Long  ' 有弹窗
Public Delay_dis_CALIB_RTC          As Long
Public Delay_dis_CALIB_VOLTAGE      As Long
Public Delay_dis_CALIB_CURRENT      As Long
Public Delay_dis_CALIB_Temp         As Long
Public Delay_dis_ReadBalckUp        As Long
Public Delay_dis_ReadMcuRAM         As Long
Public Delay_waite_muc_back_cmd     As Long ' 等待 n*100ms 时间
Public manual_time                  As Long
Public RecordTime_add               As Long         '记录时间 判断有多少CYINFOTIME
Public RecordTime                   As Long         '记录时间 间隔 500 ms
Public cyInfoTime                   As Long         '记录时间 间隔 500 ms
Public FileBin(131072)              As Byte         '二进制文件
Public Flen                         As Long         '文件长度
Public IapCmd                       As Byte          '下发 命令 00 文件头 01-FF 文件 最后一帧 加总校验的 校验
Public jindu100                     As Integer           ' 下发进度%
Public jindu                        As Long          ' 下发进度
Public IAP_MCU_START_FLAG           As Long         ' 出现过 整个IAP完成后，又发了一组 9528去了，然后又要进入IAP了
Public sentIAPflag                  As Byte          ' 下发标志 =1 在发中 =0 没有
Public Getringht_sentF              As Byte          '  =1 收到后 再发 =0等待
Public IAP_CHONGSHI                 As Long          ' 100次升级失败
Public Claib_temp(10)               As Byte          ' 下发给MCU 设置的值
Public sent_result                  As Long          ' 下发给MCU 设置的值
Public mscomm_delay                 As Long          ' 串口延时2s 再开接收
Public goto_reset_mcu               As Long               ' =1 发送 ：9527 =0 发送 IAP头文件
Public goto_reset_mcu_into          As Long          ' =1 9527 已经=0
Public PC_ADDR                      As Long          ' PC 地址
Public PC_VER                       As Long          ' 程序版本
Public CMD_cmd_No                   As Long          ' 同一个命令的子顺序 如校正电流的零点 和 线性
Public puse_blackup_button          As Integer           '  按暂停=1
Public Const CMD_ReadSN = &H1               '0x01 读取SN码
Public Const CMD_ReadSOCSOP = &H2           '0x02 读取SOC,SOP SOP能量%比，计算电压电流时间乘积
Public Const CMD_ReadVOLTAGE_CURREN = &H3   '0x03 读取总压，电流
Public Const CMD_ReadInfo = &H4             '0x04 读取实时参数
Public Const CMD_ReadSysConfig = &H5        '0x05 读取保护参数
Public Const CMD_ReadBalckUp = &H6          '0x06 读取备份数据    RD_EEPROM
Public Const CMD_ReadSys2Config = &H7       '0x07 读取出厂参数    RD_MCUSYSTEM
Public Const CMD_ReadAFEseg = &H8           '0x08 读取寄存器数据  RD_MTP
Public Const CMD_ReadRTC = &H9              '0x09 读RTC
'//Public Const CMD_ReadMcuRAM = &HB           '0x0A 读取内部状态
Public Const CMD_ReadSOC_OCV = &HA          '0x0B 读SOC配置参数
Public Const CMD_Readcap = &HB        '0x0B 读SOC配置参数
Public Const CMD_WriteAFEseg = &H20         '0x20 设置寄存器数据  WR_MTP
Public Const CMD_SetFET = &H21              '0x21 下发FET操作
Public Const CMD_WriteSysConfig = &H22      '0x22 下发设置参数
Public Const CMD_EraseBalckUp = &H23        '0x23 下发擦除备份数据
Public Const CMD_CALIB_VOLTAGE = &H24       '0x24 下发校正总电压  CALIB_VOLTAGE
Public Const CMD_CALIB_CURRENT = &H25       '0x25 下发校正电流    CALIB_CURRENT
Public Const CMD_CALIB_TEMPE = &H26         '0x26 下发校正温度    CALIB_TEMPE
Public Const CMD_CALIB_RTC = &H27           '0x27 下发更新RTC CALIB_RTC
Public Const CMD_Enter_Sleep_Mode = &H28    '0x28 下发BMS进入关机 Enter_Sleep_Mode
Public Const CMD_ISP_HANDSHAKE = &H29       '0x29 下发进入IAP_升级    ISP_HANDSHAKE
Public Const CMD_WriteSOC_OCV = &H30        '0x30 下发设置SOC配置参数
Public Const CMD_Writecap = &H33        '0x30 下发设置SOC配置参数
Public Const CMD_WriteSys2Config = &H31     '0x31 下发系统参数2设置
Public Const CMD_ActiveBms = &H91     '0x31 下发系统参数2设置
Public Const CMD_ReSet_MCU = &H32           '0x31 下发BMS复位
Public Const CMD_Blue_name = &H41           '0x41 下发BMS修改蓝牙BMS 名称
Public Const V82_SET_POWERON = &H34           '0x34 下发BMS开机命令
Public Const CMD_ReSet_OFFSET = &H35           '0x35 下发复位温度 电流 等校正值
 '通讯设计
 ' 当在主界面时  和 校正界面时 每0.5S发送 读取实时数据
 ' 进入升级后 界面后  发送后 每0.5S 发送
 ' 其它界面  不按下 不发送
 ' 所有发送 由20ms 定时器 函数调用
