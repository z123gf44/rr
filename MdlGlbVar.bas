Attribute VB_Name = "MdlGlbVar"


Public Type MyV82SysConfig
EngDesign            As Long         '   //设计容量,(0_1AH)
BalanceCur           As Long         '  //"均衡启动最小充电电流(mA)"    原来这个不用    采样电阻大小    0_01mR）
BalanceDelay         As Long         '    //  软件均衡延时(S）    原来这个不用    参考电压    mv  10
B_VStart             As Long         '    //均衡启动电压（mV）
B_Vdiff              As Long         ' //均衡启动压差（mV）10
W_Vcell_H            As Long         '   //单节高压警告值mv
W_VCell_L            As Long         '    //单节低压警告值
W_VBAT_H             As Long         '    //电池高压警告值
W_VBAT_L             As Long         '    //电池低压警告值    26
W_CURR_C             As Long         '    //充电电流警告值0_01A   20
W_CURR_D             As Long         '    //放电电流警告值
W_VDIFF_H           As Long         '   //压差报警值
W_VDIFF_L           As Long         '   //压差报警解除值
OVPVal              As Long         '  //单体过充电压
OVPDly              As Long         '  //单体过充保护延时  30
OVPRel              As Long         '  //单体过充恢复电压
UVPVal           As Long         '  //单体过放电压
UVPDly           As Long         '  //单体过放保护延时
UVPRel           As Long         '  //单体过放恢复电压
BOVPVal           As Long         ' //电池总体过充电压  40
BOVPDly           As Long         ' //电池总体过充保护延时
BOVPRel           As Long         ' //电池总体过充恢复电压
BUVPVal           As Long         ' //电池过放电压
BUVPDly           As Long         ' //电池过放保护延时
BUVPRel           As Long         ' //电池过放恢复电压  50
CC_PRO_VAL           As Long         '  //充电电流保护值
CC_PRO_PDLY           As Long         ' //充电电流保护延时
CC_PRO_RDLY           As Long         ' //充电电流恢复延时
CC_PRO_LOCK           As Long         ' //充电电流保护锁定
CD1_PRO_VAL           As Long         ' //  一级放电保护值  60
CD1_PRO_PDLY           As Long         '    //一级放电电流保护延时
CD1_PRO_RDLY           As Long         '    //一级放电电流恢复延时
CD1_PRO_LOCK           As Long         '    //一级放电电流保护锁定
CD2_PRO_VAL           As Long         ' //  二级放电保护值
CD2_PRO_PDLY           As Long         '    //二级放电电流保护延时  70
CD2_PRO_RDLY           As Long         '    //二级放电电流恢复延时
CD2_PRO_LOCK           As Long         '    //二级放电电流保护锁定
SHORT_RDLY           As Long         '  //短路延时值
SHORT_LOCK           As Long         '  //短路锁定值
CTcellHPro           As Long         '  //电芯充电高温保护
CTcellHRel           As Long         '  //电芯充电高温保护恢复80
CTcellLPro           As Long         '  //电芯充电低温保护
CTcellLRel           As Long         '  //电芯充电低温保护恢复
DTcellHPro           As Long         '  //电芯放电高温保护
DTcellHRel           As Long         '  //电芯放电高温保护恢复
DTcellLPro           As Long         '  //电芯放电低温保护85
DTcellLRel           As Long         '  //电芯放电低温保护恢复
TenvHPro           As Long         '    //电芯环境高温保护
TenvHRel           As Long         '    //电芯环境高温保护恢复
TenvLPro           As Long         '    //电芯环境低温保护
TenvLRel           As Long         '    //电芯环境低温保护恢复90
TfetHPro           As Long         '    //电芯功率高温保护
TfetHRel           As Long         '    //电芯功率高温保护恢复
TfetLPro           As Long         '    //电芯功率低温保护
TfetLRel           As Long         '    //电芯功率低温保护恢复
W_Tcell_H           As Long         '   //电芯高温警告值95
W_Tcell_L              As Long         '    //电芯低温警告值
W_Tenv_H           As Long         '    //环境高温警告值
W_Tenv_L           As Long         '    //环境低温警告值
W_Tfet_H           As Long         '    //功率高温警告值
W_Tfet_L           As Long         '    //功率低温警告值    100
B_Mode           As Long         '  //均衡模式  0~2，0           as     long     '  //不均衡    1           as     long     '   //充电均衡  2   充电+静态均衡
B_THDIS           As Long         ' //均衡高温禁止值    40  表示0℃ 65  表示25℃
B_TLDIS           As Long         ' //均衡低温禁止值
Addr           As Long         '    //保护板    RS485   地址    1~255
CellNum           As Long         ' //电池节数  5~16    105
SHORT_VAL   As Long     ';    // 短路电压保护值
TempsetNum           As Long         '  //温度个娄
HEAT_EN           As Long         ' //加热功能使能
HEAT_TSTART           As Long         ' //  加热开启温度
HEAT_TEND           As Long         '   //  加热关闭温度    110


End Type


Public Type MyV82Type
        Time_t              As String ': 时间,7bit 分别是年、月、日、 周、时、分、秒'
        mcu_powerStatu          As Long            '    MCU 工作状态
        Vbat                As String                 ' 电池电压，输出为总电压的0.5倍'
        Vcell_num           As Long    ': 电池串数，1-16'
        RealTempNum             As Long    ': 温度采样个数'
        Vcell(50)           As String   '：每一节电压 mV'
        Curr                As String   ': Curr[0]充电电流，Curr[1]放电电流'
        temp(32)            As String  ': 每个温度的数据，65 表示 25℃，正偏量 40' '
        vstate              As Long ' '
        Cstate              As Long  ' '
        Tstate              As Long    '  '数据结构如下
        Alarm               As Long '     '数据结构如下
        Fetstate            As Long  '' 数据结构如下
        NUM_VOV             As Long  '单体高压对应的电池的序号，例如 5 表示第 5 节高压
        NUM_VUV             As Long  ' ：单体欠压对应的电池的序号
        NUM_WARN_VHIGH      As Long ' ：单体高压警告对应的电池的序号
        NUM_WARN_VLOW       As Long ' ：单体低压警告对应的电池的序号
        BlanceState         As Long ' ： 均衡状态，表示那一节电压开启均衡
        DchgNum             As Long ' ：放电次数'
        BatStatus           As Long ' ：充电次数'
        SOC                 As String '  : 电池 soc ，百分比 0-100'
        CapNow              As String ' : 当前容量 (0.1AH)
        CapFull             As String ' : 满充容量(0.1AH)
        FET_code               As Long   '       // 产品序列号[10]
        afe_Temp(4)             As String
End Type

Public Type MySys2Config
       DesignVol            As Long   '       // 系统建议充电电压(mV)
       PackConfigMap        As Long   '       // MCU 系统配置参数
       FCC                  As Long   '       // FullChargeCapacity 系统满充容量(mAH)
       CycleThreshold       As Long   '       // 系统单次循环放电总量(mAH)
       CycleCount           As Long   '       // 循环放电次数
       NearFCC              As Long   '       // 有效放电开始时剩余容量与满充容量的最大差值(mAH)
       LearnLowTemp         As Long   '       // 满充容量更新允许的最低温度
       DfilterCur           As Long   '       // 系统能检测的最小电流，小于有效放电窗口的电流为 0
       SWVersion            As Long   '       // 软件版本：V1.00
       HWVersion            As Long   '       // 硬件版本：V1.00

       ShutDownDelay        As Long   '       // 进入低功耗模式等待时间(S)
       SelfDsgRate          As Long   '       // 自放电率(0.01%)
  '     IdleDelay            As Long   '       // 进入静置功耗模式等待时间(S)
       CommOffDelay         As Long   '       // 系统无上位机允许进入低功耗模式的延时(S)
       MNFDate          As String   '       // 生产日期：前 2 字节存放“年”，第 3 字节存放“月”，第 4 字节存放“日”
       MNFName          As String   '       // ManufactureName 生产厂商 ManufactureName[16]
       DeviceName       As String   '       // 产品名称：DeviceName [16]
       SN               As String   '       // 产品序列号[10]
       SOH               As String   '       // 产品序列号[10]
        MCU_ID               As String   '       // 产品序列号[10]
        KEY_CODE               As String   '       // 产品序列号[10]
End Type



Public Type MybackupRecord
 
    Time_t(7)               As Long            '    uint8 * 7   时间,7bit 分别是年、月、日、 周、时、分、秒'
    RecordType              As Long            '    uint8 * 1  记录类型
    PackStatus              As Long            '    uint16 系统状态
    BatStatus               As Long            '    uint16 电池状态
    FCC                     As Long            '    uint32//系统满充容量(mAH)
    rc                      As Long            '    uint32电池包当前剩余电量(mAh)
    PackVol                 As Long            '    uint32_t电池包总电压值(mV)
    Current                 As Long            '    int32_t实时电流值(mA)
    RSOC                    As Long            '    uint16_t电池包的剩余电量百分比(%)
    CellVol(16)             As Long            '    uint16_t电芯 1~16 的电压(mV)
    AmbientTemp             As Long            '    uint8
    PowerTemp               As Long            '    uint8
    CellTemp(7)             As Long            '    uint8;
    Protect_state           As Long            '    uint16_t电芯 软件保护状态
End Type

