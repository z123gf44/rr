Attribute VB_Name = "MdlGlbVar"


Public Type MyV82SysConfig
EngDesign            As Long         '   //�������,(0_1AH)
BalanceCur           As Long         '  //"����������С������(mA)"    ԭ���������    ���������С    0_01mR��
BalanceDelay         As Long         '    //  ���������ʱ(S��    ԭ���������    �ο���ѹ    mv  10
B_VStart             As Long         '    //����������ѹ��mV��
B_Vdiff              As Long         ' //��������ѹ�mV��10
W_Vcell_H            As Long         '   //���ڸ�ѹ����ֵmv
W_VCell_L            As Long         '    //���ڵ�ѹ����ֵ
W_VBAT_H             As Long         '    //��ظ�ѹ����ֵ
W_VBAT_L             As Long         '    //��ص�ѹ����ֵ    26
W_CURR_C             As Long         '    //����������ֵ0_01A   20
W_CURR_D             As Long         '    //�ŵ��������ֵ
W_VDIFF_H           As Long         '   //ѹ���ֵ
W_VDIFF_L           As Long         '   //ѹ������ֵ
OVPVal              As Long         '  //��������ѹ
OVPDly              As Long         '  //������䱣����ʱ  30
OVPRel              As Long         '  //�������ָ���ѹ
UVPVal           As Long         '  //������ŵ�ѹ
UVPDly           As Long         '  //������ű�����ʱ
UVPRel           As Long         '  //������Żָ���ѹ
BOVPVal           As Long         ' //�����������ѹ  40
BOVPDly           As Long         ' //���������䱣����ʱ
BOVPRel           As Long         ' //����������ָ���ѹ
BUVPVal           As Long         ' //��ع��ŵ�ѹ
BUVPDly           As Long         ' //��ع��ű�����ʱ
BUVPRel           As Long         ' //��ع��Żָ���ѹ  50
CC_PRO_VAL           As Long         '  //����������ֵ
CC_PRO_PDLY           As Long         ' //������������ʱ
CC_PRO_RDLY           As Long         ' //�������ָ���ʱ
CC_PRO_LOCK           As Long         ' //��������������
CD1_PRO_VAL           As Long         ' //  һ���ŵ籣��ֵ  60
CD1_PRO_PDLY           As Long         '    //һ���ŵ����������ʱ
CD1_PRO_RDLY           As Long         '    //һ���ŵ�����ָ���ʱ
CD1_PRO_LOCK           As Long         '    //һ���ŵ������������
CD2_PRO_VAL           As Long         ' //  �����ŵ籣��ֵ
CD2_PRO_PDLY           As Long         '    //�����ŵ����������ʱ  70
CD2_PRO_RDLY           As Long         '    //�����ŵ�����ָ���ʱ
CD2_PRO_LOCK           As Long         '    //�����ŵ������������
SHORT_RDLY           As Long         '  //��·��ʱֵ
SHORT_LOCK           As Long         '  //��·����ֵ
CTcellHPro           As Long         '  //��о�����±���
CTcellHRel           As Long         '  //��о�����±����ָ�80
CTcellLPro           As Long         '  //��о�����±���
CTcellLRel           As Long         '  //��о�����±����ָ�
DTcellHPro           As Long         '  //��о�ŵ���±���
DTcellHRel           As Long         '  //��о�ŵ���±����ָ�
DTcellLPro           As Long         '  //��о�ŵ���±���85
DTcellLRel           As Long         '  //��о�ŵ���±����ָ�
TenvHPro           As Long         '    //��о�������±���
TenvHRel           As Long         '    //��о�������±����ָ�
TenvLPro           As Long         '    //��о�������±���
TenvLRel           As Long         '    //��о�������±����ָ�90
TfetHPro           As Long         '    //��о���ʸ��±���
TfetHRel           As Long         '    //��о���ʸ��±����ָ�
TfetLPro           As Long         '    //��о���ʵ��±���
TfetLRel           As Long         '    //��о���ʵ��±����ָ�
W_Tcell_H           As Long         '   //��о���¾���ֵ95
W_Tcell_L              As Long         '    //��о���¾���ֵ
W_Tenv_H           As Long         '    //�������¾���ֵ
W_Tenv_L           As Long         '    //�������¾���ֵ
W_Tfet_H           As Long         '    //���ʸ��¾���ֵ
W_Tfet_L           As Long         '    //���ʵ��¾���ֵ    100
B_Mode           As Long         '  //����ģʽ  0~2��0           as     long     '  //������    1           as     long     '   //������  2   ���+��̬����
B_THDIS           As Long         ' //������½�ֵֹ    40  ��ʾ0�� 65  ��ʾ25��
B_TLDIS           As Long         ' //������½�ֵֹ
Addr           As Long         '    //������    RS485   ��ַ    1~255
CellNum           As Long         ' //��ؽ���  5~16    105
SHORT_VAL   As Long     ';    // ��·��ѹ����ֵ
TempsetNum           As Long         '  //�¶ȸ�¦
HEAT_EN           As Long         ' //���ȹ���ʹ��
HEAT_TSTART           As Long         ' //  ���ȿ����¶�
HEAT_TEND           As Long         '   //  ���ȹر��¶�    110


End Type


Public Type MyV82Type
        Time_t              As String ': ʱ��,7bit �ֱ����ꡢ�¡��ա� �ܡ�ʱ���֡���'
        mcu_powerStatu          As Long            '    MCU ����״̬
        Vbat                As String                 ' ��ص�ѹ�����Ϊ�ܵ�ѹ��0.5��'
        Vcell_num           As Long    ': ��ش�����1-16'
        RealTempNum             As Long    ': �¶Ȳ�������'
        Vcell(50)           As String   '��ÿһ�ڵ�ѹ mV'
        Curr                As String   ': Curr[0]��������Curr[1]�ŵ����'
        temp(32)            As String  ': ÿ���¶ȵ����ݣ�65 ��ʾ 25�棬��ƫ�� 40' '
        vstate              As Long ' '
        Cstate              As Long  ' '
        Tstate              As Long    '  '���ݽṹ����
        Alarm               As Long '     '���ݽṹ����
        Fetstate            As Long  '' ���ݽṹ����
        NUM_VOV             As Long  '�����ѹ��Ӧ�ĵ�ص���ţ����� 5 ��ʾ�� 5 �ڸ�ѹ
        NUM_VUV             As Long  ' ������Ƿѹ��Ӧ�ĵ�ص����
        NUM_WARN_VHIGH      As Long ' �������ѹ�����Ӧ�ĵ�ص����
        NUM_WARN_VLOW       As Long ' �������ѹ�����Ӧ�ĵ�ص����
        BlanceState         As Long ' �� ����״̬����ʾ��һ�ڵ�ѹ��������
        DchgNum             As Long ' ���ŵ����'
        BatStatus           As Long ' ��������'
        SOC                 As String '  : ��� soc ���ٷֱ� 0-100'
        CapNow              As String ' : ��ǰ���� (0.1AH)
        CapFull             As String ' : ��������(0.1AH)
        FET_code               As Long   '       // ��Ʒ���к�[10]
        afe_Temp(4)             As String
End Type

Public Type MySys2Config
       DesignVol            As Long   '       // ϵͳ�������ѹ(mV)
       PackConfigMap        As Long   '       // MCU ϵͳ���ò���
       FCC                  As Long   '       // FullChargeCapacity ϵͳ��������(mAH)
       CycleThreshold       As Long   '       // ϵͳ����ѭ���ŵ�����(mAH)
       CycleCount           As Long   '       // ѭ���ŵ����
       NearFCC              As Long   '       // ��Ч�ŵ翪ʼʱʣ����������������������ֵ(mAH)
       LearnLowTemp         As Long   '       // ���������������������¶�
       DfilterCur           As Long   '       // ϵͳ�ܼ�����С������С����Ч�ŵ細�ڵĵ���Ϊ 0
       SWVersion            As Long   '       // ����汾��V1.00
       HWVersion            As Long   '       // Ӳ���汾��V1.00

       ShutDownDelay        As Long   '       // ����͹���ģʽ�ȴ�ʱ��(S)
       SelfDsgRate          As Long   '       // �Էŵ���(0.01%)
  '     IdleDelay            As Long   '       // ���뾲�ù���ģʽ�ȴ�ʱ��(S)
       CommOffDelay         As Long   '       // ϵͳ����λ���������͹���ģʽ����ʱ(S)
       MNFDate          As String   '       // �������ڣ�ǰ 2 �ֽڴ�š��ꡱ���� 3 �ֽڴ�š��¡����� 4 �ֽڴ�š��ա�
       MNFName          As String   '       // ManufactureName �������� ManufactureName[16]
       DeviceName       As String   '       // ��Ʒ���ƣ�DeviceName [16]
       SN               As String   '       // ��Ʒ���к�[10]
       SOH               As String   '       // ��Ʒ���к�[10]
        MCU_ID               As String   '       // ��Ʒ���к�[10]
        KEY_CODE               As String   '       // ��Ʒ���к�[10]
End Type



Public Type MybackupRecord
 
    Time_t(7)               As Long            '    uint8 * 7   ʱ��,7bit �ֱ����ꡢ�¡��ա� �ܡ�ʱ���֡���'
    RecordType              As Long            '    uint8 * 1  ��¼����
    PackStatus              As Long            '    uint16 ϵͳ״̬
    BatStatus               As Long            '    uint16 ���״̬
    FCC                     As Long            '    uint32//ϵͳ��������(mAH)
    rc                      As Long            '    uint32��ذ���ǰʣ�����(mAh)
    PackVol                 As Long            '    uint32_t��ذ��ܵ�ѹֵ(mV)
    Current                 As Long            '    int32_tʵʱ����ֵ(mA)
    RSOC                    As Long            '    uint16_t��ذ���ʣ������ٷֱ�(%)
    CellVol(16)             As Long            '    uint16_t��о 1~16 �ĵ�ѹ(mV)
    AmbientTemp             As Long            '    uint8
    PowerTemp               As Long            '    uint8
    CellTemp(7)             As Long            '    uint8;
    Protect_state           As Long            '    uint16_t��о �������״̬
End Type

