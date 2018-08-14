Attribute VB_Name = "Module1"
'#define  IPOL_STATE_Sleeping            0        //ֹͣ״̬
'#define  IPOL_STATE_Stoped          1        //����ֹͣ�������
'#define  IPOL_STATE_LineEndStoped   2        //���н���ֹͣ����ֹͣ
'#define  IPOL_STATE_Ended               3        //��End����
'#define  IPOL_STATE_Awaiting            4        //�岹��������
'#define  IPOL_STATE_FRateZero           5        //����������Ϊ0
'#define  IPOL_STATE_Suspended           6        //��������ͣ����
'#define  IPOL_STATE_Running             7        //�岹���ڽ���

'#define IPOL_DATA_BUFF_NUM    256


'#define NULLITY                         0               //0=����Ч;
'#define OPEN_LOOP_PULSE         1               //1=��������ģʽ;
'#define CLOSE_LOOP_PULSE        2               //2=λ�ñջ��������ģʽ;
'#define CLOSE_LOOP_DAV          3               //3=λ�ñջ�ģ�������ģʽ;
'#define SIMPLE_DV_OUT           4               //4=������ѹ���ģʽ;
'#define SIMPLE_PWM_OUT          5               //5=����PWM�������ģʽ;  9030�̼�
'#define AUTO_USER_SET           6               //6=�Զ��û��趨ģʽ;           ���ֻ����1��3


Declare Function InitCard_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal Axis0 As Integer, ByVal Axis1 As Integer, ByVal Axis2 As Integer, ByVal Axis3 As Integer, _
ByVal PWM_DA_Mode As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : ExitCard_9030
' ' ������� : 23

' ' ����     :�˳�9030��,���ͷ���ռ��Դ
' ' ����:
        
'          Board_NO: 0-3,���
          

' ' ����ֵ:  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ExitCard_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisWorkMode_9030
' ' ������� : 98

' ' ����     : �����Ṥ��ģʽ

' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'WorkMode:            �����Ṥ��ģʽ:
'                                                                        0=����Ч;
'                                                                        1=��������ģʽ;
'                                                                        2=λ�ñջ��������ģʽ;
'                                                                        3=λ�ñջ�ģ�������ģʽ;
'                                                                        4=������ѹ���ģʽ;
'                                                                        5=����PWM�������ģʽ;  9030�̼�
'
' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisWorkMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal WorkMode As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisKP_9030
' ' ������� : 99
'
' ' ����     : ������PID���ڱ���ϵ��

' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          Kp            : PID���ڱ���ϵ��; ��Χ: ������ Lib "dfjzh9030dll.dll" (0-65535) As Integer  ��λ: �ޡ�

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisKP_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal Axis_No As Integer, ByVal Kp As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisKI_9030
' ' ������� : 100

' ' ����     : ������PID���ڻ���ϵ��

' ' ����     :
'
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          Ki            : ��PID���ڻ���ϵ��; ��Χ: ������ Lib "dfjzh9030dll.dll" (0-65535) As Integer  ��λ: �ޡ�

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisKI_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal Ki As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisKD_9030
' ' ������� : 101

' ' ����     : ������PID����΢��ϵ��
'
' ' ����     :
'
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
'
'          Kd            : ��PID����΢��ϵ��; ��Χ: ������ Lib "dfjzh9030dll.dll" (0-65535) As Integer  ��λ: �ޡ�

 '' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisKD_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal Kd As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisIL_9030
' ' ������� : 102

' ' ����     : ������PID���ڻ�����

' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          IL            : ��PID���ڻ�����; ��Χ: ������ Lib "dfjzh9030dll.dll" (0-65535) As Integer  ��λ: �ޡ�

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisIL_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal IL As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisFVRate_9030
' ' ������� : 103

' ' ����     : ������PID�����ٶ�ǰ��ϵ��

' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'EcLine:                   ������������������  Lib "dfjzh9030dll.dll" (�ı�Ƶǰ)
'          MaxSpeed      : �������ת��;  ��λ: RPM  Lib "dfjzh9030dll.dll" (ת/����)��
'
' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFVRate_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal EcLine As Long, ByVal MaxSpeed As Double) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisFV_9030
' ' ������� : 103

' ' ����     : ������PID�����ٶ�ǰ��

' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          FV            : ��PID�����ٶ�ǰ��; ��Χ: ������ Lib "dfjzh9030dll.dll" (0-65535) As Integer  ��λ: �ޡ�

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFV_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal FV As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisPEL_9030
' ' ������� : 104

' ' ����     : ������λ�ñջ�����

' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          PosErrL  : ������λ�ñջ�����; ��Χ: ������ Lib "dfjzh9030dll.dll" (0-65535) As Integer  ��λ: �ޡ�

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisPEL_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal PosErrL As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisDAOut_9030
' ' ������� : 105
'
' ' ����     : �����������ѹ

' ' ����     :
'
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
'
'          DA_Avlue :   -10 - +10 V; ����: 1/6000; ��: ���ȴ���12λ Lib "dfjzh9030dll.dll" (4096),����13λ Lib "dfjzh9030dll.dll" (8192).
'                       ��DA�����PWM���ռ��ͬһ��Ӳ����Դ , ͨ��9030���Ӱ��ϵ�����ѡ����DA���
'                                   ����PWM���.

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisDAOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal DA_Avlue As Double) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisPWMOut_9030
' ' ������� : 106

' ' ����     : ������PWM�������
' ' ����     :
        
'          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          frequency:   18-1500000;    PWM���Ƶ��; ��λ: ���� Lib "dfjzh9030dll.dll" (Hz)
'          Pulse_Highf: ռ�ձ�,��Χ:0.0-1.0;  ����:��С��1%

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
'Declare Function SetAxisPWMOut_9030 Lib "dfjzh9030dll.dll" (ByVal    As Integer Board_NO,ByVal    As Integer Axis_No,ByVal     As Long frequency,ByVal    As Single Pulse_Highf) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : AxisPWMStop_9030
' ' ������� : 107

' ' ����   : ��PWMƵ�����ֹͣ

 '' ����   :
        
 '         Board_NO :   0��3, ���
 '          Axis_No  : 0��3, ���

 '' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 ''
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function AxisPWMStop_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : Home_9030
' ' ������� : 3

' ' ����     : ��λ������

' ' ����:
        
'          Board_NO: 0��3, ���
'          Axis_No : 0��3, ���
'
' ' ����ֵ:  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Home_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : HomeFB_9030
' ' ������� : 108

' ' ����     :  ��λ�ñ���������

' ' ����:
'
'          Board_NO: 0��3, ���
'          Axis_No : 0��3, ���

' ' ����ֵ:  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function HomeFB_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisIO_9030
' ' ������� : 15

' ' ����     :  ������IO�Ƿ���Ч
'
' ' ����     :
        
'          Board_NO: 0��3, ���
'          Axis_No : 0��3, ���
'          PHN_flag: ����λ��Home�㡢����λ��1=����λ��2=Home�㣻3=����λ
'          Mode    : 0��1��2��3, 0=�õ���Ч��3=���ɹҽ�ģʽ��1��2=��������λ���ù̶����ж�ģʽ��1��2=��Home��̶���ģʽ
'          H_L_Act : 0��1, 0=�õ�͵�ƽ��Ч,1=�õ�ߵ�ƽ��Ч
'          IO_index: 1-16, ��Mode=3ʱ Lib "dfjzh9030dll.dll" (���ɹҽ�ģʽ) As Integer�õ�ҽӵ�ͨ����������һ��

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisIO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
                                                          ByVal PHN_flag As Integer, ByVal Mode As Integer, _
                                                          ByVal H_L_Act As Integer, ByVal IO_Index As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : SetAxisIOHL_9030
' ' ������� : 85

' ' ����     :  ������IO�Ǹߵ�ƽ��Ч���ǵ͵�ƽ��Ч

' ' ����     :
        
'          Board_NO: 0��3, ���
'          Axis_No : 0��3, ���
'          PLimit  : 0��1, 0=��������λ�͵�ƽ��Ч,1=��������λ�ߵ�ƽ��Ч
'          NLimit  : 0��1, 0=���Ḻ��λ�͵�ƽ��Ч,1=���Ḻ��λ�ߵ�ƽ��Ч
'          Home    : 0��1, 0=����Home��͵�ƽ��Ч,1=����Home��ߵ�ƽ��Ч

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
'//Declare Function SetAxisIOHL_9030 Lib "dfjzh9030dll.dll" (ByVal    As Integer Board_NO,ByVal    As Integer Axis_No,ByVal    As Integer PLimit,ByVal    As Integer NLimit,ByVal    As Integer Home) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : GoHome_9030
' ' ������� : 24
'
' ' ����   : ���Home��

' ' ����   :
        
          '          Board_NO : 0��3, ���
'          Axis_No  : 0��3, ���
          
'          goHomeVel    : ���Home���ٶ�; ��Χ: ��������;  ��λ: Hz���������Ƶ�ʡ�
'          LeaveHomeVel : ���뿪Home���ٶ�; ��Χ: ��������;    ��λ: Hz���������Ƶ�ʡ�
'          LeaveHomePos : ���뿪Home�����; ��Χ: ��������,��������ֵ;  ��λ: ���������
'          LookZIndexVel: ����һת�����ٶ�; ��Χ: ��������;    ��λ: Hz���������Ƶ�ʡ�
'PulseNum:                    �����������ÿת������  Lib "dfjzh9030dll.dll" (��һת����ļ��������� Lib "dfjzh9030dll.dll" (��դ�߷���))
'          Z_IndexFlag  : 0��1;����Ϊλ�ñջ�ģʽ Lib "dfjzh9030dll.dll" (2��3ģʽ)ʱ,���������һת����ģʽ: 0=���㲻��һת����;1=������һת����

 '' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 ''
 ''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GoHome_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal goHomeVel As Long, ByVal LeaveHomeVel As Long, ByVal LeaveHomePos As Long, ByVal LookZIndexVel As Long, ByVal PulseNum As Long, ByVal Z_IndexFlag As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' ������   : LookZIndex_9030
' ' ������� : 114
'
' ' ����   : ��һת���� Lib "dfjzh9030dll.dll" (���Home��)

' ' ����   :
'
'          Board_NO : 0��3, ���
 '         Axis_No  : 0��3, ���
'
'          LookZIndexVel: ����һת�����ٶ�; ��Χ: ��������;    ��λ: Hz���������Ƶ�ʡ�
'PulseNum:                    �����������ÿת������  Lib "dfjzh9030dll.dll" (��һת����ļ��������� Lib "dfjzh9030dll.dll" (��դ�߷���))

' ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LookZIndex_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal LookZIndexVel As Long, ByVal PulseNum As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Set_Emergency_Stop_9030
 ' ������� : 29

 ' ����   : ����IO��ͣ�ź�

 ' ����   :
        
 '         Board_NO : 0��3, ���
 '         Mask     : 0��8, ��ͣ�ź���ͨ�������I1-I8����,0��ʾ��Ч,ȱʡΪ��Ч
 '         Mode     :  Lib "dfjzh9030dll.dll" (����)

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_Emergency_Stop_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Mask As Integer, ByVal Mode As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisPos_9030
 ' ������� : 4

 ' ����     : �������λ��

 ' ����     :
        
  '        Board_NO : 0��3, ���
   '       Axis_No  : 0��3, ���
  '
   '       position : ���λ��; ��Χ: ��������;  ��λ: ���������

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal position As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisVel_9030
 ' ������� : 5

 ' ����     : ��������ٶ�

 ' ����     :
        
          'Board_NO : 0��3, ���
          'Axis_No  : 0��3, ���
          
         ' velocity : ����ٶ�; ��Χ: ��������;    ��λ: Hz���������Ƶ�ʡ�

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisStartVel_9030
 ' ������� : 88

 ' ����   : ����������ٶ�

 ' ����   :
        
       '   Board_NO : 0��3, ���
       '   Axis_No  : 0��3, ���
          
       '   velocity : ��������ٶ�; ��Χ: ����������;    ��λ: Hz���������Ƶ�ʡ�

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisStartVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisStopVel_9030
 ' ������� : 95

 ' ����   : �����ֹͣ�ٶ�,�̼�3.0��

 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���
          
        '  velocity : ���ֹͣ�ٶ�; ��Χ: ����������;    ��λ: Hz���������Ƶ�ʡ�ȱʡֵ: 16

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisStopVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisDec_9030
 ' ������� : 96

 ' ����   : ����ļ��ٶ�,�̼�3.0��

 ' ����   :
        
        '  Board_NO : 0��3, ���
       '   Axis_No  : 0��3, ���
          
        '  deceleration: ��ļ��ٶ�; ��Χ: ��������,��������ֵ;  ��λ: ������ / ��ƽ����

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal deceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisStopDec_9030
 ' ������� : 122

 ' ����   : �����Stop����ļ��ٶ�,�̼�5.1��

 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          
         ' deceleration: ��ļ��ٶ�; ��Χ: ��������,��������ֵ;  ��λ: ������ / ��ƽ����

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisStopDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal deceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisAcc_9030
 ' ������� : 6

 ' ����     : ������ļ��ٶ�

 ' ����     :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���
          
       '   acceleration: ��ļ��ٶ�; ��Χ: ��������,��������ֵ;  ��λ: ������ / ��ƽ����

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisAcc_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal acceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : StartAxis_9030
 ' ������� : 7

 ' ����     : �Ὺʼ����,λ��ģʽ

 ' ����     :
        
        '  Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
'

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function StartAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������ : StopAxis_9030
 ' ������� : 9

 ' ����   : ��ֹͣ,��������ٶȼ���ֹͣ

 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function StopAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : AbortAxis_9030
 ' ������� : 19

 ' ����   : ��ֹͣ,��10����֮�ڼ���ֹͣ
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
         '

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function AbortAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : CeaseAxis_9030
 ' ������� : 20

 ' ����   : ��ֹͣ,����ֹͣ,�޼��ٹ���

 ' ����   :
        
        '  Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function CeaseAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : StartAxisVel_9030
 ' ������� : 8

 ' ����   : �������ٶ�ģʽ,���������ٶȿ�ʼ����

 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          
         ' velocity : ����ٶ�; ��Χ: ��������;  ��λ: Hz���������Ƶ�ʡ�

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function StartAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisOffset_9030
 ' ������� : 84

 ' ����   : ������λ��ƫ��ֵ

 ' ����   :
        
          'Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          
         ' offset   : ��λ�õ�ƫ��ֵ; ��Χ: ��������;  ��λ: ���������

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisOffset_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal offset As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisFBOffset_9030
 ' ������� : 109

 ' ����     :   ����λ�ñ�����ƫ��ֵ

 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          
         ' offset   : ��λ�õ�ƫ��ֵ; ��Χ: ��������;  ��λ: ���������

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFBOffset_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal offset As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisOPC_9030
 ' ������� : 112

 ' ����     :   �������ѹ���0�㲹��

 ' ����   :
        
          'Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          
         ' OPC_value: ���ѹ���0�㲹��;  ��Χ: -1000 - +1000 ��

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisOPC_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal OPC_value As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisMotorOnOff_9030
 ' ������� : 110

 ' ����     :   ������On��Off��ʹ��

 ' ����   :
        
       '   Board_NO : 0��3, ���
       '   Axis_No  : 0��3, ���
       '
       '   OnOff    : 0,1;       0=Off,1=On��

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisMotorOnOff_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal OnOff As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisOutMode_9030
 ' ������� : 17

 ' ����   : ����������ģʽ
 ' ����   :
        
       '   Board_NO : 0��3, ���
       '   Axis_No  : 0��3, ���

       '   Mode_A   : 0��1; 0=���������,1=���������
       '   Mode_B   : 0��1; 0=����-����ģʽ���,1=����-����ģʽ���
       '   Mode_C   : 0��1; 0=�᷽������,1=�᷽��ת
  

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisOutMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal Mode_A As Integer, ByVal Mode_B As Integer, ByVal Mode_C As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisSAcce_9030
 ' ������� : 72

 ' ����   : ������S�ͼ��ٶ�
 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���

       '   PowerFlag: 1-4; S�ͼ��ٶ� Lib "dfjzh9030dll.dll" (1,2,3,4)ָ��
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisSAcce_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal PowerFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisTEC_9030
 ' ������� : 73

 ' ����   : �������ݾ෴���϶����
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
        '  ErrorV   : ���ݾ෴���϶����ֵ; ��Χ: 0-32767;  ��λ: ���������
        '  TimeNum  : ��������ʱ��; ��Χ: 1-20;  ��λ: ���롣
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisTEC_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal ErrorV As Long, ByVal TimeNum As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisTECData_9030
 ' ������� : 111

 ' ����   : �������ݾ���������
 ' ����   :
        
         ' Board_NO :    0��3, ���
       '   Axis_No  :    0��3, ���
       '   Mode          :       2��3            2=�����϶+�ݾ�����;3=˫���ݾ�����
        '  EffectNum :   Mode=2ʱ��1-512��       Mode=3ʱ��0-256��       ��Ч������
        '  BasePoint :   ����λ�ã� ����TEData[0]��Ӧ��λ��,��λ: ���������
        '  NodeLen       :       ���ݼ���룬��λ: ���������
        '  Direc         :       1 �� -1������: Mode=2ʱ�� �����϶��������; Mode=3ʱ, ǰ Lib "dfjzh9030dll.dll" (˫��)һ�����ݷ���
        '  ReverseGap:   �����϶�������ݣ���Χ: 0-32767;  ��λ: ���������

        '  TEData        :       �ݾ������������飬��������̶�Ϊ512����ÿ�����ݷ�Χ: -32768 - +32767;  ��λ: ���������
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisTECData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal Mode As Integer, ByVal EffectNum As Integer, ByVal BasePoint As Long, ByVal NodeLen As Long, ByVal Direc As Integer, ByVal ReverseGap As Integer, ByVal TEData As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisTECWork_9030
 ' ������� : 111

 ' ����   : ����/ֹͣ ���ݾ�����
 ' ����   :
        
         ' Board_NO :    0��3, ���
         ' Axis_No  :    0��3, ���
         ' Work     :    0��1            0=ֹͣ�ݾ�����;1=�����ݾ�����
           

 ' ����ֵ :  0��-1��-2��        -1=���ɹ�,-2=���ݾ���������ʧЧ��0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisTECWork_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal Work As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisFE_9030
 ' ������� : 74

 ' ����   : ���������������˶�
 ' ����   :
        
        '  Board_NO : 0��3, ���
       '   Axis_No  : 0��3, ���
       '   Rate     : ��������,������; ��Χ: ����ֵ0.001-1000;  ��λ: ��;  �ֱ���: 0.001
       '   Kp       : ����PID����ϵ��; ��Χ: 0.001-1000;  ��λ: ��;  �ֱ���: 0.001
       '   Mode1    : 0=�ٶ�ģʽ,1=λ��ģʽ
        '  Mode2    : 0= ���Զ����㣬 1=�Զ�����
        '  Mode3    : 0=ֹͣ״̬���棬1=���˶�״̬����
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFE_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal Rate As Double, ByVal Kp As Double, ByVal Mode1 As Integer, ByVal Mode2 As Integer, ByVal Mode3 As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetAxisEGear_9030
 ' ������� : 75

 ' ����   : ��������ӳ����˶�
 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���
       '   F_Axis_No: 0��3, �������

       '   Rate     : ��������,������; ��Χ: ����ֵ0.001-1000;  ��λ: ��;  �ֱ���: 0.001
       '   Kp       : ����PID����ϵ��; ��Χ: 0.001-1000;  ��λ: ��;  �ֱ���: 0.001
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisEGear_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal F_Axis_No As Integer, ByVal Rate As Double, ByVal Kp As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : CancelAxisFEG_9030
 ' ������� : 76

 ' ����   : ȡ�������������˶�
 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���


 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function CancelAxisFEG_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ResetEn0Flag_9030
 ' ������� : 79

 ' ����   : �������Զ������־��0
 ' ����   :
        
         ' Board_NO : 0��3, ���


 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ResetEn0Flag_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : GetEnAuto0Flag_9030
 ' ������� : 80

 ' ����   : ��ñ������Զ������־
                        
 ' ����   :
        
 '         Board_NO : 0��3, ���
          
 ' ����ֵ :
'0:                                ��������û�б��Զ�����
'1:                                �������Զ�����
'255:                          ����ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetEnAuto0Flag_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadAxisPos_9030
 ' ������� : 2

 ' ����   : ��ȡ��ĵ�ǰλ��
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
  

 ' ����ֵ : ���λ��; ��Χ: ��������;  ��λ: ���������

         '   ��������λ��=-2147483648 ʱ,��ʾ���ɹ�,�д��������
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadAxisTheoryPos_9030
 ' ������� : 2

 ' ����   : ��ȡ��ĵ�ǰ����λ��
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
  

 ' ����ֵ : ���λ��; ��Χ: ��������;  ��λ: ���������

            '��������λ��=-2147483648 ʱ,��ʾ���ɹ�,�д��������
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisTheoryPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadAxisEncodePos_9030
 ' ������� : 2

 ' ����   : ��ȡ��ı�����λ��
 ' ����   :
        
          'Board_NO : 0��3, ���
          'Axis_No  : 0��3, ���
  

 ' ����ֵ : ���λ��; ��Χ: ��������;  ��λ: ���������

           ' ��������λ��=-2147483648 ʱ,��ʾ���ɹ�,�д��������
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisEncodePos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadAxisTECV_9030
 ' ������� : 2

 ' ����   : ��ȡ���ݾ���������
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
  

 ' ����ֵ : ����ݾ���������; ��Χ: ������;  ��λ: ���������

           ' ����������=-32768 ʱ,��ʾ���ɹ�,�д��������
 ''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisTECV_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadAxisVel_9030
 ' ������� : 26

 ' ����   : ��ȡ��ĵ�ǰ�ٶ�
 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���
          

 ' ����ֵ : ��ĵ�ǰ�ٶ�; ��Χ: ��������;    ��λ: Hz���������Ƶ�ʡ�

           ' ������=-2147483648 ʱ,��ʾ���ɹ�,�д��������
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadAxisState_9030
 ' ������� : 27

 ' ����   : ��ȡ���״̬
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���
          

 ' ����ֵ :
'1:                  �����˶���
'0:                  ����ֹͣ״̬ , ��λ�õ���, ������ Lib "dfjzh9030dll.dll" (Home_9030����)
'                -1: ��GoHome OK
''                -2: ����GoHome����ʱֹͣ
 '               -3: �ᱻStopAxis_9030����ֹͣ
'                -4: �ᱻAbortAxis_9030����ֹͣ
'                -5: �ᱻCeaseAxis_9030����ֹͣ
'                -6: �ᱻ �岹 ����ֹͣ
'                -7: �ᱻ����λֹͣ
'                -8: �ᱻ����λֹͣ
 '               -9: ��λ�üĴ��������ֹͣ
'           -10: �ᱻ�ⲿIO��ֹͣͣ
 '          -11: ���ڸ���ģʽ���ٶ�Ϊ0
'           -12: �������λ�üĴ������
'           -13: ���������
' '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadEncoderPos_9030
 ' ������� : 10

 ' ����   : ��ȡ������λ��
 ' ����   :
        
         ' Board_NO : 0��3, ���

 ' ����ֵ : ������λ��; ��Χ: ��������;  ��λ: ���������
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadEncoderPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : HomeEncode_9030
 ' ������� : 11

 ' ����   : ��λ������,������λ������
 ' ����   :
        
         ' Board_NO : 0��3, ���
            

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function HomeEncode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetEncodeCount_9030
 ' ������� : 113

 ' ����     :   �踽�ӱ�������ֵ

 ' ����   :
        
         ' Board_NO : 0��3, ���
          
          'offset   : ��������ֵ; ��Χ: ��������;  ��λ: ���������

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetEncodeCount_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal offset As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadFirmwareVersion_9030
 ' ������� : 28

 ' ����   : ��ȡ9030���ƿ��̼��汾��
 ' ����   :
        
        '  Board_NO : 0��3, ���
            

 ' ����ֵ :  0�� ����1, 0=���ɹ�, ����1=�汾��, ����:  10=1.0�汾
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadFirmwareVersion_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : GetHWID_9030
 ' ������� : 1020

 ' ����   : ��ȡ9030���ƿ�Ӳ��ID��
 ' ����   :
        
         ' Board_NO : 0��3, ���
            

 ' ����ֵ :  0�� ����1, 0=���ɹ�, ����1=ID��, 57=9030,58=9011
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetHWID_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadDllVersion_9030
 ' ������� : 67

 ' ����   : ��ȡ9030��̬���ӿ�汾��
 ' ����   :
        
  '                       Lib "dfjzh9030dll.dll" (��)

 ' ����ֵ :  0�� ����1, 0=���ɹ�, ����1=�汾��, ����:  10=1.0�汾
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadDllVersion_9030 Lib "dfjzh9030dll.dll" () As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : GetDriverVersion_9030
 ' ������� : 70

 ' ����   : ��ȡ9030���������汾��
 ' ����   :
        
         '                Lib "dfjzh9030dll.dll" (��)
'
 ' ����ֵ :  0�� ����1, 0=���ɹ�, ����1=�汾��, ����:  100=1.00�汾,����:  110=1.10�汾
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetDriverVersion_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadIO_9030
 ' ������� : 14

 ' ����   : ��ȡͨ�������I1-I20״̬
 ' ����   :
        
       '   Board_NO : 0��3, ���
            

 ' ����ֵ :  0-19λ��Ч,��Ӧ�����I1-I20״̬
 '                        -1=���ɹ�,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadIO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadIOBit_9030
 ' ������� : 14

 ' ����   : ��λ��ȡͨ�������I1-I20״̬
 ' ����   :
        
       '   Board_NO : 0�� 3,  ���
       '   Index    : 1 - 20, �����������

 ' ����ֵ :  0-1,�����״̬
 '                        -1=���ɹ�,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadIOBit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : WriteIo_9030
 ' ������� : 16

 ' ����   : ����ͨ�������O1-O8״̬
 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' IO_V     : �����ֵ,��8λ��Ч,��ӦO1-O8,8�������״̬

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function WriteIo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_V As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : WriteIoBit_9030
 ' ������� : 81

 ' ����   : ��λ����ͨ�������O1-O8״̬
 ' ����   :
        
         ' Board_NO : 0��3, ���
        '  IO_V     : 0 - 1�������ֵ
        '  Index    : 1 - 8, �����������

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function WriteIoBit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_V As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadOs_9030
 ' ������� : 82

 ' ����   : ��ȡ�����O1-O8״̬
 ' ����   :
        
         ' Board_NO : 0��3, ���
            

 ' ����ֵ :  0-7λ��Ч,��Ӧ�����O1-O8״̬

         '                -1=���ɹ�,
 ''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadOs_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadOsBit_9030
 ' ������� : 83

 ' ����   : ��λ��ȡ�����O1-O8״̬
 ' ����   :
        
       '   Board_NO : 0��3, ���
       '   Index    : 1 - 8, �����������

 ' ����ֵ :  0-1,�����״̬
        '                 -1=���ɹ�,
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadOsBit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadMPGIO_9030
 ' ������� : 118

 ' ����   : ������IO�����MPG_I1-MPG_I7״̬
 ' ����   :
        
       '   Board_NO : 0��3, ���
       '   Index    : 0=ȫ��������ֵ��bit0-bit6 ��ӦMPG_I1-MPG_I7״̬��1-7=��λ��ȡ�����ض�Ӧλ״̬

 ' ����ֵ :  Index=0ʱ 0-6λ��Ч,��Ӧ�����MPG_I1-MPG_I7״̬��Index=1-7ʱ�����ض�Ӧλ״̬ Lib "dfjzh9030dll.dll" (0��1��
       '                  -1=���ɹ�,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadMPGIO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : GetAxisMode_9030
 ' ������� : 119

 ' ����   : �����û��趨ģʽ
 ' ����   :
        
        '  Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���

 ' ����ֵ :  0��1,��Ӧ����û��趨ģʽ
         '                -1=���ɹ�,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetAxisMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : PwmOut_9030
 ' ������� : 12

 ' ����   : PWM�������
 
 ' ����   :
        
        ' Board_NO :   0��3, ���
        '  frequency:   18-1500000;    PWM���Ƶ��; ��λ: ���� Lib "dfjzh9030dll.dll" (Hz)
        '  Pulse_Highf: ռ�ձ�,��Χ:0.0-1.0;  ����:��С��1%

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function PwmOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal frequency As Long, ByVal Pulse_Highf As Single) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   :PwmOut2_9030
 ' ������� : 12

 ' ����   : PWM�������
 
 ' ����   :
        
        '  Board_NO   :   0��3, ���
       '   frequency  :   18-1500000;    PWM���Ƶ��; ��λ: ���� Lib "dfjzh9030dll.dll" (Hz)
        'Pulse_Highf:             �ߵ�ƽ����  Lib "dfjzh9030dll.dll" (ms)

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function PwmOut2_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal frequency As Long, ByVal Pulse_Highf As Single) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : PwmStop_9030
 ' ������� : 13

 ' ����   : PWM ����ֹͣ���

 ' ����   :
        
          'Board_NO :   0��3, ���

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function PwmStop_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : DAOut_9030
 ' ������� : 68

 ' ����   : DA Lib "dfjzh9030dll.dll" (��ģת��)ģ�������
 
 ' ����   :
        
         ' Board_NO :   0��3, ���
         ' DA_Avlue :   -10 - +10 V; ����: 1/6000; ��: ���ȴ���12λ Lib "dfjzh9030dll.dll" (4096),����13λ Lib "dfjzh9030dll.dll" (8192).
         '              ��DA�����PWM���ռ��ͬһ��Ӳ����Դ , ͨ��9030���Ӱ��ϵ�����ѡ����DA���
         '                          ����PWM���.

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function DAOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal DA_Avlue As Double) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetXAxis_9030
 ' ������� : 32

 ' ����   : ��ʵ������岹�����X����ƥ��

 ' ����   :
        
          'Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���

'factor_c_t:             ������嵱��?
'delta:                  ����岹��λ�����ֵ?��λΪ�û���λ?

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetXAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_OffXAxis_9030
 ' ������� : 33

 ' ����   : �����岹�����X��ƥ��
 ' ����   :
        
          'Board_NO : 0��3, ���
          

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffXAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetYAxis_9030
 ' ������� : 34

 ' ����   : ��ʵ������岹�����Y����ƥ��

 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���

'factor_c_t:             ������嵱��?
'delta:                  ����岹��λ�����ֵ?��λΪ�û���λ?

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetYAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_OffYAxis_9030
 ' ������� : 35

 ' ����   : �����岹�����Y��ƥ��
 ' ����   :
        
         ' Board_NO : 0��3, ���
          

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffYAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetZAxis_9030
 ' ������� : 36

 ' ����   : ��ʵ������岹�����Z����ƥ��

 ' ����   :
        
        '  Board_NO : 0��3, ���
        '  Axis_No  : 0��3, ���

'factor_c_t:             ������嵱��?
'delta:                  ����岹��λ�����ֵ?��λΪ�û���λ?

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetZAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_OffZAxis_9030
 ' ������� : 37

 ' ����   : �����岹�����Z��ƥ��
 ' ����   :
        
        '  Board_NO : 0��3, ���
          

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffZAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetWAxis_9030
 ' ������� : 38

 ' ����   : ��ʵ������岹�����W����ƥ��

 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 0��3, ���

'factor_c_t:             ������嵱��?
'delta:                  ����岹��λ�����ֵ?��λΪ�û���λ?

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetWAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_OffWAxis_9030
 ' ������� : 39

 ' ����   : �����岹�����W��ƥ��
 ' ����   :
        
        '  Board_NO : 0��3, ���
          

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffWAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetAxisMaxErrLtd_9030
 ' ������� : 87

 ' ����   : �������岹λ���������

 ' ����   :
        
         ' Board_NO : 0��3, ���
         ' Axis_No  : 1��4, ���־��;1=X��,2=Y��,3=Z��,4=W��
          
         ' ErrLid   :    �������岹��λ�������ֵ����λΪ�û���λ�����������ֵ,ϵͳ������.

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetAxisMaxErrLtd_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal ErrLid As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_MicroAdjustPos_9030
 ' ������� : 89

 ' ����   :   �岹��ͣ��΢����λ��

 ' ����   :
        
         ' Board_NO : 0��3, ���
        '  Axis_No  : 1��4, ���־��;1=X��,2=Y��,3=Z��,4=W��
          
'MA_Pos:               ����λ��΢��ֵ?���ֵ , ��λΪ�û���λ?

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_MicroAdjustPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal MA_Pos As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetDecMagnifyCoeff_9030
 ' ������� : 86

 ' ����   : ��岹���ٶȷŴ�ϵ��

 ' ����   :
        
         ' Board_NO : 0��3, ���
          
         ' MagnifyCoeff: 1.0 - 2.0,  Ϊ�˸��Ʋ岹����ʱ�ĳ����

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetDecMagnifyCoeff_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MagnifyCoeff As Double) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetACCDec_9030
 ' ������� : 40

 ' ����   : ���ò岹���ٶȺͼ��ٶ�

 ' ����   :
        
        '  Board_NO : 0��3, ���

        ' acceleration:  �岹���ٶȣ���Χ��1-10000����λ���û���λ / ��/ �롣ȱʡֵ: 500
        ' deceleration:  �岹���ٶȣ���Χ��1-10000����λ���û���λ / ��/ �롣ȱʡֵ: 500

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetACCDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal acceleration As Long, ByVal deceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Start_9030
 ' ������� : 41

 ' ����   : �岹��ʼ

 ' ����   :
        
         ' Board_NO : 0��3, ���

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Start_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ObligeFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetXStartPos_9030
 ' ������� : 54

 ' ����   : ��ò岹�ᵱǰλ��,Ҳ����岹�Ŀ�ʼλ��, �ڷ��Ͳ岹���� Lib "dfjzh9030dll.dll" (LM_Line_9030,IpolArc_6030)��ʼǰ����

 ' ����   :
        
        '  Board_NO : 0��3, ���
        ' Pos:                 ��岹�Ŀ�ʼλ�� , ��λΪ�û���λ?ָ�������, ����ΪNULL Lib "dfjzh9030dll.dll" (��ָ��)

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetXStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer
Declare Function LM_GetYStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer
Declare Function LM_GetZStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer
Declare Function LM_GetWStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetAxisStartPos_9030
 ' ������� : 54

 ' ����   : ��ò岹�ᵱǰλ��,Ҳ����岹�Ŀ�ʼλ��,                     �ڷ��Ͳ岹���� Lib "dfjzh9030dll.dll" (LM_Line_9030,IpolArc_6030)��ʼǰ����

 ' ����   :
        
          'Board_NO : 0��3, ���
          'AxisFlag : 1- 4, ָʾ���ĸ��᣻1��X�ᣬ 2��Y�ᣬ3��Z�ᣬ4��W�ᡣ

 ' ����ֵ :  ��ĵ�ǰλ�á��û���λ
           '              ��������λ��=-2147483648 ʱ,��ʾ���ɹ�,�д��������
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetAxisStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal AxisFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Line_9030
 ' ������� : 42

 ' ����   : ֱ�߲岹
        '                ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��
'
 ' ����   :
        
      '    Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
 '         Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Line_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_LineMaxV_9030
 ' ������� : 69

 ' ����   : ֱ�߲岹
  '                      ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
  '        Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'          MaxSpeed : ���в岹����ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineMaxV_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal Speed As Double, _
ByVal MaxSpeed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_LineMeasure_9030
 ' ������� : 65

 ' ����   : ֱ�߲岹������IO������
   '                     ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          IO_Index : 1-8;  ͨ��IO����������,
'          Mode     : 1,2,3,4��ģʽ
'                                                                1,2:ͨ��IO�������Ϊ1�� 1=����ֹͣ��2=10ms����ֹͣ
'                                                                3,4:ͨ��IO�������Ϊ0�� 3=����ֹͣ��4=10ms����ֹͣ
          
'          Speed    : �岹�ٶ�,Ҳ�Ǳ��в岹����ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineMeasure_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal IO_Index As Integer, ByVal Mode As Integer, _
ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_LineFE_9030
 ' ������� : 77

 ' ����   : ֱ�߲岹 ���������λ��
 '                       ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
  '        Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'
'          Mode     : 0,1; 0=����Ŀ��,����ֹͣ����; 1=����Ŀ��,��ֹͣ����,�к����������;
'          Rate     : ��������,������;
                                 
'                                 ��Χ: �ɸ���������嵱������, ����ֵ��Χ: 0.001'����������嵱�� - 1000 '����������嵱�� ;
'                                 ��λ: ��;  �ֱ���: 0.001'����������嵱��
'
'          Kp       : ����PID����ϵ��; ��Χ: 0.001-1000;  ��λ: ��;  �ֱ���: 0.001
'          Speed    : �岹Ԥ���ٶ�,Ҳ�Ǳ��в岹����ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineFE_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal Rate As Double, ByVal Kp As Double, ByVal Mode As Integer, ByVal Speed As Double, ByVal LineNO As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_ArcCW_9030
 ' ������� : 43

 ' ����   : Բ���岹,˳Բ
'                        ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'xcPos:               x���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'ycPos:               y���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨

'                                ��z���w���������յ����겻�غ�ʱ,z���w����xy�����������˶�,Z��w��ͬʱ���������˶�.
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCW_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal xcPos As Double, ByVal ycPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_ArcCW_ZX_9030
 ' ������� : 43

 ' ����   : Բ���岹,˳Բ,��ZXƽ��
  '                      ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zcPos:               z���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'xcPos:               x���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨

'                                ��y���w���������յ����겻�غ�ʱ,y���w����zx�����������˶�,y��w��ͬʱ���������˶�.
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCW_ZX_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal zcPos As Double, ByVal xcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_ArcCW_YZ_9030
 ' ������� : 43

 ' ����   : Բ���岹,˳Բ,��YZƽ��
   '                     ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'ycPos:               y���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zcPos:               z���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨

'                                ��x���w���������յ����겻�غ�ʱ,x���w����yz�����������˶�,x��w��ͬʱ���������˶�.
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCW_YZ_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal ycPos As Double, ByVal zcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_ArcCCW_9030
 ' ������� : 44

 ' ����   : Բ���岹,��Բ
 '                       ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'xcPos:               x���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'ycPos:               y���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨

 '                               ��z���w���������յ����겻�غ�ʱ,z���w����xy�����������˶�,Z��w��ͬʱ���������˶�.
'
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCCW_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal xcPos As Double, ByVal ycPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_ArcCCW_ZX_9030
 ' ������� : 44

 ' ����   : Բ���岹,��Բ,��ZXƽ��
 '                       ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zcPos:               z���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'xcPos:               x���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨

'                                ��y���w���������յ����겻�غ�ʱ,y���w����xy�����������˶�,y��w��ͬʱ���������˶�.
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCCW_ZX_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal zcPos As Double, ByVal xcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_ArcCCW_YZ_9030
 ' ������� : 44

 ' ����   : Բ���岹,��Բ,��YZƽ��
 '                       ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��

 ' ����   :
 '
 '         Board_NO : 0��3, ���

'xPos:                x��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'ycPos:               y���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zcPos:               z���Բ��λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
 '         Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨

'                                ��x���w���������յ����겻�غ�ʱ,x���w����xy�����������˶�,x��w��ͬʱ���������˶�.
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCCW_YZ_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal ycPos As Double, ByVal zcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_End_9030
 ' ������� : 45

 ' ����   : �岹����
           '             ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��
          '              ���岹�������е�����ʱ��ֹͣ���� , ��Ҫ���¿�ʼ, ��������岹��������
          '               Lib "dfjzh9030dll.dll" (LM_Line_9030,LM_ArcCW_9030,LM_ArcCCW_9030),����LM_Start_9030����

 ' ����   :
        
       '   Board_NO : 0��3, ���
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_End_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Wait_9030
 ' ������� : 60

 ' ����   : �岹�ȴ�
           '             ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��
           '             ���岹�������е�����ʱ����ͣ����,�ڵȴ��û�����ʱ���,�����¿�ʼ.

 ' ����   :
        
      '    Board_NO    : 0��3, ���
'Millisecond:            �ȴ�ʱ�� , ����
'LineNO:                 �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Wait_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Millisecond As Long, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_PWM_9030
 ' ������� : 61

 ' ����   : �岹ģʽ��PWM���ռ�ձ����ٶȱ仯
  '                      ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��
  '                      ���岹�������е�����ʱ�����ݲ�������PWM���.

 ' ����   :
        
 '         Board_NO    : 0��3, ���
  '        frequency   : 18-1500000;    PWM���Ƶ��; ��λ: ���� Lib "dfjzh9030dll.dll" (Hz)
 '                       ��Ϊ0ʱ , ֹͣPWM���
 '         Pulse_Highf : ռ�ձ�,��Χ:0.0-1.0;  ����:��С��1%
  '                      ��ֵ��Ӧ�岹�ٶȴﵽSpeed�����ٶ�ʱ�����ռ�ձ� , ���岹�ٶȴ���
 '                                       Speed�����ٶ�ʱ,Ҳ���ò������.
'
 '         Speed       : ռ�ձ��涯�����岹�ٶ�,��λ: �û���λ/����
  '                      ����ֵΪ0ʱ , ���������涯ģʽ, ���岹�������е�����ʱ
  '                                      ������frequency,Pulse_Highf����ֵ����PWM���.
'LineNO:                 �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_PWM_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal frequency As Long, ByVal Pulse_Highf As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_IOOut_9030
 ' ������� : 62

 ' ����   : �岹ģʽ��IO�����
     '                   ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��
   '                     ���岹�������е�����ʱ�����ݲ������IO��.

 ' ����   :
        
   '       Board_NO    : 0��3, ���
    '      IO_Index    : 1-8;  ͨ��IO����������,
     '     IO_Value    : 0 - 1; �����ֵ
'LineNO:                 �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_IOOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_Index As Integer, ByVal IO_Value As Integer, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Wait_I_9030
 ' ������� : 63

 ' ����   : �岹�ȴ�ͨ��IO������
   '                     ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���16��
  '                      ���岹�������е�����ʱ���ȴ�ֱ�������Ϊ1.

 ' ����   :
        
'          Board_NO    : 0��3, ���
'          IO_Index    : 1-8;  ͨ��IO����������,
'LineNO:                 �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Wait_I_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_Index As Integer, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_CleanBuff_9030
 ' ������� : 46

 ' ����   : ����岹����
                        
 ' ����   :
        
 '         Board_NO : 0��3, ���
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_CleanBuff_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetMinVel_9030
 ' ������� : 47

 ' ����   : ���ò岹��С�ٶ�
                        
 ' ����   :
        
 '         Board_NO  : 0��3, ���
 '         MinLineVel: ֱ����С�岹�ٶ�, ��λ: �û���λ/����, ȱʡֵ: 30
 '         MinArcVel : Բ����С�岹�ٶ�, ��λ: �û���λ/����, ȱʡֵ: 30
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetMinVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MinLineVel As Double, ByVal MinArcVel As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetMaxVel_9030
 ' ������� : 57

 ' ����   : ���ò岹����ٶ�
                        
 ' ����   :
        
       '   Board_NO  : 0��3, ���
       '   MaxLineVel: ֱ�����岹�ٶ�, ��λ: �û���λ/����, ȱʡֵ: 4000
       '   MaxArcVel : Բ�����岹�ٶ�, ��λ: �û���λ/����, ȱʡֵ: 4000
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetMaxVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MaxLineVel As Double, ByVal MaxArcVel As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetSpeedPri_9030
 ' ������� : 58

 ' ����   : ���ò岹�ٶ��ٶ����Ȼ��Ǿ�������
                        
 ' ����   :
        
'          Board_NO    : 0��3, ���
 '         PriorityFlag: 0-1; 1:�岹�ٶ�����, 0:�岹�������ȡ�ȱʡֵ: 1 �岹�ٶ�����
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSpeedPri_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal PriorityFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetSAccePower_9030
 ' ������� : 71

 ' ����   : ���ò岹S�ͼ��ٶ� Lib "dfjzh9030dll.dll" (1,2,3,4)ָ��
                        
 ' ����   :
        
'          Board_NO    : 0��3, ���
  '        PowerFlag   : 1-4;  S�ͼӼ��� Lib "dfjzh9030dll.dll" (ָ������)��ָ��
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSAccePower_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal PowerFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetParaAngle_9030
 ' ������� : 48

 ' ����   : ��岹����,�Ƕ�
                        
 ' ����   :
        
  ''        Board_NO : 0��3, ���
'angle1:
'angle2:
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetParaAngle_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal angle1 As Integer, ByVal angle2 As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetSysPara_9030
 ' ������� : 59

 ' ����   : ��岹ϵͳ����
                        
 ' ����   :
  '
  '        Board_NO : 0��3, ���
 '         MinLength: ϵͳ��С����,��λ: �û���λ, ȱʡֵ: 0.001;��ֱ�߻�Բ���岹ʱ,ֱ�߻�Բ������
 '                    ����С�ڸ��趨ֵ
 '         MinSpeed : ϵͳ��С�ٶ�,��λ: �û���λ/����, ȱʡֵ: 0.001;��ֱ�߻�Բ���岹ʱ,ֱ�߻�Բ��
 '                                �岹�ٶȲ���С�ڸ��趨ֵ
  '        ArcError : ϵͳ���Բ�����,��λ: �û���λ, ȱʡֵ: 0.2;��Բ���岹ʱ,Բ�������յ�뾶
 '                    ���ܴ��ڸ��趨ֵ
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSysPara_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MinLength As Double, ByVal MinSpeed As Double, ByVal ArcError As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetBuffLen_9030
 ' ������� : 55

 ' ����   : ��ò岹����ʣ�೤��
                        
 ' ����   :
        
 '         Board_NO : 0��3, ���
          
 ' ����ֵ :  0-32,�岹���泤��
 '            -1=���ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetBuffLen_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetState_9030
 ' ������� : 56

 ' ����   : ��ò岹״̬
                        
 ' ����   :
        
 '         Board_NO : 0��3, ���
          
 ' ����ֵ :
'0:                                ֹͣ״̬
'1:                                ����ֹͣ�������  Lib "dfjzh9030dll.dll" (������Abort_9030����)
'2:                                ���н���ֹͣ����ֹͣ
'3:                                ��LM_End_9030����
'4:                                �岹��������
'5:                                ����������Ϊ0��ͣ
'6:                                ��������ͣ����
'7:                                �岹���ڽ���
'255:                  ����ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetMeasureState_9030
 ' ������� : 66

 ' ����   : ��ò岹����״̬
                        
 ' ����   :
        
'          Board_NO : 0��3, ���
          
 ' ����ֵ :
'0:                                û�м�⵽�����ź�
'1:                                ��⵽�����ź�
'255:                          ����ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetMeasureState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetFEnState_9030
 ' ������� : 78

 ' ����   : ���������������ѴﵽĿ��״̬
                        
 ' ����   :
        
'          Board_NO : 0��3, ���
          
 ' ����ֵ :
'0:                                ���滹û�дﵽĿ��
'1:                                �����ѴﵽĿ��
'255:                          ����ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetFEnState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetLineNO_9030
 ' ������� : 49

 ' ����   : ��ò岹��ǰ�к�
 ' ����   :
        
 '         Board_NO : 0��3, ���

 ' ����ֵ : ��ǰ�к�; ��Χ: ��������;  �û��趨ֵ��

 '           ������ֵ=-2147483648 ʱ,��ʾ���ɹ�,�д��������
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetLineNO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Pause_9030
 ' ������� : 50

 ' ����   :  �岹��ͣ,�岹������10����֮�ڼ���ֹͣ
                        
 ' ����   :
        
'          Board_NO : 0��3, ���
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Pause_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Resume_9030
 ' ������� : 51

 ' ����   :  �ָ��岹��ͣ,�岹�����Բ岹���ٶȻָ�����
                        
 ' ����   :
        
        '  Board_NO : 0��3, ���
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Resume_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetSpeedRate_9030
 ' ������� : 52

 ' ����   :  ��岹����
                        
 ' ����   :
        
      '    Board_NO : 0��3,  ���
      '    Rate     : 1-160, 100=100%,��:��ԭ�趨�ٶ�ִ��;10=10%,��:��ԭ�趨�ٶȵİٷ�֮ʮִ��,
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSpeedRate_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Rate As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_LineEnd_9030
 ' ������� : 53

 ' ����   : �岹�����굱ǰ��ֹͣ
       '                 ���������岹������ , ���ڲ岹����ʱ, ��ʱִ��
     '                   ���岹���������굱ǰ��ֹͣʱ , ��Ҫ���¿�ʼ, ����ִ��LM_Start_9030����

 ' ����   :
        
   '      Board_NO : 0��3, ���
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineEnd_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetForceCtrl_9030
 ' ������� : 94

 ' ����   : ���ò岹������ģʽ,�̼�3.0��
                        
 ' ����   :
        
       '   Board_NO    : 0��3, ���
      '    ForceFlag   : 0-1;  0:�岹��������Ч,1=�岹��������Ч
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetForceCtrl_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ForceFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetNurbsScanMode_9030
 ' ������� : 1000

 ' ����   : ����Nurbs���߲岹Ԥ�����ɨ��ģʽ,�̼�3.0��

 ' ����   :
        
        '  Board_NO              : 0��3, ���

       '   Mode                  :       Nurbs���߲岹Ԥ��ɨ���ٶ�ģʽ: 0=�����ٶ�ɨ��,1=�������ٶȰٷֱ�ɨ��
'ScanSpeed:                              ������ɨ���ٶ�
      '    ScanSpeedRate :       �������ٶȰٷֱȵı�ֵ: ��Χ: 1-100
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsScanMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Mode As Integer, ByVal ScanSpeed As Double, ByVal ScanSpeedRate As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetNurbsVelCtrl_9030
 ' ������� : 1001

 ' ����   : ����Nurbs���߲岹ʱ���ٶȿ���,�̼�3.0��

 ' ����   :
        
      '    Board_NO              : 0��3, ���

     '     BSErrEnable   :       0=�����Ҹ������Ч,1=�����Ҹ������Ч
     '     BSErrV                :       ����Ҹ����ֵ,��Χ: 0.0001-10.0, ��λ: �û���λ
  
      '    RAccEnable    :       0=���Ʒ�����ٶ���Ч,1=���Ʒ�����ٶ���Ч
     '     RAccV                 :       �������ٶ�ֵ,��Χ: 1-1000000,       ��λ: �û���λ/��/��
 

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsVelCtrl_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal BSErrEnable As Integer, ByVal BSErrV As Double, ByVal RAccEnable As Integer, ByVal RAccV As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetNurbsAccDec_9030
 ' ������� : 93

 ' ����   : ����Nurbs���߲岹���ٶȺͼ��ٶ�,�̼�3.0��

 ' ����   :
        
       '   Board_NO : 0��3, ���

     '     Nurbs_Acc:    �岹���ٶȣ���Χ��1-10000����λ���û���λ / ��/ �롣ȱʡֵ: 500
      '    Nurbs_Dec:    �岹���ٶȣ���Χ��1-10000����λ���û���λ / ��/ �롣ȱʡֵ: 500

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsAccDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Nurbs_Acc As Double, ByVal Nurbs_Dec As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetNurbsCompCoef_9030
 ' ������� : 97

 ' ����   : ��Nurbs���߲岹����ϵ��  �̼�3.0��

 ' ����   :
        
      '    Board_NO : 0��3, ���

     '     Coef: ��Nurbs���߲岹����ϵ������Χ��0.0-0.05����λ���ޡ�ȱʡֵ: 0.01

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsCompCoef_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Coef As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Nurbs_9030
 ' ������� : 92

 ' ����   : Nurbs���߲岹,�̼�3.0��
  '                      ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���64��

 ' ����   :
        
     '     Board_NO : 0��3, ���

'knot1:               �ڵ�ֵ1?
'knot2:               �ڵ�ֵ2?
'knot3:               �ڵ�ֵ3?
'knot4:               �ڵ�ֵ4?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Nurbs_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal knot1 As Double, ByVal knot2 As Double, ByVal knot3 As Double, ByVal knot4 As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_Nurbs4Axis_9030
 ' ������� : 92

 ' ����   : 4��Nurbs���߲岹,�̼�3.0��
      '                 ������Ѳ岹��������9030���Ĳ岹��������,9030���Ĳ岹�������ܹ���64��

 ' ����   :
        
 '         Board_NO : 0��3, ���

'knot1:               �ڵ�ֵ1?
'knot2:               �ڵ�ֵ2?
'knot3:               �ڵ�ֵ3?
'knot4:               �ڵ�ֵ4?
'wPos:                w��岹���յ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'          Speed    : �岹�ٶ�,��λ: �û���λ/����
'LineNO:              �岹�к� , �û������趨
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Nurbs4Axis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal knot1 As Double, ByVal knot2 As Double, ByVal knot3 As Double, ByVal knot4 As Double, ByVal wPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_NurbsData_9030
 ' ������� : 91

 ' ����   : ��Nurbs���߲岹����,�̼�3.0��

 ' ����   :
        
'          Board_NO : 0��3, ���

'xPos:                X��Ŀ��Ƶ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��Ŀ��Ƶ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��Ŀ��Ƶ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'knot:                �ڵ�ֵ?
'weight:              Ȩֵ?
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_NurbsData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal knot As Double, ByVal weight As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_NurbsInit_9030
 ' ������� : 90

 ' ����     : ��ʼ��Nurbs����,ΪNurbs���߲岹��׼��

 ' ����   :
        
       '   Board_NO : 0��3, ���

       '   _deg     :  Lib "dfjzh9030dll.dll" (��������,�����3) 3=����Nurbs����
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_NurbsInit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal deg As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SendNurbsData_9030
 ' ������� : 1013

 ' ����   :   ��9030��ת��NURBS��������,�̼�3.0��
                        
 ' ����   :
        
 '         Board_NO    : 0��3, ���
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SendNurbsData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetNubrsExecPara_9030
 ' ������� : 1012

 ' ����   : ��ȡ9030��NURBS���߲岹���в���ֵ
 ' ����   :
        
       '   Board_NO : 0��3, ���
            

 ' ����ֵ :  ��ǰ9030��NURBS���߲岹���в���ֵ
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetNubrsExecPara_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Single

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetFactVel_9030
 ' ������� : 1002

 ' ����   :       ���ʵ�ʲ岹�ٶ�                      ��̬���ӿ�3.0��
 ' ����   :
        
        '  Board_NO : 0��3, ���
            

 ' ����ֵ :  ʵ�ʲ岹�ٶ�       ��λ: �û���λ/����
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetFactVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Single

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetNurbsInBuffLen_9030
 ' ������� : 1003

 ' ����   : ���Nurbs�����ڲ岹����������               ��̬���ӿ�3.0��
                        
 ' ����   :
        
       ' Board_NO : 0��3, ���
       '
 ' ����ֵ :  0-64,Nurbs�����ڲ岹����������
       '      -1=���ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetNurbsInBuffLen_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Set_NurbsInit_9030
 ' ������� : 1004

 ' ����     : ��ʼ��Nurbs����,ΪNurbs���߼�����׼��

 ' ����   :
        
        '  Nurbs_NO : 0��7, Nurbs���ߺ�

       '   _deg     :  Lib "dfjzh9030dll.dll" (��������,�����3) 3=����Nurbs����
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_NurbsInit_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal deg As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Set_NurbsData_9030
 ' ������� : 1005

 ' ����   : ��Nurbs���߼�������,��̬���ӿ�3.0��

 ' ����   :
        
     '     Nurbs_NO : 0��7, Nurbs���ߺ�

'xPos:                X��Ŀ��Ƶ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'yPos:                y��Ŀ��Ƶ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'zPos:                z��Ŀ��Ƶ�λ��  Lib "dfjzh9030dll.dll" (����ֵ), ��λΪ�û���λ?
'knot:                �ڵ�ֵ?
'weight:              Ȩֵ?
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_NurbsData_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal knot As Double, ByVal weight As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Set_NurbsEnd_9030
 ' ������� : 1006

 ' ����   : ��Nurbs�������ݽ���,��̬���ӿ�3.0��

 ' ����   :
        
    '     Nurbs_NO : 0��7, Nurbs���ߺ�

'knot1:               �ڵ�ֵ1?
'knot2:               �ڵ�ֵ2?
'knot3:               �ڵ�ֵ3?
'knot4:               �ڵ�ֵ4?
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_NurbsEnd_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, _
ByVal knot1 As Double, ByVal knot2 As Double, ByVal knot3 As Double, ByVal knot4 As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Get_NurbsPos_9030
 ' ������� : 1007

 ' ����   : ����Nurbs������λ�� Lib "dfjzh9030dll.dll" (��ֵ��)����,��̬���ӿ�3.0��

 ' ����   :
        
    '      Nurbs_NO : 0��7, Nurbs���ߺ�

    '      Up       : Nurbs���߲�������,��Χ: 0-  Lib "dfjzh9030dll.dll" (���Ƶ���-3)��
'xPos:                X��λ��ָ��?
'yPos:                Y��λ��ָ��?
'zPos:                Z��λ��ָ��?
          
 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsPos_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal Up As Double, ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Get_NurbsPosVB_9030
 ' ������� : 1007

 ' ����   : ����Nurbs������λ�� Lib "dfjzh9030dll.dll" (��ֵ��)����,��̬���ӿ�3.0��

 ' ����   :
        
      '    Nurbs_NO : 0��7, Nurbs���ߺ�

    '      Up       : Nurbs���߲�������,��Χ: 0-  Lib "dfjzh9030dll.dll" (���Ƶ���-3)��
'xflag:               Ϊ1ʱ , ָʾ����X��λ��?
'yflag:               Ϊ1ʱ , ָʾ����Y��λ��?
'zflag:               Ϊ1ʱ , ָʾ����Z��λ��?
          
 ' ����ֵ :  Nurbs�����ڲ�������ΪUp����λ��
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsPosVB_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal Up As Double, ByVal xflag As Integer, ByVal yflag As Integer, ByVal zflag As Integer) As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Get_NurbsLen_9030
 ' ������� : 1009

 ' ����   : ����Nurbs���߳���,��̬���ӿ�3.0��

 ' ����   :
        
      '    Nurbs_NO : 0��7, Nurbs���ߺ�

       '   Up       : Nurbs���߲�������,��Χ: 0-  Lib "dfjzh9030dll.dll" (���Ƶ���-3)��

          
 ' ����ֵ :  Nurbs�����ڲ�������ΪUp�ĳ���
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsLen_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal Up As Double) As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : Get_NurbsErrorNo_9030
 ' ������� : 1011

 ' ����   : ��ò���Nurbs���� Lib "dfjzh9030dll.dll" (����)�Ĵ����             ��̬���ӿ�3.0��
                        
 ' ����   :
        
    '      Nurbs_NO : 0��7, Nurbs���ߺ�
          
 ' ����ֵ :
'Nurbs�����:
'
'                        1=�ڴ治��
'                        2=�������㷶Χ
'                        3=�����߼�����
'                        4=ɨ���ٶȹ���
''                        5=���ݲ�����
'                        6=�ڵ�ʸ�������ݳ�����ѧ��ʽ�Ķ���
'                        7=���Ƶ�����������4
'                        8=��������������Χ
 '                       9=������ƿ���ʼ�����ɹ�
'                        -1=���ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_GetNurbsErrorNo_9030
 ' ������� : 1010

 ' ����   : ��ò���Nurbs���� Lib "dfjzh9030dll.dll" (�岹)�Ĵ����             ��̬���ӿ�3.0��
                        
 ' ����   :
        
 '         Board_NO : 0��3, ���
          
 ' ����ֵ :
'Nurbs�����:

'                        1=�ڴ治��
'                        2=�������㷶Χ
'                        3=�����߼�����
'                        4=ɨ���ٶȹ���
'                        5=���ݲ�����
 '                       6=�ڵ�ʸ�������ݳ�����ѧ��ʽ�Ķ���
''                        7=���Ƶ�����������4
 '                       8=��������������Χ
'                        -1=���ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetNurbsErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : GetErrorNo_9030
 ' ������� : 100

 ' ����   : ��ô����
 '                       �����ú����ķ���ֵ Ϊ"���ɹ�"ʱ,������������ʱ,
 '                       �ɵ��øú�����ô�����Ϣ
''
 ' ����   :
        
'          Board_NO  : 0��3, ���
'          CleanFlag : 0��1;  0=���������;  1=�������;
 '         'ErrorNo  : ���ش�����                ָ�������,����ΪNULL Lib "dfjzh9030dll.dll" (��ָ��)
 '         'FuncNo   : ���ز�������ĺ������      ָ�������,����ΪNULL Lib "dfjzh9030dll.dll" (��ָ��)
          
 ' ����ֵ :  0��1, 0=�޴���,1=�д���
 '
 ' �����ű�:

'1                   ��λ�üĴ�������
'2                   ���ڲ岹�˶���
'3                   �岹����
'4                   �����˶���
'5                   �߼�����
'6                   �岹��ѹ����λ��

'7                   �ڴ治��
'8                   ��������
'9                   �岹��������
'10              ������ͨѶʧ��
'11              ������ͨѶ��ʱ
'12              �岹����С��ϵͳ��С�ٶ�
'13              �岹����С��ϵͳ��С����
'14              �岹���ݴ���ϵͳԲ���뾶���
'15              �岹����Բ���������
'           16   ����ϵͳ���ֵ��С��ϵͳ��Сֵ:  ϵͳֵ��Χ:  ��λ��:       -2147483648 �� 2147483647
'                                                                                                                  �岹��������: <1073741823      Lib "dfjzh9030dll.dll" (ֱ�߳��Ⱥ�Բ������)
'                                                                                                                  Բ�����뾶: <1048575
'                                                                                                                  �岹Բ��λ��: -2147483648 �� 2147483647
'           17   9030��̬���ӿ�汾��9030���Ĺ̼��汾��һ��
'18              �в岹�������� , ���������ټ���岹����
'19              ���ڸ����˶�
'20              ����ֹͣ����
'21                  Nurbs���߼��������ڴ治��
'22                  Nurbs���߳������㷶Χ
'23                  Nurbs���߼����߼�����
'24                  Nurbs����ɨ���ٶȹ���
'25                  Nurbs�������ݲ�����
'26                  Nurbs���ߵ���㲻����
'27                  Nurbs���ߵ�����û�м�ʱ���뿨��
'
'
'28                  ���90300���Ĺ̼��汾��һ��
'                29  9030���̼��汾̫��

'31                              �����������
'32                              ���ղ岹���ݴ���
'33                              ����NURBS���ݴ���
'34                              �����Ǳջ�ģʽʱ , ���ܵ�������λ�ñ�����
'35                              �����������λ�üĴ�������
'36                              ��λ�ø�������
'37                              ����һת����ʧ��
'38                              ��Home����Ч , GoHomeʧЧ


 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal CleanFlag As Integer, ByVal ErrorNo As Integer, ByVal FuncNo As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : DM_SetAxisVel_9030
 ' ������� : 21

 ' ���� :��ֱ���ٶ����ģʽ
 '  input       :
 '                               num          ���1��4
 '                               velocity      ���ٶ�ģʽ
 '
 '  output      :none
 '  return      :none
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function DM_SetAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal position As Long, ByVal velocity As Long, ByVal direction As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : DM_SetAxisPos_9030
 ' ������� : 22

 ' ���� :��ֱ���ٶ����ģʽ
 '  input       :
 '                               num          ���1��4
 '                               velocity      ���ٶ�ģʽ
 '
 '  output      :none
 '  return      :none
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function DM_SetAxisPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal position As Long, ByVal velocity As Long, ByVal direction As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : RegCANExp_9030
 ' ������� : 2000

 ' ����     :�Ǽ�CAN������չ��
 ' ����:
        
 '         Board_NO: 0-3,���
'ID:                 CAN������չ��ID��
'CardType:           ��չ���ͺ�


 ' ����ֵ:  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function RegCANExp_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long, ByVal CardType As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : EnableCANExp_9030
 ' ������� : 116

 ' ����     : CAN������չ��ʹ�� Lib "dfjzh9030dll.dll" (��ʼ������)
 ' ����:
        
      '    Board_NO: 0-3,���

 ' ����ֵ:  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function EnableCANExp_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SendCANData_9030
 ' ������� : 117

 ' ����     : CAN������չ����������  Lib "dfjzh9030dll.dll" (���䷢����Ϣ)
 ' ����:
        
'          Board_NO: 0-3,���
'ID:                 ��չ����ID��
'CardType:           ��չ���ͺ�
'D1234:              ��4�ֽ�
'D5678:              ��4�ֽ�

 ' ����ֵ:  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SendCANData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long, ByVal CardType As Integer, ByVal D1234 As Long, ByVal D5678 As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadCANL_9030
 ' ������� : 2001

 ' ����   : ��ȡCAN������չ������ ��4λ
 ' ����   :
        
  '        Board_NO : 0��3, ���
'ID:                 ��չ����ID��
            

 ' ����ֵ :  0-7��8-15��16-23��24-31λ,�ֱ��Ӧ��1��2��3��4�ֽ�
                         
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadCANL_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : ReadCANH_9030
 ' ������� : 2001

 ' ����   : ��ȡCAN������չ������ ��4λ
 ' ����   :
        
  '        Board_NO : 0��3, ���
'ID:                 ��չ����ID��
            

 ' ����ֵ :  0-7��8-15��16-23��24-31λ,�ֱ��Ӧ��5��6��7��8�ֽ�
                         
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadCANH_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : GetCANErrorNo_9030
 ' ������� : 2002

 ' ����   : ���CAN���ߴ����
  '                      �����ú����ķ���ֵ Ϊ"���ɹ�"ʱ,������������ʱ,
  '                      �ɵ��øú�����ô�����Ϣ

 ' ����   :
        
  '        Board_NO  : 0��3, ���
  '        CleanFlag : 0��1;  0=���������;  1=�������;

          
 ' ����ֵ :  0-9, 0=�޴���,�����
 '
 ' �����ű�:

'1                                       �����Խ��
'2                                       ���ݳ���Խ��
'3                                       �����Ѿ�ռ��
'4                                       �����ID���ظ�
'5                                       CAN����ֻ��ʹ��һ��
'6                                       ���䲻ƥ��
'                7                       CAN���߽������ݳ�������  Lib "dfjzh9030dll.dll"  Lib "dfjzh9030dll.dll" (>64�ֽ�)
'8                                       ��������
'9                                       CAN ���߽���������Ч

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetCANErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal CleanFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetLineArcCompCoef_9030
 ' ������� : 120

 ' ����   : ��ֱ��/Բ���岹����ϵ��  �̼�3.7������

 ' ����   :
        
   '       Board_NO : 0��3, ���
   '       flag     : 0-1   0=ֱ�߲岹,1=Բ���岹
   '       Coef: ��ֱ��/Բ���岹����ϵ����ֱ�߲岹��Χ��0.0,0.03-0.08����λ���ޡ�ȱʡֵ: 0.00
    '                                                                             Բ���岹��Χ��0.0,0.03��         ��λ���ޡ�ȱʡֵ: 0.00

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetLineArcCompCoef_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal flag As Integer, ByVal Coef As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : SetActiveEncoder_9030
 ' ������� : 121

 ' ����   : �������������˶�֮ǰ,��������������
 ' ����   :
        
 '         Board_NO : 0��3, ���
 '         Axis_No  : 0��4, 0-3��Ӧ���,4=���ӱ�����,ȱʡΪ4
           

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetActiveEncoder_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' ������   : LM_SetIpolCompCoef_9030
 ' ������� : 120

 ' ����   : ��岹����ϵ��  �̼�3.6������

 ' ����   :
    
 '     Board_NO : 0��3, ���
 '     flag1    : 0-3   0=ֱ�߲岹,1=ֱ�߲岹,2=Բ���岹,3=Nurbs���߲岹
 '     flag2    : 0-2   0=ֱ�߲岹,1=ֱ�߲岹,2=Բ���岹,3=Nurbs���߲岹
 '     Coef: ��ֱ��/Բ���岹����ϵ����ֱ�߲岹��Χ��0.0,0.03-0.08����λ���ޡ�ȱʡֵ: 0.00
 '                                        Բ���岹��Χ��0.0,0.03��     ��λ���ޡ�ȱʡֵ: 0.00

 ' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/

Declare Function LM_SetIpolCompCoef_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal flag1 As Integer, ByVal flag2 As Integer, ByVal Coef As Double) As Integer

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/

' ������   : LM_SetSysDelay_9030
' ������� :

' ����   : ��岹ϵͳ����
            
' ����   :
    
'      Board_NO : 0��3, ���
'      DelayFlag: 0,1,  0=����ʱ; 1=��ʱ1����  ���岹����ʣ�೤��ʱ��ʱ��־.  ȱʡֵ: 1;
      
      
' ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
'
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Declare Function LM_SetSysDelay_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal DelayFlag As Integer) As Integer

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
' * ������   : UnlockFlash_9030
' * ������� : 126

' * ����   : 9030���� �û�flash ����
' * ����   :
    
'      Board_NO : 0��3, ���
'      password1: ������,�������룬ȱʡΪ0
'      password2: ������,�������룬ȱʡΪ0
       

' * ����ֵ :  1��0��-1,0=���ɹ�,1=�ɹ���-1=����ʧ��
' *
' *
' *****************************************************************************/
'short APIENTRY UnlockFlash_9030(unsigned short Board_NO,unsigned long password1,unsigned long password2);

'/******************************************************************************
' *
' * ������   : LockFlash_9030
' * ������� : 1023

' * ����   : 9030���� �û�flash ����
' * ����   :
    
'      Board_NO : 0��3, ���
'password1:       ������ , ��������
'password2:       ������ , ��������
       

' * ����ֵ :  1��0��-1,0=���ɹ�,1=�ɹ���-1=����ʧ��
' *
' *
' *****************************************************************************/
'short APIENTRY LockFlash_9030(unsigned short Board_NO,unsigned long password1,unsigned long password2);

'/******************************************************************************
' *
' * ������   : WriteFlash_9030
' * ������� : 128

' * ����   : д 9030���� �û�flash
' * ����   :
    
'      Board_NO : 0��3, ���
'      offset    : 0-199 ƫ��ֵ
'      len       : 1-8
'word1:       ������ , ��������
'word2:       ������ , ��������
       

' * ����ֵ :  1��0��-1,0=���ɹ�,1=�ɹ���-1=����ʧ��
' *
' *
' *****************************************************************************/
'short APIENTRY WriteFlash_9030(unsigned short Board_NO,unsigned short offset,unsigned short len,unsigned long word1,unsigned long word2);

'/******************************************************************************
' *
' * ������   : WriteFlashChar_9030
' * ������� : 128

' * ����   : д 9030���� �û�flash
' * ����   :
    
'      Board_NO : 0��3, ���
'      offset    : 0-199 ƫ��ֵ
'      len       : 1-8
'Data:             ����ָ��
       

' * ����ֵ :  1��0��-1,0=���ɹ�,1=�ɹ���-1=����ʧ��
' *
' *
' *****************************************************************************/
'short APIENTRY WriteFlashChar_9030(unsigned short Board_NO,unsigned short offset,unsigned short len,unsigned char *Data);

'/******************************************************************************
' *
' * ������   : UpdateFlash_9030
' * ������� : 129

' * ����   : ���� 9030���� �û�flash ����
' * ����   :
    
'      Board_NO : 0��3, ���
     
       

' * ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' *
' *
' *****************************************************************************/
'short APIENTRY UpDateFlash_9030(unsigned short Board_NO);


'/******************************************************************************
' *
' * ������   : ReadFlash_9030
' * ������� : 128

' * ����   :  �� 9030���� �û�flash
' * ����   :
    
'      Board_NO : 0��3, ���
'      offset    : 0-199 ƫ��ֵ
'      len       : 1-4
       

' * ����ֵ :
' *
' *
' *****************************************************************************/
'unsigned long APIENTRY ReadFlash_9030(unsigned short Board_NO,unsigned short offset,unsigned short len);

'/******************************************************************************
' *
' * ������   : ReadFlashChar_9030
' * ������� : 128

' * ����   :  �� 9030���� �û�flash
' * ����   :
    
'      Board_NO : 0��3, ���
'      offset    : 0-199 ƫ��ֵ
'      len       : 1-4
'Data:             ����ָ��
       

' * ����ֵ :  0��-1,-1=���ɹ�,0=�ɹ�
' *
' *
' *****************************************************************************/
'short APIENTRY ReadFlashChar_9030(unsigned short Board_NO,unsigned short offset,unsigned short len,unsigned char *Data);


