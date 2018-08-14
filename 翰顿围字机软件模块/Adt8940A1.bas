Attribute VB_Name = "adt8940a1"
Option Explicit
'******************************�����⺯��****************************
Declare Function adt8940a1_initial Lib "8940A1.dll" () As Integer
' ���ܣ���ʼ����
'����ֵ>0ʱ����ʾ8940A1�������������Ϊ3��������Ŀ��ÿ��ŷֱ�Ϊ0��1��2as integer
'����ֵ=0ʱ��˵��û�а�װ8940A1��as integer
'����ֵ<0ʱ��-1��ʾû�а�װ�˿���������-2��ʾPCI�Ŵ��ڹ��ϡ�

Declare Function get_lib_version Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'���ܣ���ȡ��ǰ��汾

Declare Function set_pulse_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Integer, ByVal logic As Long, ByVal dir_logic As Long) As Integer
'���ܣ������������Ĺ�����ʽ
'cardno ����
'axis ���(1 - 4)
'value       0������+���巽ʽ        1������+����ʽ
'logic       0: ���߼�����           1: ���߼�����
'dir-logic   0����������ź����߼�    1����������źŸ��߼�
'����ֵ      0: ��ȷ 1: ����
'Ĭ��ģʽ������+�������߼����壬��������ź����߼�

Declare Function set_limit_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal v1 As Integer, ByVal v2 As Integer, ByVal dir_logic As Integer) As Integer
'���ܣ��趨����������λ����nLMT�źŵ�ģʽ
'����:
'cardno ����
'axis ���(1 - 4)
'v1 0: ����λ��Ч 1: ����λ��Ч
'v2 0: ����λ��Ч 1: ����λ��Ч
'logic 0: �͵�ƽ��Ч 1: �ߵ�ƽ��Ч
'����ֵ 0: ��ȷ 1: ����
'Ĭ��ģʽΪ������λ��Ч������λ��Ч���͵�ƽ��Ч

Declare Function set_stop0_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal v As Integer, ByVal logic As Long) As Integer
'���ܣ��趨stop0�����źŵ�ģʽ
'cardno ����
'axis   ���(1 - 4)
'v      0: ��Ч       1: ��Ч
'logic  0: �͵�ƽ��Ч 1: �ߵ�ƽ��Ч
'����ֵ 0: ��ȷ       1: ����
'Ĭ��ģʽΪ: ��Ч

Declare Function set_stop1_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Long, ByVal v As Long, ByVal logic As Long) As Integer
'���ܣ��趨stop1�����źŵ�ģʽ
'cardno     ����
'axis       ���(1 - 4)
'v          0: ��Ч 1: ��Ч
'logic      0: �͵�ƽ��Ч 1: �ߵ�ƽ��Ч
'����ֵ      0: ��ȷ 1: ����
'Ĭ��ģʽΪ: ��Ч

Declare Function get_status Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'����: ��ȡ���������״̬
'cardno     ����
'axis       ���(1 - 4)
'v ����״ָ̬��
'           0:  �������� ��0: ��������
'����ֵ     0: ��ȷ 1: ����

Declare Function get_inp_status Lib "8940A1.dll" (ByVal cardno As Integer, ByRef value As Long) As Integer
'����: ��ȡ�岹������״̬
'cardno     ����
'v �岹״ָ̬��
'           0: �岹���� 1: ���ڲ岹
'����ֵ     0: ��ȷ     1: ����

Declare Function set_acc Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal add As Long) As Integer
'����: ���ٶ��趨
'cardno     ����
'axis       ���(1 - 4)
'Add        ��Χ(1 - 64000)
'���ٶ�ʵ��ֵ  add*125
'����ֵ     0: ��ȷ     1: ����

Declare Function set_startv Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal startv As Long) As Integer
'����: ��ʼ�ٶ��趨
'cardno     ����
'axis       ���(1 - 4)
'startv      ��Χ(1-2M)
'����ֵ     0: ��ȷ 1: ����

Declare Function set_speed Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal speed As Long) As Integer
'����: �����ٶ��趨
'cardno     ����
'axis       ���(1 - 4)
'speed      ��Χ(1-2M)
'����ֵ      0: ��ȷ 1: ����


Declare Function set_command_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Long) As Integer
'����: �߼�λ���趨
'cardno     ����
'axis       ���(1 - 4)
'value      ��Χ(-2147483648��+2147483647)
'����ֵ     0: ��ȷ 1: ����

Declare Function set_actual_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Long) As Integer
'����: ʵ��λ���趨
'cardno     ����
'axis       ���(1 - 4)
'value      ��Χ(-2147483648��+2147483647)
'����ֵ     0: ��ȷ 1: ����

Declare Function get_command_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'����: ��ȡ������߼�λ��
'cardno     ����
'axis       ���(1 - 4)
'value      �߼�λ�õ�ָ��
'����ֵ     0: ��ȷ 1: ����

Declare Function get_actual_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'����: ��ȡ�����ʵ��λ��
'cardno     ����
'axis       ���(1 - 4)
'value      ʵ��λ�õ�ָ��
'����ֵ     0: ��ȷ 1: ����

Declare Function get_speed Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'����: ��ȡ����ĵ�ǰ�����ٶ�
'cardno     ����
'axis       ���(1 - 4)
'value      ��ǰ�����ٶȵ�ָ��
'����ֵ     0: ��ȷ 1: ����

Declare Function get_out Lib "8940A1.dll" (ByVal cardno As Integer, ByVal number As Integer) As Integer
'*****************************************************
'����: ��ȡ�����
'����:
'    cardno ����
'    number �����
'����ֵ      ��ȡ����˿ڵĵ�ǰ״̬,0: �͵�ƽ   1: �ߵ�ƽ  -1:����
'*****************************************************/

Declare Function pmove Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Long) As Integer
'����: ��������
'cardno     ����
'axis       ���(1 - 4)
'value      �����������(-268435455��+268435455)
'           >0������������      <0������������
'����ֵ     0: ��ȷ     1: ����

Declare Function dec_stop Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'����: ��������ֹͣ
'cardno     ����
'axis       ���(1 - 4)
'����ֵ     0: ��ȷ 1: ����

Declare Function sudden_stop Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'����: ��������ֹͣ
'cardno     ����
'axis       ���(1 - 4)
'����ֵ     0: ��ȷ 1: ����

Declare Function inp_move2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long) As Long
'����: ����ֱ�߲岹
'cardno         ����
'axis1,axis2    ����岹�����
'pulse1,pulse2  �ƶ�����Ծ���(-8388608��+8388607)
'����ֵ         0: ��ȷ 1: ����

Declare Function inp_move3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long) As Long
'����: ����ֱ�߲岹
'cardno                 ����
'axis1,axis2,axis3      ����岹�����
'pulse1,pulse2,pulse3   �ƶ�����Ծ���(-8388608��+8388607)
'����ֵ                 0: ��ȷ 1: ����

Declare Function inp_move4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long) As Long
'����: ����ֱ�߲岹
'cardno ����
'pulse1,pulse2,pulse3,pulse4 XYZA�����ƶ�����Ծ���(-8388608��+8388607)
'����ֵ 0: ��ȷ 1: ����

Declare Function read_bit Lib "8940A1.dll" (ByVal cardno As Integer, ByVal number As Long) As Long
'����: ��ȡ�����
'cardno ����
'number �����(0 - 39)
'����ֵ 0: �͵�ƽ 1: �ߵ�ƽ -1: ����

Declare Function write_bit Lib "8940A1.dll" (ByVal cardno As Integer, ByVal number As Long, ByVal value As Long) As Long
'����: ���
'cardno ����
'number �����(0 - 15)
'value  0: �͵�ƽ   1: �ߵ�ƽ
'����ֵ  0: ��ȷ     1: ����

Declare Function get_hardware_ver Lib "8940A1.dll" (ByVal cardno As Integer) As Double
'����: ��ȡӲ���汾
'cardno     ����
'����ֵ     1: Ӳ����һ��         2:Ӳ���ڶ���
'�����1��2ֻ������ʱ��˵���ã�����ֵ�Ƕ��پ�Ϊ���٣�ĿǰӲ���汾Ϊ1.1

Declare Function set_suddenstop_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal v As Integer, ByVal logic As Integer) As Integer
'����: Ӳ��ֹͣģʽ����
'cardno     ����
'v          0: ��Ч 1: ��Ч
'logic      0: �͵�ƽ��Ч 1: �ߵ�ƽ��Ч
'����ֵ     0: ��ȷ 1: ����
'Ӳ��ֹͣ�źŹ̶�ʹ��P2���Ӱ�25���� (IN31)

Declare Function set_delay_time Lib "8940A1.dll" (ByVal cardno As Integer, ByVal time As Long) As Integer
'����: �趨��ʱʱ��
'cardno ����
'time   ��ʱʱ��
'����ֵ 0: ��ȷ 1: ����
'ʱ�䵥λΪ1/8us

Declare Function get_delay_status Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'����: ��ȡ��ʱ״̬
'cardno ����
'����ֵ  0: ��ʱ���� 1: ��ʱ������

'*********************************************//
'               ����������                     //
'*********************************************//
Declare Function set_symmetry_speed Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*******************************************************
'����:   �趨�ԳƼӼ��ٵ�ֵ
'����:
'    cardno ����
'    axis ���
'    lspd ���ٶ�
'    hspd �����ٶ�
'    tacc ����ʱ��
'����ֵ 0: ��ȷ 1: ����
'*******************************************************

Declare Function symmetry_relative_move Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'********************************************************
'*����:���յ�ǰλ��,�ԶԳƼӼ��ٽ��ж����ƶ�
'*����:
'      cardno -����
'      axis ---���
'      pulse --����
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'*********************************************************

Declare Function symmetry_absolute_move Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*********************************************************
'*����:�������λ��,�ԶԳƼӼ��ٽ��ж����ƶ�
'*����:
'      cardno -����
'      axis ---���
'      pulse --����
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'**********************************************************

Declare Function symmetry_relative_line2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'**********************************************************
'*����:���յ�ǰλ��,�ԶԳƼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
'      axis1 ---���1
'      axis2 ---���2
'      pulse1 --����1
'      pulse2 --����2
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'***********************************************************

Declare Function symmetry_absolute_line2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'***********************************************************
'*����:�������λ��,�ԶԳƼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
'      axis1 ---���1
'      axis2 ---���2
'      pulse1 --����1
'      pulse2 --����2
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'************************************************************/

Declare Function symmetry_relative_line3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'************************************************************
'*����:���յ�ǰλ��,�ԶԳƼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
'      axis1 ---���1
'      axis2 ---���2
'      axis3 ---���3
'      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'***************************************************************

Declare Function symmetry_absolute_line3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'**************************************************************
'����: �������λ�� , �ԶԳƼӼ��ٽ���ֱ�߲岹
'����:
'      cardno -����
'      axis1 ---���1
'      axis2 ---���2
'      axis3 ---���3
'      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'****************************************************************

Declare Function symmetry_relative_line4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*****************����ֱ�߲岹����˶�****************
'*����:���յ�ǰλ��,�ԼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
'      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      pulse4 --����4
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'******************************************************

Declare Function symmetry_absolute_line4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*****************����Գ�ֱ�߲岹�����˶�****************
'*����:�������λ��,�ԶԳƼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
'      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      pulse4 --����4
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'******************************************************


'//*********************************************//
'//               �ⲿ����                    //
'//*********************************************//

Declare Function manual_pmove Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal pos As Long) As Integer
'/************************�ⲿ�źŶ�����������**********************
'����: �ⲿ�źŶ�����������
'����:
'    cardno ����
'    axis ���(1 - 4)
'    pos ����
'����ֵ 0: ��ȷ 1: ����
'    ˵��:(1)�����������壬������û���������У���Ҫ�ȵ��ⲿ�źŵ�ƽ�����仯
'         (2)����ʹ����ͨ��ť,Ҳ���Խ�����
'******************************************************************/

Declare Function manual_continue Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'/************************�ⲿ�ź�������������**********************
'����: �ⲿ�ź�������������
'����:
'    cardno ����
'    axis ���(1 - 4)
'����ֵ 0: ��ȷ 1: ����
'    ˵��:(1)�����������壬������û���������У���Ҫ�ȵ��ⲿ�źŵ�ƽ�����仯
'         (2)����ʹ����ͨ��ť,Ҳ���Խ�����
'******************************************************************/

Declare Function manual_disable Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'/***********************�ر��ⲿ�ź�����ʹ��***********************
'����: �ر��ⲿ�ź�����ʹ��
'����:
'    cardno ����
'    axis ���(1 - 4)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/

'//*********************************************//
'//               λ������                    //
'//*********************************************//

Declare Function set_lock_position Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal mode As Integer, ByVal regi As Integer, ByVal logical As Integer) As Integer
'/****************************λ���������ú���**********************
'����: ���õ�λ�źŹ��� , ������������߼�λ�ú�ʵ��λ��
'����:
'    axis��������
'    mode��λ�����湤��ģʽ|0:��Ч
'                        |1:��Ч
'    regi��������ģʽ  |0:�߼�λ��
'                      |1:ʵ��λ��
'    logical����ƽ�ź� |0:�ɸߵ���
'                      |1:�ɵ͵���
'����ֵ 0: ��ȷ 1: ����
'˵��:    ʹ��ָ����axis��IN�ź���Ϊ�����ź�
'*******************************************************************/

Declare Function get_lock_status Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef v As Integer) As Integer
'/*************************��ȡ����״̬***********************
'����: ��ȡ����״̬
'����:
'    cardno ����
'    axis ���(1 - 4)
'    V            0|δִ��ͬ������
'                 1|ִ�й�ͬ������
'����ֵ 0: ��ȷ 1: ����
'˵��:    ���øú������Բ�׽λ�������Ƿ�ִ��
'******************************************************************/

Declare Function get_lock_position Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef pos As Long) As Integer
'/**************************��ȡ������λ��**************************
'����: ��ȡ������λ��
'����:
'    cardno ����
'    axis ���(1 - 4)
'    pos �����λ��
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/

Declare Function clr_lock_status Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'/**************************�������״̬**************************
'����: �������״̬
'����:
'    cardno ����
'    axis ���(1 - 4)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/

'//*********************************************//
'//               Ӳ������                    //
'//*********************************************//
Declare Function fifo_inp_move1 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal pulse1 As Long, ByVal speed As Long) As Integer
'/**************************���Ỻ��**************************
'����: ���Ỻ��
'����:
'    cardno ����
'    axis1 ���(1 - 4)
'    pulse1 ���������
'    speed ������ٶ�
'����ֵ 0: ��ȷ 1: ����
'˵��:����2048������ռ䣬ÿ�����Ỻ��ָ��ռ��3���ռ䣬�ɻ���682��ָ��
'******************************************************************/

Declare Function fifo_inp_move2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal speed As Long) As Integer
'/**************************���Ỻ��**************************
'����: ���Ỻ��
'����:
'    cardno ����
'    axis1 ���(1 - 4)
'    axis2 ���(1 - 4)
'    pulse1 �����������
'    pulse2 �����������
'    speed ������ٶ�
'����ֵ 0: ��ȷ 1: ����
'˵��:����2048������ռ䣬ÿ�����Ỻ��ָ��ռ��4���ռ䣬�ɻ���512��ָ��
'******************************************************************/

Declare Function fifo_inp_move3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal speed As Long) As Integer
'/**************************���Ỻ��**************************
'����: ���Ỻ��
'����:
'    cardno ����
'    axis1 ���(1 - 4)
'    axis2 ���(1 - 4)
'    axis3 ���(1 - 4)
'    pulse1 �����������
'    pulse2 �����������
'    pulse3 �����������
'    speed ������ٶ�
'����ֵ 0: ��ȷ 1: ����
'˵��:����2048������ռ䣬ÿ�����Ỻ��ָ��ռ��5���ռ䣬�ɻ���409��ָ��
'******************************************************************/

Declare Function fifo_inp_move4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal speed As Long) As Integer
'/**************************���Ỻ��**************************
'����: ���Ỻ��
'����:
'    cardno ����
'    axis1 ���(1 - 4)
'    axis2 ���(1 - 4)
'    axis3 ���(1 - 4)
'    axis4 ���(1 - 4)
'    pulse1 �����������
'    pulse2 �����������
'    pulse3 �����������
'    pulse4 �����������
'    speed ������ٶ�
'����ֵ 0: ��ȷ 1: ����
'˵��:����2048������ռ䣬ÿ�����Ỻ��ָ��ռ��6���ռ䣬�ɻ���341��ָ��
'******************************************************************/

Declare Function reset_fifo Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'/**************************���軺��**************************
'����: �������
'����:
'    cardno ����
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/

Declare Function read_fifo_count Lib "8940A1.dll" (ByVal cardno As Integer, ByRef value As Integer) As Integer
'/**************************��ȡ������**********************
'����:��ȡ����������Ž�ȥ��ָ�ʣ������δִ��
'����:
'    cardno ����
'    value  δִ�е�ָ����ռ���ֽ���
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/

Declare Function read_fifo_empty Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'/**************************��ȡ����״̬**********************
'����: ��ȡ�����Ƿ�Ϊ��
'����:
'    cardno ����
'����ֵ 0: �ǿ� 1: ��
'******************************************************************/

Declare Function read_fifo_full Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'/**************************��ȡ����״̬**********************
'����:��ȡ�����Ƿ����ˣ�����֮�󽫲����ٴ�����
'����:
'    cardno ����
'����ֵ 0: δ�� 1: ��
'******************************************************************/

Declare Function home1 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal backDir As Integer, ByVal logical0 As Integer, ByVal logical1 As Integer, ByVal homeStartV As Long, ByVal homeSpeed As Long, ByVal homeAcc As Long, ByVal searchRange As Long, ByVal searchSpeed As Long, ByVal phaseSpeed As Long, ByVal pulseUnit As Long) As Integer
'**************************�����ԭ��**********************
'����: ִ�е����ԭ���˶�
'����:
'    cardno ����
'    axis ���(1 - 4)
'    backDir                         ��ԭ�㷽��  0������    1������
'    logical0                        ��ԭ��stop0����  0:�͵�ƽ��Ч 1:�ߵ�ƽ��Ч
'    logical1                        ��ԭ��stop1����  0:�͵�ƽ��Ч 1:�ߵ�ƽ��Ч   -1����Ч��������Z�ࣩ
'    homeStartV                      ��ԭ����ʼ�ٶȣ�ȡֵ��Χ��0-2M
'    homeSpeed                       ��ԭ�������ٶȣ�ȡֵ��Χ��0-2M
'    homeAcc                         ��ԭ����ٶȣ�ȡֵ��Χ��0-64000
'    searchRange ԭ�㷶Χ(���˹���)
'    searchSpeed stop0�����ٶ�(���˹���)
'    phaseSpeed Z�������ٶ�(���˹���)
'    pulseUnit ÿת����
'
'����ֵ  0:��ԭ��ɹ�;   -1:��������;    -2����ԭ��ʧ��,(������λ��ԭ�㷶Χ��С);     1����ԭ�㱻��ֹ
'˵��:
' (1) ��ԭ���Ϊ�Ĵ�:
'     ��һ��:���ٽӽ�stop0(logical0ԭ������)���ҵ�stop0;
'     �ڶ���:���ٷ����뿪stop0�������ƶ�ָ��ԭ�㷶Χ������;
'     ������:�ٴ����ٽӽ�stop0;
'     ���Ĳ�:���ٽӽ�stop1(logical1������Z��).
' (2) ���Ĳ�����ѡ���Ƿ�ִ��,ͨ��logical1��ѡ��.
' (3) ��������ԭ��,����ȴ���һ���ԭ������󣬲���ִ����һ��Ļ�ԭ�㶯��.
'*****************************************************

Declare Function inp_arc Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal cood As Long) As Integer
'*************************����:����Բ���岹**************************
'���ܣ�     ��������Բ���岹�˶� ��������������岹ָ���װ��ͨ����ͨ�岹ʵ��
'����:
'    cardno ����
'    axis1 axis2 ���(1 - 4)
'    dir                 ��Բ����    0:˳ʱ��Բ ;1����ʱ��Բ
'    cood[]              Բ�������������(���,�м��,�յ�)��������Ԫ��
'
'  ����ֵ��  -3:���㲻�ܹ���Բ���� -2:��λ�ź�ֹͣ��-1:��������    0:�ɹ���  1:Բ���岹��ֹ.
'  ע�⣺Ĭ�ϲ���Բ���岹�����������嵱����ͬ;
'  ����岹�켣Ϊ��Բ���м�������ó���������Բ�ĶԳƵĵ�.
'********************************************************************
Declare Function fifo_arc Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal speed As Integer, ByVal ccood As Long) As Integer
'*************************����:����Բ���岹����ʵ��**************************
'���ܣ�     ��������Բ���岹�˶�����������Ӳ������岹ָ���װ��ͨ������ʵ�֡�
'����:
'    cardno ����
'    axis1 axis2 ���(1 - 4)
'    speed �岹�ٶ�
'    cood[]              Բ�������������(���,�м��,�յ�)��������Ԫ��
'
'����ֵ��  -3:���㲻�ܹ���Բ��;  -2-��λ�ź�ֹͣ;    -1:��������;    0:�ɹ�;     1:Բ���岹��ֹ.
'ע�⣺Ĭ�ϲ���Բ���岹�����������嵱����ͬ;
'      ����岹�켣Ϊ��Բ���м�������ó���������Բ�ĶԳƵĵ�.
'
'********************************************************************
Declare Function continue_move1 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal dir As Integer) As Integer
'*************************����:���������˶�**************************
'����:      ���������˶�
'����:
'    cardno ����
'    axis ���(1 - 4)
'    dir                 0:���� ;1������
'
'����ֵ��   -1:��λ�ź�ֹͣ; 1:����;     0:��ȷ.
'ע��:д����������ǰ,һ��Ҫ��ȷ���趨�ٶȲ���.
'********************************************************************

Declare Function continue_move2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal dir1 As Integer, ByVal dir2 As Integer) As Integer
'*************************����:���������˶�**************************
'����:      ���������˶�
'����:
'    cardno ����
'    axis1 ���(1 - 4)
'    axis2 ���(1 - 4)
'    dir1                0:����; 1������
'    dir2                0:����; 1������
'
'����ֵ��   -1:��λ�ź�ֹͣ; 1:����;    0:��ȷ.
'ע��:д����������ǰ,һ��Ҫ��ȷ���趨�ٶȲ���.
'********************************************************************


Public Sub MyProc()

    DoEvents

End Sub

