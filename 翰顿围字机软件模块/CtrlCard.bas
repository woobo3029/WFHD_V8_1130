Attribute VB_Name = "CtrlCard"
'********************** �˶�����ģ�� ********************

    'Ϊ�˼򵥡����㡢��ݵؿ�����ͨ���Ժá�����չ��ǿ��
    
    'ά�������Ӧ��ϵͳ�������ڿ��ƿ�������Ļ����Ͻ�
    
    '���п⺯�������˷����װ�������ʾ��ʹ��һ���˶�
    
    '���ƿ�

'********************************************************

''������ƿ�����
'Public Const CtrlCardType = 0     ' 0 ����adt8940a, 1 ���� 9030��
''�ı�忨���ͣ���Ҫ�ı���Ӧ����Ŷ���
'Public Const FeedAxis = 1
'Public Const BendAxis = 2
'Public Const VertAxis = 3
'Public Const VertUpDownAxis = 4

Public Const CtrlCardType = 4       '0=adt8940a, 1=9030�� 2=6052, 4=GALIL

'Public Const CtrlCardType = 4       '0=adt8940a, 1=9030�� 2=6052, 4=GALIL
'�ı�忨���ͣ���Ҫ�ı���Ӧ����Ŷ���
Public Const FeedAxis = 0
Public Const BendAxis = 1
Public Const VertAxis = 2
Public Const VertUpDownAxis = 3

Public Result As Integer      '����ֵ

Public hDmc As Long

Const MAXAXIS = 4           '�������

'*******************��ʼ������************************

    '�ú����а����˿��ƿ���ʼ�����õĿ⺯�������ǵ���
    
    '���������Ļ��������Ա�����ʾ�����������ȵ���
    
    '����ֵ<=0��ʾ��ʼ��ʧ�ܣ�����ֵ>0��ʾ��ʼ���ɹ�

'*****************************************************
Public Function Init_Card() As Integer
       
If 0 = CtrlCardType Then
    Result = adt8940a1_initial           '����ʼ��
    
    If Result <= 0 Then
     
       Init_Card = Result
       
       Exit Function
       
    End If
    
    For I = 1 To MAXAXIS
       
       set_command_pos 0, I, 0         '�߼�λ�ü���������
       
       set_actual_pos 0, I, 0          'ʵλλ�ü���������
       
       set_startv 0, I, 1000            '���ó�ʼ�ٶ�
       
       set_speed 0, I, 2000             '���������ٶ�
       
       set_acc 0, I, 625               '���ü��ٶ�
     
    Next I
    Init_Card = Result
ElseIf CtrlCardType = 4 Then
    rc = DMCOpen(1, 0, hDmc)
    If rc = 0 Then
        SetFEdir hDmc, 1
        SetAxisOutMode hDmc, VertUpDownAxis, -2
    End If
    Init_Card = rc
Else
   Result = InitCard_9030(0, 1, 1, 1, 1, 0)
   If Result <> 0 Then
        Init_Card = Result
        Exit Function
    End If
    
    '�趨home�㡢��λ��
    SetAxisIO_9030 0, BendAxis, 2, 3, 1, 5          '�仡��λ����5
    SetAxisIO_9030 0, VertUpDownAxis, 2, 3, 1, 9    '������λ����9
    SetAxisIO_9030 0, VertAxis, 2, 3, 1, 6          'ϳ���Ƕȸ�λ����6
    SetAxisMotorOnOff_9030 0, FeedAxis, 1
    SetAxisMotorOnOff_9030 0, BendAxis, 1
    SetAxisMotorOnOff_9030 0, VertUpDownAxis, 1
    SetAxisMotorOnOff_9030 0, VertAxis, 1
    
    '�ı��᷽��
    SetAxisOutMode_9030 0, VertUpDownAxis, 0, 0, 1
    SetAxisOutMode_9030 0, VertAxis, 0, 0, 0
   Init_Card = Result
End If
    
       
End Function

'********************��ȡ�汾��Ϣ************************
'
'    �ú������ڻ�ȡ������汾
'
'    ����:     libver -��汾��
'
'*********************************************************
Public Function Get_Version(libver As Double, hardwarever As Double) As Integer

    Dim ver As Integer
    
    ver = get_lib_version(0)
    
    libver = (ver)
    
    hardwarever = get_hardware_ver(0)
    
End Function

'**********************�����ٶ�ģ��***********************

'   ���ݲ�����ֵ���ж������ٻ��ǼӼ���

'    ������ĳ�ʼ�ٶȡ������ٶȺͼ��ٶ�

'    ����:       axis -���

'               StartV -��ʼ�ٶ�

'               Speed -�����ٶ�

'               Add -���ٶ�
    
'    ����ֵ=0��ȷ������ֵ=1����

'*********************************************************
Public Function Setup_Speed(ByVal axis As Long, ByVal startv As Long, ByVal speed As Long, ByVal add As Long, ByVal tacc As Double) As Integer

        If (startv - speed >= 0) Then
        
            Result = set_startv(0, axis, startv)
        
            set_speed 0, axis, startv
            
'            set_symmetry_speed 0, axis, startv, startv, tacc
            
        Else
        
            Result = set_startv(0, axis, startv)
        
            set_speed 0, axis, speed
        
            set_acc 0, axis, add / 125
            
'          set_symmetry_speed 0, axis, startv, speed, tacc
            
        End If
       
End Function

'*********************������������**********************

    '�ú����������������˶����˶�
    
    '������axis-��ţ�pulse-�����������
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function Axis_Pmove(ByVal axis As Long, ByVal pulse As Long) As Integer
    
    Result = pmove(0, axis, pulse)
    
    Axis_Pmove = Result
    
End Function

'*******************��������岹����********************

    '�ú���������������������в岹�˶�
    
    '����:     axis1 , axis2 - ����岹�����
    
    '          pulse1,pulse2-��Ӧ������������
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function Interp_Move2(ByVal axis1 As Long, ByVal axis2 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long) As Integer

    Result = inp_move2(0, axis1, axis2, pulse1, pulse2)
    
    Interp_Move2 = Result
    
End Function

'*******************��������岹����********************

    '�ú���������������������в岹�˶�
    
    '����:     axis1 , axis2,axis3 - ����岹�����
    
    '          pulse1,pulse2,pulse3-��Ӧ������������
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************

Public Function Interp_Move3(ByVal axis1 As Long, ByVal axis2 As Long, ByVal axis3 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long) As Integer

    Result = inp_move3(0, axis1, axis2, axis3, pulse1, pulse2, pulse3)
    
    Interp_Move3 = Result
    
End Function


'*******************����岹����********************

    '�ú�����������XYZW������в岹�˶�
    
    '����: pulse1,pulse2,pulse3,pulse4-��Ӧ������������
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function Interp_Move4(ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long) As Integer
    
    Result = inp_move4(0, pulse1, pulse2, pulse3, pulse4)
    
    Interp_Move4 = Result
    
End Function

'*******************ֹͣ��������********************

    '�ú�������ֹͣ��������Ϊ����ֹͣ�ͼ���ֹͣ
    
    '����: axis-��ţ�mode: 0-����ֹͣ��1-����ֹͣ
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function StopRun(ByVal axis As Long, ByVal mode As Long) As Integer

    If mode = 0 Then
        
        Result = sudden_stop(0, axis)
        
    Else
    
        Result = dec_stop(0, axis)
    
    End If

End Function

'*******************����λ�ú���********************

    '�ú������������߼�λ�ú�ʵ��λ��
    
    '����: axis-���            pos-λ������ֵ
    
    '      mode
    
    '         0 - �����߼�λ��     1 - ����ʵ��λ��
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function Setup_Pos(ByVal axis As Long, ByVal pos As Long, ByVal mode As Long) As Integer

    If mode = 0 Then
    
        Result = set_command_pos(0, axis, pos)
        
    Else
    
        Result = set_actual_pos(0, axis, pos)
        
    End If
    
End Function

'*******************��ȡ�˶���Ϣ����********************

    '�ú������ڻ�ȡ�߼�λ�á�ʵ��λ�ú������ٶ�
    
    '����: axis-��ţ�logps-�߼�λ��
    
    '      actpos-ʵ��λ�ã�speed-�����ٶ�
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function Get_CurrentInf(ByVal axis As Long, LogPos As Long, actpos As Long, speed As Long) As Integer

    Result = get_command_pos(0, axis, LogPos)
    
    get_actual_pos 0, axis, actpos
    
    get_speed 0, axis, speed
    
    Get_CurrentInf = Result
    
End Function


'*******************��ȡ�˶�״̬����********************

    '�ú������ڻ�ȡ���������״̬�Ͳ岹������״̬
    
    '����: axis-��ţ�value-״̬(0-������������0-��������)
    
    '      mode 0-��ȡ���������״̬����0-��ȡ�岹������״̬
    
    '����ֵ=0��ȷ������ֵ=1����

'*******************************************************
Public Function Get_MoveStatus(ByVal axis As Long, value As Long, ByVal mode As Integer) As Integer

    If mode = 0 Then
    
        GetMove_Status = get_status(0, axis, value)
        
    Else
    
        GetMove_Status = get_inp_status(0, value)
        
    End If
    
End Function

'***********************��ȡ�����*******************************
'
'     �ú������ڶ�ȡ���������
'
'     ������number-�����(0 ~ 39)
'
'     ����ֵ��0 �� �͵�ƽ��1 �� �ߵ�ƽ��-1 �� ����
'
'****************************************************************
Public Function Read_Input(ByVal number As Long) As Integer
    
    Read_Input = read_bit(0, number)
    
End Function

'*********************������㺯��******************************
'
'    �ú���������������ź�
'
'    ������ number-�����(0 ~ 15)

'           value 0-�͵�ƽ       1���ߵ�ƽ
'
'    ����ֵ=0��ȷ������ֵ=1����
'****************************************************************
Public Function Write_Output(ByVal number As Long, ByVal value As Long) As Integer

    Write_Output = write_bit(0, number, value)
    
End Function


'********************�������������ʽ**********************
'
'    �ú���������������Ĺ�����ʽ
'
'    ������axis-��ţ� value-���巽ʽ 0�����士���巽ʽ 1�����士����ʽ
'
'    ����ֵ=0��ȷ������ֵ=1����
'
'    Ĭ�����巽ʽΪ����+����ʽ
'
'    ���������Ĭ�ϵ����߼�����ͷ�������ź����߼�
'
'*********************************************************
Public Function Setup_pulseMode(ByVal axis As Long, ByVal value As Long) As Integer

    Setup_pulseMode = set_pulse_mode(0, axis, value, 0, 0)
    
End Function

'********************������λ�źŷ�ʽ**********************
'
'   �ú��������趨��/��������λ����nLMT�źŵ�ģʽ
'
'   ����:      axis -���
'              value1   0������λ��Ч  1������λ��Ч
'              value2   0������λ��Ч  1������λ��Ч
'              logic    0���͵�ƽ��Ч  1���ߵ�ƽ��Ч
'   Ĭ��ģʽΪ:    ����λ��Ч,����λ��Ч,�͵�ƽ��Ч
'
'   ����ֵ=0��ȷ������ֵ=1����
'  *********************************************************
Public Function Setup_LimitMode(ByVal axis As Long, ByVal value1 As Long, ByVal value2 As Long, ByVal logic As Long) As Integer

    Setup_LimitMode = set_limit_mode(0, axis, value1, value2, logic)
    
End Function

'
'********************����stop0�źŷ�ʽ**********************
'
'   �ú��������趨stop0�źŵ�ģʽ
'
'   ����:     axis -���

'             value   0����Ч        1����Ч

'             logic   0���͵�ƽ��Ч  1���ߵ�ƽ��Ч
'   Ĭ��ģʽΪ:    ��Ч
'
'   ����ֵ=0��ȷ������ֵ=1����
'  *********************************************************
Public Function Setup_Stop0Mode(ByVal axis As Long, ByVal value As Long, ByVal logic As Long) As Integer

    Setup_Stop0Mode = set_stop0_mode(0, axis, value, logic)
    
End Function


'********************����stop1�źŷ�ʽ**********************
'
'   �ú��������趨stop1�źŵ�ģʽ
'
'   ����:     axis -���
'             value   0����Ч       1����Ч

'             logic   0���͵�ƽ��Ч  1���ߵ�ƽ��Ч
'   Ĭ��ģʽΪ:    ��Ч
'
'   ����ֵ=0��ȷ������ֵ=1����
'  *********************************************************
Public Function Setup_Stop1Mode(ByVal axis As Long, ByVal value As Long, ByVal logic As Long) As Integer

    Setup_Stop1Mode = set_stop1_mode(0, axis, value, logic)
    
End Function

'********************����Ӳ��ֹͣ**************************
'
'   �ú��������趨Ӳ��ֹͣ��ģʽ
'
'   ����:     value   0����Ч        1����Ч

'             logic   0���͵�ƽ��Ч  1���ߵ�ƽ��Ч

'   Ĭ��ģʽΪ:    ��Ч
'
'   ����ֵ=0��ȷ������ֵ=1����

'   Ӳ��ֹͣ�źŹ̶�ʹ��P3���Ӱ�34����(IN31)
'  *********************************************************

Public Function Setup_HardStop(ByVal value As Long, ByVal logic As Long) As Integer

    Setup_HardStop = set_suddenstop_mode(0, value, logic)
    
End Function

'********************������ʱ**************************
'
'   �ú��������趨��ʱ
'
'   ����:     time - ��ʱʱ�䣨��λΪus��
'
'   ����ֵ=0��ȷ������ֵ=1����

'  *********************************************************

Public Function Setup_Delay(ByVal Time As Long) As Integer

    Setup_Delay = set_delay_time(0, Time * 8)
    
End Function

'**********************��ȡ��ʱ״̬**********************

'   �ú������ڻ�ȡ��ʱ��״̬

'   ����ֵ    0 - ��ʱ����    1 - ��ʱ������

'********************************************************

Public Function Get_DelayStatus() As Integer

    Get_DelayStatus = get_delay_status(0)
    
End Function

'------------------------����������--------------------------
'˵��:���º�����Ϊ�˷���ͻ���ʹ�ö����ӵĺ���
'-----------------------------------------------------------

'*****************************��������˶�*********************
'����:���յ�ǰλ��,�ԼӼ��ٽ��ж����ƶ�
'����:
'      cardno -����
'      axis ---���
'      pulse --����
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'*******************************************************************/
Public Function Sym_RelativeMove(ByVal axis As Long, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_move(0, axis, pulse, lspd, hspd, tacc)

    Symmetry_RelativeMove = Result
End Function
'/***************************��������ƶ�************************
'*����:�������λ��,�ԼӼ��ٽ��ж����ƶ�
'*����:
'      cardno -����
'      axis ---���
'      pulse --����
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'********************************************************************/
Public Function Sym_AbsoluteMove(ByVal axis As Integer, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
    
    Result = symmetry_absolute_move(0, axis, pulse, lspd, hspd, tacc)
    
    Symmetry_AbsoluteMove = Result
    
End Function

'**********************����ֱ�߲岹����ƶ�********************
'*����:���յ�ǰλ��,�ԼӼ��ٽ���ֱ�߲岹
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
'******************************************************************/
Public Function Sym_RelativeLine2(ByVal axis1 As Long, ByVal axis2 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_line2(0, axis1, axis2, pulse1, pulse2, lspd, hspd, tacc)

    Symmetry_RelativeLine2 = Result

End Function
'********************����ֱ�߲岹�����ƶ�**********************
'*����:�������λ��,�ԼӼ��ٽ���ֱ�߲岹
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
'******************************************************************/
Public Function Sym_AbsoluteLine2(ByVal axis1 As Long, ByVal axis2 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
    
    Result = symmetry_absolute_line2(0, axis1, axis2, pulse1, pulse2, lspd, hspd, tacc)
    
    Symmetry_AbsoluteLine2 = Result

End Function

'**********************����ֱ�߲岹����˶�********************
'*����:���յ�ǰλ��,�ԼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
'      axis1 ---���1
'      axis2 ---���2
'      axis3 ---���3
''      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/
Public Function Sym_RelativeLine3(ByVal axis1 As Long, ByVal axis2 As Long, ByVal axis3 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_line3(0, axis1, axis2, axis3, pulse1, pulse2, pulse3, lspd, hspd, tacc)

    Symmetry_RelativeLine3 = Result

End Function
'*********************����ֱ�߲岹�����˶�*********************
'����: �������λ�� , �ԼӼ��ٽ���ֱ�߲岹
'����:
'      cardno -����
''      axis1 ---���1
'      axis2 ---���2
'      axis3 ---���3
'      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/
Public Function Sym_AbsoluteLine3(ByVal axis1 As Long, ByVal axis2 As Long, ByVal axis3 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_absolute_line3(0, axis1, axis2, axis3, pulse1, pulse2, pulse3, lspd, hspd, tacc)

    Symmetry_AbsoluteLine3 = Result

End Function


'**********************����ֱ�߲岹����˶�********************
'*����:���յ�ǰλ��,�ԼӼ��ٽ���ֱ�߲岹
'*����:
'      cardno -����
''      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      pulse4 --����4
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/
Public Function Sym_RelativeLine4(ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_line4(0, pulse1, pulse2, pulse3, pulse4, lspd, hspd, tacc)

    Symmetry_RelativeLine4 = Result

End Function
'*********************����ֱ�߲岹�����˶�*********************
'����: �������λ�� , �ԼӼ��ٽ���ֱ�߲岹
'����:
'      cardno -����
'      pulse1 --����1
'      pulse2 --����2
'      pulse3 --����3
'      pulse4 --����4
'      lspd ---����
'      hspd ---����
'      tacc---����ʱ��(��λ:��)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/
Public Function Sym_AbsoluteLine4(ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_absolute_line4(0, pulse1, pulse2, pulse3, pulse4, lspd, hspd, tacc)

    Symmetry_AbsoluteLine4 = Result

End Function


'------------------------�ⲿ�ź�����--------------------------
'˵��:�ⲿ�źſ��������ֻ�ͨ�������ź�
'-----------------------------------------------------------
'********************�ⲿ�źŶ�������***********************************************
'����: �ⲿ�źŶ�����������
'����:
'    axis ���
'    pulse ����
'����ֵ 0: ��ȷ 1: ����
'    ˵��:(1)�����������壬������û���������У���Ҫ�ȵ��ⲿ�źŵ�ƽ�����仯
'         (2)����ʹ����ͨ��ť,Ҳ���Խ�����
'******************************************************************/
Public Function Manu_Pmove(ByVal axis As Long, ByVal pulse As Long) As Integer

    Result = manual_pmove(0, axis, pulse)
    
    Manu_Pmove = Result
    
End Function

'************************�ⲿ�ź�������������**********************
'����: �ⲿ�ź�������������
'����:
'    axis ���
'����ֵ 0: ��ȷ 1: ����
'    ˵��:(1)�����������壬������û���������У���Ҫ�ȵ��ⲿ�źŵ�ƽ�����仯
'         (2)����ʹ����ͨ��ť,Ҳ���Խ�����
'******************************************************************/
Public Function Manu_Continue(ByVal axis As Long) As Integer

    Result = manual_continue(0, axis)
    
    Manu_Continue = Result

End Function

'***********************�ر��ⲿ�ź�����ʹ��***********************
'����: �ر��ⲿ�ź�����ʹ��
'����:
'    axis ���
'����ֵ 0: ��ȷ 1: ����
'******************************************************************/
Public Function Disable_Manu(ByVal axis As Long) As Integer

   Result = manual_disable(0, axis)

   Disable_Manu = Result

End Function

'------------------------λ�����湦��--------------------------
'˵��:�������źű���������������ǰλ�þ����������񡣸ù�������λ�ò���ʮ��׼ȷ�����㡣
'-----------------------------------------------------------
'*************************��ȡ����״̬***********************
'����: ��ȡ����״̬
'����:
'    axis ���
'    status��0|δִ������״̬
'            1|ִ�й�����״̬
'����ֵ 0: ��ȷ 1: ����
'˵��:    ���øú������Բ�׽λ�������Ƿ�ִ��
'******************************************************************/
Public Function Get_LockStatus(ByVal axis As Long, Status As Long) As Integer
    Dim istatus As Integer

    Result = get_lock_status(0, axis, istatus)
 
    Status = istatus
    Get_LockStatus = Result
    
End Function

'****************************λ���������ú���**********************
'����: ���õ�λ�źŹ��� , ������������߼�λ�ú�ʵ��λ��
'����:
'    axis��������
'    mode��λ�����湤��ģʽ|0:��Ч
'                         |1:��Ч
'    regi��������ģʽ  |0:�߼�λ��
'                      |1:ʵ��λ��
'    logical����ƽ�ź� |0:�ɸߵ���
'                      |1:�ɵ͵���
'����ֵ 0: ��ȷ 1: ����
'˵��:    ʹ��ָ����axis��IN�ź���Ϊ�����ź�
'*******************************************************************/
Public Function Setup_LockPosition(ByVal axis As Long, ByVal mode As Long, ByVal regi As Long, ByVal logical As Long) As Integer
    
    Result = set_lock_position(0, axis, mode, regi, logical)
    
    Setup_LockPosition = Result
    
End Function


'**************************��ȡ������λ��**************************
'����: ��ȡ������λ��
'����:
'    axis ���
'    pos �����λ��
'����ֵ 0: ��ȷ 1: ����
'******************************************************************
Public Function Get_LockPosition(ByVal axis As Long, pos As Long) As Integer

    Result = get_lock_position(0, axis, pos)
    
    Get_LockPosition = Result
    
End Function

'**************************�������״̬**************************
'����: �������״̬
'����:
'    axis ���(1 - 4)
'����ֵ 0: ��ȷ 1: ����
'******************************************************************
Public Function Clr_LockStatus(ByVal axis As Long) As Integer

    Result = clr_lock_status(0, axis)
    
    Clr_LockStatus = Result
    
End Function


