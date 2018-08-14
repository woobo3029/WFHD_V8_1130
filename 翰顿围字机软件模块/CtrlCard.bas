Attribute VB_Name = "CtrlCard"
'********************** 运动控制模块 ********************

    '为了简单、方便、快捷地开发出通用性好、可扩展性强、
    
    '维护方便的应用系统，我们在控制卡函数库的基础上将
    
    '所有库函数进行了分类封装。下面的示例使用一块运动
    
    '控制卡

'********************************************************

''定义控制卡类型
'Public Const CtrlCardType = 0     ' 0 代表adt8940a, 1 代表 9030，
''改变板卡类型，需要改变相应的轴号定义
'Public Const FeedAxis = 1
'Public Const BendAxis = 2
'Public Const VertAxis = 3
'Public Const VertUpDownAxis = 4

Public Const CtrlCardType = 4       '0=adt8940a, 1=9030， 2=6052, 4=GALIL

'Public Const CtrlCardType = 4       '0=adt8940a, 1=9030， 2=6052, 4=GALIL
'改变板卡类型，需要改变相应的轴号定义
Public Const FeedAxis = 0
Public Const BendAxis = 1
Public Const VertAxis = 2
Public Const VertUpDownAxis = 3

Public Result As Integer      '返回值

Public hDmc As Long

Const MAXAXIS = 4           '最大轴数

'*******************初始化函数************************

    '该函数中包含了控制卡初始化常用的库函数，这是调用
    
    '其他函数的基础，所以必须在示例程序中最先调用
    
    '返回值<=0表示初始化失败，返回值>0表示初始化成功

'*****************************************************
Public Function Init_Card() As Integer
       
If 0 = CtrlCardType Then
    Result = adt8940a1_initial           '卡初始化
    
    If Result <= 0 Then
     
       Init_Card = Result
       
       Exit Function
       
    End If
    
    For I = 1 To MAXAXIS
       
       set_command_pos 0, I, 0         '逻辑位置计数器清零
       
       set_actual_pos 0, I, 0          '实位位置计数器清零
       
       set_startv 0, I, 1000            '设置初始速度
       
       set_speed 0, I, 2000             '设置驱动速度
       
       set_acc 0, I, 625               '设置加速度
     
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
    
    '设定home点、限位点
    SetAxisIO_9030 0, BendAxis, 2, 3, 1, 5          '弯弧复位开关5
    SetAxisIO_9030 0, VertUpDownAxis, 2, 3, 1, 9    '升降复位开关9
    SetAxisIO_9030 0, VertAxis, 2, 3, 1, 6          '铣刀角度复位开关6
    SetAxisMotorOnOff_9030 0, FeedAxis, 1
    SetAxisMotorOnOff_9030 0, BendAxis, 1
    SetAxisMotorOnOff_9030 0, VertUpDownAxis, 1
    SetAxisMotorOnOff_9030 0, VertAxis, 1
    
    '改变轴方向
    SetAxisOutMode_9030 0, VertUpDownAxis, 0, 0, 1
    SetAxisOutMode_9030 0, VertAxis, 0, 0, 0
   Init_Card = Result
End If
    
       
End Function

'********************获取版本信息************************
'
'    该函数用于获取函数库版本
'
'    参数:     libver -库版本号
'
'*********************************************************
Public Function Get_Version(libver As Double, hardwarever As Double) As Integer

    Dim ver As Integer
    
    ver = get_lib_version(0)
    
    libver = (ver)
    
    hardwarever = get_hardware_ver(0)
    
End Function

'**********************设置速度模块***********************

'   依据参数的值，判断是匀速还是加减速

'    设置轴的初始速度、驱动速度和加速度

'    参数:       axis -轴号

'               StartV -初始速度

'               Speed -驱动速度

'               Add -加速度
    
'    返回值=0正确，返回值=1错误

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

'*********************单轴驱动函数**********************

    '该函数用于驱动单个运动轴运动
    
    '参数：axis-轴号，pulse-输出的脉冲数
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function Axis_Pmove(ByVal axis As Long, ByVal pulse As Long) As Integer
    
    Result = pmove(0, axis, pulse)
    
    Axis_Pmove = Result
    
End Function

'*******************任意两轴插补函数********************

    '该函数用于驱动任意两轴进行插补运动
    
    '参数:     axis1 , axis2 - 参与插补的轴号
    
    '          pulse1,pulse2-对应轴的输出脉冲数
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function Interp_Move2(ByVal axis1 As Long, ByVal axis2 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long) As Integer

    Result = inp_move2(0, axis1, axis2, pulse1, pulse2)
    
    Interp_Move2 = Result
    
End Function

'*******************任意三轴插补函数********************

    '该函数用于驱动任意三轴进行插补运动
    
    '参数:     axis1 , axis2,axis3 - 参与插补的轴号
    
    '          pulse1,pulse2,pulse3-对应轴的输出脉冲数
    
    '返回值=0正确，返回值=1错误

'*******************************************************

Public Function Interp_Move3(ByVal axis1 As Long, ByVal axis2 As Long, ByVal axis3 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long) As Integer

    Result = inp_move3(0, axis1, axis2, axis3, pulse1, pulse2, pulse3)
    
    Interp_Move3 = Result
    
End Function


'*******************四轴插补函数********************

    '该函数用于驱动XYZW四轴进行插补运动
    
    '参数: pulse1,pulse2,pulse3,pulse4-对应轴的输出脉冲数
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function Interp_Move4(ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long) As Integer
    
    Result = inp_move4(0, pulse1, pulse2, pulse3, pulse4)
    
    Interp_Move4 = Result
    
End Function

'*******************停止驱动函数********************

    '该函数用于停止驱动，分为立即停止和减速停止
    
    '参数: axis-轴号，mode: 0-立即停止，1-减速停止
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function StopRun(ByVal axis As Long, ByVal mode As Long) As Integer

    If mode = 0 Then
        
        Result = sudden_stop(0, axis)
        
    Else
    
        Result = dec_stop(0, axis)
    
    End If

End Function

'*******************设置位置函数********************

    '该函数用于设置逻辑位置和实际位置
    
    '参数: axis-轴号            pos-位置设置值
    
    '      mode
    
    '         0 - 设置逻辑位置     1 - 设置实际位置
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function Setup_Pos(ByVal axis As Long, ByVal pos As Long, ByVal mode As Long) As Integer

    If mode = 0 Then
    
        Result = set_command_pos(0, axis, pos)
        
    Else
    
        Result = set_actual_pos(0, axis, pos)
        
    End If
    
End Function

'*******************获取运动信息函数********************

    '该函数用于获取逻辑位置、实际位置和运行速度
    
    '参数: axis-轴号，logps-逻辑位置
    
    '      actpos-实际位置，speed-运行速度
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function Get_CurrentInf(ByVal axis As Long, LogPos As Long, actpos As Long, speed As Long) As Integer

    Result = get_command_pos(0, axis, LogPos)
    
    get_actual_pos 0, axis, actpos
    
    get_speed 0, axis, speed
    
    Get_CurrentInf = Result
    
End Function


'*******************获取运动状态函数********************

    '该函数用于获取各轴的驱动状态和插补的驱动状态
    
    '参数: axis-轴号，value-状态(0-驱动结束，非0-正在驱动)
    
    '      mode 0-获取单轴的驱动状态，非0-获取插补的驱动状态
    
    '返回值=0正确，返回值=1错误

'*******************************************************
Public Function Get_MoveStatus(ByVal axis As Long, value As Long, ByVal mode As Integer) As Integer

    If mode = 0 Then
    
        GetMove_Status = get_status(0, axis, value)
        
    Else
    
        GetMove_Status = get_inp_status(0, value)
        
    End If
    
End Function

'***********************读取输入点*******************************
'
'     该函数用于读取单个输入点
'
'     参数：number-输入点(0 ~ 39)
'
'     返回值：0 － 低电平，1 － 高电平，-1 － 错误
'
'****************************************************************
Public Function Read_Input(ByVal number As Long) As Integer
    
    Read_Input = read_bit(0, number)
    
End Function

'*********************输出单点函数******************************
'
'    该函数用于输出单点信号
'
'    参数： number-输出点(0 ~ 15)

'           value 0-低电平       1－高电平
'
'    返回值=0正确，返回值=1错误
'****************************************************************
Public Function Write_Output(ByVal number As Long, ByVal value As Long) As Integer

    Write_Output = write_bit(0, number, value)
    
End Function


'********************设置脉冲输出方式**********************
'
'    该函数用于设置脉冲的工作方式
'
'    参数：axis-轴号， value-脉冲方式 0－脉冲＋脉冲方式 1－脉冲＋方向方式
'
'    返回值=0正确，返回值=1错误
'
'    默认脉冲方式为脉冲+方向方式
'
'    本程序采用默认的正逻辑脉冲和方向输出信号正逻辑
'
'*********************************************************
Public Function Setup_pulseMode(ByVal axis As Long, ByVal value As Long) As Integer

    Setup_pulseMode = set_pulse_mode(0, axis, value, 0, 0)
    
End Function

'********************设置限位信号方式**********************
'
'   该函数用于设定正/负方向限位输入nLMT信号的模式
'
'   参数:      axis -轴号
'              value1   0－正限位有效  1－正限位无效
'              value2   0－负限位有效  1－负限位无效
'              logic    0－低电平有效  1－高电平有效
'   默认模式为:    正限位有效,负限位有效,低电平有效
'
'   返回值=0正确，返回值=1错误
'  *********************************************************
Public Function Setup_LimitMode(ByVal axis As Long, ByVal value1 As Long, ByVal value2 As Long, ByVal logic As Long) As Integer

    Setup_LimitMode = set_limit_mode(0, axis, value1, value2, logic)
    
End Function

'
'********************设置stop0信号方式**********************
'
'   该函数用于设定stop0信号的模式
'
'   参数:     axis -轴号

'             value   0－无效        1－有效

'             logic   0－低电平有效  1－高电平有效
'   默认模式为:    无效
'
'   返回值=0正确，返回值=1错误
'  *********************************************************
Public Function Setup_Stop0Mode(ByVal axis As Long, ByVal value As Long, ByVal logic As Long) As Integer

    Setup_Stop0Mode = set_stop0_mode(0, axis, value, logic)
    
End Function


'********************设置stop1信号方式**********************
'
'   该函数用于设定stop1信号的模式
'
'   参数:     axis -轴号
'             value   0－无效       1－有效

'             logic   0－低电平有效  1－高电平有效
'   默认模式为:    无效
'
'   返回值=0正确，返回值=1错误
'  *********************************************************
Public Function Setup_Stop1Mode(ByVal axis As Long, ByVal value As Long, ByVal logic As Long) As Integer

    Setup_Stop1Mode = set_stop1_mode(0, axis, value, logic)
    
End Function

'********************设置硬件停止**************************
'
'   该函数用于设定硬件停止的模式
'
'   参数:     value   0－无效        1－有效

'             logic   0－低电平有效  1－高电平有效

'   默认模式为:    无效
'
'   返回值=0正确，返回值=1错误

'   硬件停止信号固定使用P3端子板34引脚(IN31)
'  *********************************************************

Public Function Setup_HardStop(ByVal value As Long, ByVal logic As Long) As Integer

    Setup_HardStop = set_suddenstop_mode(0, value, logic)
    
End Function

'********************设置延时**************************
'
'   该函数用于设定延时
'
'   参数:     time - 延时时间（单位为us）
'
'   返回值=0正确，返回值=1错误

'  *********************************************************

Public Function Setup_Delay(ByVal Time As Long) As Integer

    Setup_Delay = set_delay_time(0, Time * 8)
    
End Function

'**********************获取延时状态**********************

'   该函数用于获取延时的状态

'   返回值    0 - 延时结束    1 - 延时进行中

'********************************************************

Public Function Get_DelayStatus() As Integer

    Get_DelayStatus = get_delay_status(0)
    
End Function

'------------------------复合驱动类--------------------------
'说明:以下函数是为了方便客户的使用而增加的函数
'-----------------------------------------------------------

'*****************************单轴相对运动*********************
'功能:参照当前位置,以加减速进行定量移动
'参数:
'      cardno -卡号
'      axis ---轴号
'      pulse --脉冲
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'*******************************************************************/
Public Function Sym_RelativeMove(ByVal axis As Long, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_move(0, axis, pulse, lspd, hspd, tacc)

    Symmetry_RelativeMove = Result
End Function
'/***************************单轴绝对移动************************
'*功能:参照零点位置,以加减速进行定量移动
'*参数:
'      cardno -卡号
'      axis ---轴号
'      pulse --脉冲
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'********************************************************************/
Public Function Sym_AbsoluteMove(ByVal axis As Integer, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
    
    Result = symmetry_absolute_move(0, axis, pulse, lspd, hspd, tacc)
    
    Symmetry_AbsoluteMove = Result
    
End Function

'**********************两轴直线插补相对移动********************
'*功能:参照当前位置,以加减速进行直线插补
'*参数:
'      cardno -卡号
'      axis1 ---轴号1
'      axis2 ---轴号2
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Sym_RelativeLine2(ByVal axis1 As Long, ByVal axis2 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_line2(0, axis1, axis2, pulse1, pulse2, lspd, hspd, tacc)

    Symmetry_RelativeLine2 = Result

End Function
'********************两轴直线插补绝对移动**********************
'*功能:参照零点位置,以加减速进行直线插补
'*参数:
'      cardno -卡号
'      axis1 ---轴号1
'      axis2 ---轴号2
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Sym_AbsoluteLine2(ByVal axis1 As Long, ByVal axis2 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
    
    Result = symmetry_absolute_line2(0, axis1, axis2, pulse1, pulse2, lspd, hspd, tacc)
    
    Symmetry_AbsoluteLine2 = Result

End Function

'**********************三轴直线插补相对运动********************
'*功能:参照当前位置,以加减速进行直线插补
'*参数:
'      cardno -卡号
'      axis1 ---轴号1
'      axis2 ---轴号2
'      axis3 ---轴号3
''      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Sym_RelativeLine3(ByVal axis1 As Long, ByVal axis2 As Long, ByVal axis3 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_line3(0, axis1, axis2, axis3, pulse1, pulse2, pulse3, lspd, hspd, tacc)

    Symmetry_RelativeLine3 = Result

End Function
'*********************三轴直线插补绝对运动*********************
'功能: 参照零点位置 , 以加减速进行直线插补
'参数:
'      cardno -卡号
''      axis1 ---轴号1
'      axis2 ---轴号2
'      axis3 ---轴号3
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Sym_AbsoluteLine3(ByVal axis1 As Long, ByVal axis2 As Long, ByVal axis3 As Long, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_absolute_line3(0, axis1, axis2, axis3, pulse1, pulse2, pulse3, lspd, hspd, tacc)

    Symmetry_AbsoluteLine3 = Result

End Function


'**********************四轴直线插补相对运动********************
'*功能:参照当前位置,以加减速进行直线插补
'*参数:
'      cardno -卡号
''      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      pulse4 --脉冲4
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Sym_RelativeLine4(ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_relative_line4(0, pulse1, pulse2, pulse3, pulse4, lspd, hspd, tacc)

    Symmetry_RelativeLine4 = Result

End Function
'*********************四轴直线插补绝对运动*********************
'功能: 参照零点位置 , 以加减速进行直线插补
'参数:
'      cardno -卡号
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      pulse4 --脉冲4
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Sym_AbsoluteLine4(ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer

    Result = symmetry_absolute_line4(0, pulse1, pulse2, pulse3, pulse4, lspd, hspd, tacc)

    Symmetry_AbsoluteLine4 = Result

End Function


'------------------------外部信号驱动--------------------------
'说明:外部信号可以是手轮或通用输入信号
'-----------------------------------------------------------
'********************外部信号定量驱动***********************************************
'功能: 外部信号定量驱动函数
'参数:
'    axis 轴号
'    pulse 脉冲
'返回值 0: 正确 1: 错误
'    说明:(1)发出定量脉冲，但驱动没有立即进行，需要等到外部信号电平发生变化
'         (2)可以使用普通按钮,也可以接手轮
'******************************************************************/
Public Function Manu_Pmove(ByVal axis As Long, ByVal pulse As Long) As Integer

    Result = manual_pmove(0, axis, pulse)
    
    Manu_Pmove = Result
    
End Function

'************************外部信号连续驱动函数**********************
'功能: 外部信号连续驱动函数
'参数:
'    axis 轴号
'返回值 0: 正确 1: 错误
'    说明:(1)发出定量脉冲，但驱动没有立即进行，需要等到外部信号电平发生变化
'         (2)可以使用普通按钮,也可以接手轮
'******************************************************************/
Public Function Manu_Continue(ByVal axis As Long) As Integer

    Result = manual_continue(0, axis)
    
    Manu_Continue = Result

End Function

'***********************关闭外部信号驱动使能***********************
'功能: 关闭外部信号驱动使能
'参数:
'    axis 轴号
'返回值 0: 正确 1: 错误
'******************************************************************/
Public Function Disable_Manu(ByVal axis As Long) As Integer

   Result = manual_disable(0, axis)

   Disable_Manu = Result

End Function

'------------------------位置锁存功能--------------------------
'说明:当锁存信号被触发，编码器当前位置就立即被捕获。该功能用于位置测量十分准确、方便。
'-----------------------------------------------------------
'*************************获取锁存状态***********************
'功能: 获取锁存状态
'参数:
'    axis 轴号
'    status—0|未执行锁存状态
'            1|执行过锁存状态
'返回值 0: 正确 1: 错误
'说明:    利用该函数可以捕捉位置锁存是否执行
'******************************************************************/
Public Function Get_LockStatus(ByVal axis As Long, Status As Long) As Integer
    Dim istatus As Integer

    Result = get_lock_status(0, axis, istatus)
 
    Status = istatus
    Get_LockStatus = Result
    
End Function

'****************************位置锁存设置函数**********************
'功能: 设置到位信号功能 , 锁定所有轴的逻辑位置和实际位置
'参数:
'    axis—参照轴
'    mode—位置锁存工作模式|0:无效
'                         |1:有效
'    regi—计数器模式  |0:逻辑位置
'                      |1:实际位置
'    logical—电平信号 |0:由高到低
'                      |1:由低到高
'返回值 0: 正确 1: 错误
'说明:    使用指定轴axis的IN信号作为触发信号
'*******************************************************************/
Public Function Setup_LockPosition(ByVal axis As Long, ByVal mode As Long, ByVal regi As Long, ByVal logical As Long) As Integer
    
    Result = set_lock_position(0, axis, mode, regi, logical)
    
    Setup_LockPosition = Result
    
End Function


'**************************获取锁定的位置**************************
'功能: 获取锁定的位置
'参数:
'    axis 轴号
'    pos 锁存的位置
'返回值 0: 正确 1: 错误
'******************************************************************
Public Function Get_LockPosition(ByVal axis As Long, pos As Long) As Integer

    Result = get_lock_position(0, axis, pos)
    
    Get_LockPosition = Result
    
End Function

'**************************清除锁存状态**************************
'功能: 清除锁存状态
'参数:
'    axis 轴号(1 - 4)
'返回值 0: 正确 1: 错误
'******************************************************************
Public Function Clr_LockStatus(ByVal axis As Long) As Integer

    Result = clr_lock_status(0, axis)
    
    Clr_LockStatus = Result
    
End Function


