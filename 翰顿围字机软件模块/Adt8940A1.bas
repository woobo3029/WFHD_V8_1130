Attribute VB_Name = "adt8940a1"
Option Explicit
'******************************基本库函数****************************
Declare Function adt8940a1_initial Lib "8940A1.dll" () As Integer
' 功能：初始化卡
'返回值>0时，表示8940A1卡的数量。如果为3，则下面的可用卡号分别为0、1、2as integer
'返回值=0时，说明没有安装8940A1卡as integer
'返回值<0时，-1表示没有安装端口驱动程序，-2表示PCI桥存在故障。

Declare Function get_lib_version Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'功能：获取当前库版本

Declare Function set_pulse_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Integer, ByVal logic As Long, ByVal dir_logic As Long) As Integer
'功能：设置输出脉冲的工作方式
'cardno 卡号
'axis 轴号(1 - 4)
'value       0：脉冲+脉冲方式        1：脉冲+方向方式
'logic       0: 正逻辑脉冲           1: 负逻辑脉冲
'dir-logic   0：方向输出信号正逻辑    1：方向输出信号负逻辑
'返回值      0: 正确 1: 错误
'默认模式：脉冲+方向，正逻辑脉冲，方向输出信号正逻辑

Declare Function set_limit_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal v1 As Integer, ByVal v2 As Integer, ByVal dir_logic As Integer) As Integer
'功能：设定正负方向限位输入nLMT信号的模式
'参数:
'cardno 卡号
'axis 轴号(1 - 4)
'v1 0: 正限位有效 1: 正限位无效
'v2 0: 负限位有效 1: 负限位无效
'logic 0: 低电平有效 1: 高电平有效
'返回值 0: 正确 1: 错误
'默认模式为：正限位有效，负限位有效，低电平有效

Declare Function set_stop0_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal v As Integer, ByVal logic As Long) As Integer
'功能：设定stop0输入信号的模式
'cardno 卡号
'axis   轴号(1 - 4)
'v      0: 无效       1: 有效
'logic  0: 低电平有效 1: 高电平有效
'返回值 0: 正确       1: 错误
'默认模式为: 无效

Declare Function set_stop1_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Long, ByVal v As Long, ByVal logic As Long) As Integer
'功能：设定stop1输入信号的模式
'cardno     卡号
'axis       轴号(1 - 4)
'v          0: 无效 1: 有效
'logic      0: 低电平有效 1: 高电平有效
'返回值      0: 正确 1: 错误
'默认模式为: 无效

Declare Function get_status Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'功能: 获取各轴的驱动状态
'cardno     卡号
'axis       轴号(1 - 4)
'v 驱动状态指针
'           0:  驱动结束 非0: 正在驱动
'返回值     0: 正确 1: 错误

Declare Function get_inp_status Lib "8940A1.dll" (ByVal cardno As Integer, ByRef value As Long) As Integer
'功能: 获取插补的驱动状态
'cardno     卡号
'v 插补状态指针
'           0: 插补结束 1: 正在插补
'返回值     0: 正确     1: 错误

Declare Function set_acc Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal add As Long) As Integer
'功能: 加速度设定
'cardno     卡号
'axis       轴号(1 - 4)
'Add        范围(1 - 64000)
'加速度实际值  add*125
'返回值     0: 正确     1: 错误

Declare Function set_startv Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal startv As Long) As Integer
'功能: 初始速度设定
'cardno     卡号
'axis       轴号(1 - 4)
'startv      范围(1-2M)
'返回值     0: 正确 1: 错误

Declare Function set_speed Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal speed As Long) As Integer
'功能: 驱动速度设定
'cardno     卡号
'axis       轴号(1 - 4)
'speed      范围(1-2M)
'返回值      0: 正确 1: 错误


Declare Function set_command_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Long) As Integer
'功能: 逻辑位置设定
'cardno     卡号
'axis       轴号(1 - 4)
'value      范围(-2147483648～+2147483647)
'返回值     0: 正确 1: 错误

Declare Function set_actual_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Long) As Integer
'功能: 实际位置设定
'cardno     卡号
'axis       轴号(1 - 4)
'value      范围(-2147483648～+2147483647)
'返回值     0: 正确 1: 错误

Declare Function get_command_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'功能: 获取各轴的逻辑位置
'cardno     卡号
'axis       轴号(1 - 4)
'value      逻辑位置的指针
'返回值     0: 正确 1: 错误

Declare Function get_actual_pos Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'功能: 获取各轴的实际位置
'cardno     卡号
'axis       轴号(1 - 4)
'value      实际位置的指针
'返回值     0: 正确 1: 错误

Declare Function get_speed Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef value As Long) As Integer
'功能: 获取各轴的当前驱动速度
'cardno     卡号
'axis       轴号(1 - 4)
'value      当前驱动速度的指针
'返回值     0: 正确 1: 错误

Declare Function get_out Lib "8940A1.dll" (ByVal cardno As Integer, ByVal number As Integer) As Integer
'*****************************************************
'功能: 获取输出点
'参数:
'    cardno 卡号
'    number 输出点
'返回值      获取输出端口的当前状态,0: 低电平   1: 高电平  -1:错误
'*****************************************************/

Declare Function pmove Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal value As Long) As Integer
'功能: 定量驱动
'cardno     卡号
'axis       轴号(1 - 4)
'value      输出的脉冲数(-268435455～+268435455)
'           >0：正方向驱动      <0：负方向驱动
'返回值     0: 正确     1: 错误

Declare Function dec_stop Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'功能: 驱动减速停止
'cardno     卡号
'axis       轴号(1 - 4)
'返回值     0: 正确 1: 错误

Declare Function sudden_stop Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'功能: 驱动立即停止
'cardno     卡号
'axis       轴号(1 - 4)
'返回值     0: 正确 1: 错误

Declare Function inp_move2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long) As Long
'功能: 两轴直线插补
'cardno         卡号
'axis1,axis2    参与插补的轴号
'pulse1,pulse2  移动的相对距离(-8388608～+8388607)
'返回值         0: 正确 1: 错误

Declare Function inp_move3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long) As Long
'功能: 三轴直线插补
'cardno                 卡号
'axis1,axis2,axis3      参与插补的轴号
'pulse1,pulse2,pulse3   移动的相对距离(-8388608～+8388607)
'返回值                 0: 正确 1: 错误

Declare Function inp_move4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long) As Long
'功能: 四轴直线插补
'cardno 卡号
'pulse1,pulse2,pulse3,pulse4 XYZA四轴移动的相对距离(-8388608～+8388607)
'返回值 0: 正确 1: 错误

Declare Function read_bit Lib "8940A1.dll" (ByVal cardno As Integer, ByVal number As Long) As Long
'功能: 读取输入点
'cardno 卡号
'number 输入点(0 - 39)
'返回值 0: 低电平 1: 高电平 -1: 错误

Declare Function write_bit Lib "8940A1.dll" (ByVal cardno As Integer, ByVal number As Long, ByVal value As Long) As Long
'功能: 输出
'cardno 卡号
'number 输出点(0 - 15)
'value  0: 低电平   1: 高电平
'返回值  0: 正确     1: 错误

Declare Function get_hardware_ver Lib "8940A1.dll" (ByVal cardno As Integer) As Double
'功能: 获取硬件版本
'cardno     卡号
'返回值     1: 硬件第一版         2:硬件第二版
'这里的1、2只是是暂时做说明用，返回值是多少就为多少，目前硬件版本为1.1

Declare Function set_suddenstop_mode Lib "8940A1.dll" (ByVal cardno As Integer, ByVal v As Integer, ByVal logic As Integer) As Integer
'功能: 硬件停止模式设置
'cardno     卡号
'v          0: 无效 1: 有效
'logic      0: 低电平有效 1: 高电平有效
'返回值     0: 正确 1: 错误
'硬件停止信号固定使用P2端子板25引脚 (IN31)

Declare Function set_delay_time Lib "8940A1.dll" (ByVal cardno As Integer, ByVal time As Long) As Integer
'功能: 设定延时时间
'cardno 卡号
'time   延时时间
'返回值 0: 正确 1: 错误
'时间单位为1/8us

Declare Function get_delay_status Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'功能: 获取延时状态
'cardno 卡号
'返回值  0: 延时结束 1: 延时进行中

'*********************************************//
'               复合驱动类                     //
'*********************************************//
Declare Function set_symmetry_speed Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*******************************************************
'功能:   设定对称加减速的值
'参数:
'    cardno 卡号
'    axis 轴号
'    lspd 起步速度
'    hspd 驱动速度
'    tacc 加速时间
'返回值 0: 正确 1: 错误
'*******************************************************

Declare Function symmetry_relative_move Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'********************************************************
'*功能:参照当前位置,以对称加减速进行定量移动
'*参数:
'      cardno -卡号
'      axis ---轴号
'      pulse --脉冲
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'*********************************************************

Declare Function symmetry_absolute_move Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal pulse As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*********************************************************
'*功能:参照零点位置,以对称加减速进行定量移动
'*参数:
'      cardno -卡号
'      axis ---轴号
'      pulse --脉冲
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'**********************************************************

Declare Function symmetry_relative_line2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'**********************************************************
'*功能:参照当前位置,以对称加减速进行直线插补
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
'***********************************************************

Declare Function symmetry_absolute_line2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'***********************************************************
'*功能:参照零点位置,以对称加减速进行直线插补
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
'************************************************************/

Declare Function symmetry_relative_line3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'************************************************************
'*功能:参照当前位置,以对称加减速进行直线插补
'*参数:
'      cardno -卡号
'      axis1 ---轴号1
'      axis2 ---轴号2
'      axis3 ---轴号3
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'***************************************************************

Declare Function symmetry_absolute_line3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'**************************************************************
'功能: 参照零点位置 , 以对称加减速进行直线插补
'参数:
'      cardno -卡号
'      axis1 ---轴号1
'      axis2 ---轴号2
'      axis3 ---轴号3
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'返回值 0: 正确 1: 错误
'****************************************************************

Declare Function symmetry_relative_line4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*****************四轴直线插补相对运动****************
'*功能:参照当前位置,以加减速进行直线插补
'*参数:
'      cardno -卡号
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      pulse4 --脉冲4
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'******************************************************

Declare Function symmetry_absolute_line4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal lspd As Long, ByVal hspd As Long, ByVal tacc As Double) As Integer
'*****************四轴对称直线插补绝对运动****************
'*功能:参照零点位置,以对称加减速进行直线插补
'*参数:
'      cardno -卡号
'      pulse1 --脉冲1
'      pulse2 --脉冲2
'      pulse3 --脉冲3
'      pulse4 --脉冲4
'      lspd ---低速
'      hspd ---高速
'      tacc---加速时间(单位:秒)
'******************************************************


'//*********************************************//
'//               外部驱动                    //
'//*********************************************//

Declare Function manual_pmove Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal pos As Long) As Integer
'/************************外部信号定量驱动函数**********************
'功能: 外部信号定量驱动函数
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'    pos 脉冲
'返回值 0: 正确 1: 错误
'    说明:(1)发出定量脉冲，但驱动没有立即进行，需要等到外部信号电平发生变化
'         (2)可以使用普通按钮,也可以接手轮
'******************************************************************/

Declare Function manual_continue Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'/************************外部信号连续驱动函数**********************
'功能: 外部信号连续驱动函数
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'返回值 0: 正确 1: 错误
'    说明:(1)发出定量脉冲，但驱动没有立即进行，需要等到外部信号电平发生变化
'         (2)可以使用普通按钮,也可以接手轮
'******************************************************************/

Declare Function manual_disable Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'/***********************关闭外部信号驱动使能***********************
'功能: 关闭外部信号驱动使能
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'返回值 0: 正确 1: 错误
'******************************************************************/

'//*********************************************//
'//               位置锁存                    //
'//*********************************************//

Declare Function set_lock_position Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal mode As Integer, ByVal regi As Integer, ByVal logical As Integer) As Integer
'/****************************位置锁存设置函数**********************
'功能: 设置到位信号功能 , 锁定所有轴的逻辑位置和实际位置
'参数:
'    axis—参照轴
'    mode—位置锁存工作模式|0:无效
'                        |1:有效
'    regi—计数器模式  |0:逻辑位置
'                      |1:实际位置
'    logical—电平信号 |0:由高到低
'                      |1:由低到高
'返回值 0: 正确 1: 错误
'说明:    使用指定轴axis的IN信号作为触发信号
'*******************************************************************/

Declare Function get_lock_status Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef v As Integer) As Integer
'/*************************获取锁存状态***********************
'功能: 获取锁存状态
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'    V            0|未执行同步操作
'                 1|执行过同步操作
'返回值 0: 正确 1: 错误
'说明:    利用该函数可以捕捉位置锁存是否执行
'******************************************************************/

Declare Function get_lock_position Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByRef pos As Long) As Integer
'/**************************获取锁定的位置**************************
'功能: 获取锁定的位置
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'    pos 锁存的位置
'返回值 0: 正确 1: 错误
'******************************************************************/

Declare Function clr_lock_status Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer) As Integer
'/**************************清除锁存状态**************************
'功能: 清除锁存状态
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'返回值 0: 正确 1: 错误
'******************************************************************/

'//*********************************************//
'//               硬件缓存                    //
'//*********************************************//
Declare Function fifo_inp_move1 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal pulse1 As Long, ByVal speed As Long) As Integer
'/**************************单轴缓存**************************
'功能: 单轴缓存
'参数:
'    cardno 卡号
'    axis1 轴号(1 - 4)
'    pulse1 缓存的脉冲
'    speed 缓存的速度
'返回值 0: 正确 1: 错误
'说明:共有2048个缓存空间，每条单轴缓存指令占用3个空间，可缓存682条指令
'******************************************************************/

Declare Function fifo_inp_move2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal speed As Long) As Integer
'/**************************两轴缓存**************************
'功能: 两轴缓存
'参数:
'    cardno 卡号
'    axis1 轴号(1 - 4)
'    axis2 轴号(1 - 4)
'    pulse1 缓存的脉冲数
'    pulse2 缓存的脉冲数
'    speed 缓存的速度
'返回值 0: 正确 1: 错误
'说明:共有2048个缓存空间，每条两轴缓存指令占用4个空间，可缓存512条指令
'******************************************************************/

Declare Function fifo_inp_move3 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal axis3 As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal speed As Long) As Integer
'/**************************三轴缓存**************************
'功能: 三轴缓存
'参数:
'    cardno 卡号
'    axis1 轴号(1 - 4)
'    axis2 轴号(1 - 4)
'    axis3 轴号(1 - 4)
'    pulse1 缓存的脉冲数
'    pulse2 缓存的脉冲数
'    pulse3 缓存的脉冲数
'    speed 缓存的速度
'返回值 0: 正确 1: 错误
'说明:共有2048个缓存空间，每条三轴缓存指令占用5个空间，可缓存409条指令
'******************************************************************/

Declare Function fifo_inp_move4 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal pulse1 As Long, ByVal pulse2 As Long, ByVal pulse3 As Long, ByVal pulse4 As Long, ByVal speed As Long) As Integer
'/**************************四轴缓存**************************
'功能: 四轴缓存
'参数:
'    cardno 卡号
'    axis1 轴号(1 - 4)
'    axis2 轴号(1 - 4)
'    axis3 轴号(1 - 4)
'    axis4 轴号(1 - 4)
'    pulse1 缓存的脉冲数
'    pulse2 缓存的脉冲数
'    pulse3 缓存的脉冲数
'    pulse4 缓存的脉冲数
'    speed 缓存的速度
'返回值 0: 正确 1: 错误
'说明:共有2048个缓存空间，每条四轴缓存指令占用6个空间，可缓存341条指令
'******************************************************************/

Declare Function reset_fifo Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'/**************************重设缓存**************************
'功能: 清除缓存
'参数:
'    cardno 卡号
'返回值 0: 正确 1: 错误
'******************************************************************/

Declare Function read_fifo_count Lib "8940A1.dll" (ByVal cardno As Integer, ByRef value As Integer) As Integer
'/**************************读取缓存数**********************
'功能:读取缓存数，存放进去的指令还剩多少条未执行
'参数:
'    cardno 卡号
'    value  未执行的指令所占的字节数
'返回值 0: 正确 1: 错误
'******************************************************************/

Declare Function read_fifo_empty Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'/**************************读取缓存状态**********************
'功能: 读取缓存是否为空
'参数:
'    cardno 卡号
'返回值 0: 非空 1: 空
'******************************************************************/

Declare Function read_fifo_full Lib "8940A1.dll" (ByVal cardno As Integer) As Integer
'/**************************读取缓存状态**********************
'功能:读取缓存是否满了，满了之后将不能再存数据
'参数:
'    cardno 卡号
'返回值 0: 未满 1: 满
'******************************************************************/

Declare Function home1 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal backDir As Integer, ByVal logical0 As Integer, ByVal logical1 As Integer, ByVal homeStartV As Long, ByVal homeSpeed As Long, ByVal homeAcc As Long, ByVal searchRange As Long, ByVal searchSpeed As Long, ByVal phaseSpeed As Long, ByVal pulseUnit As Long) As Integer
'**************************单轴回原点**********************
'功能: 执行单轴回原点运动
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'    backDir                         回原点方向  0：正向    1：负向
'    logical0                        回原点stop0设置  0:低电平有效 1:高电平有效
'    logical1                        回原点stop1设置  0:低电平有效 1:高电平有效   -1：无效（不搜索Z相）
'    homeStartV                      回原点启始速度，取值范围：0-2M
'    homeSpeed                       回原点驱动速度，取值范围：0-2M
'    homeAcc                         回原点加速度，取值范围：0-64000
'    searchRange 原点范围(不宜过大)
'    searchSpeed stop0搜索速度(不宜过高)
'    phaseSpeed Z相搜索速度(不宜过高)
'    pulseUnit 每转脉冲
'
'返回值  0:回原点成功;   -1:参数错误;    -2：回原点失败,(碰到限位或原点范围过小);     1：回原点被中止
'说明:
' (1) 回原点分为四大步:
'     第一步:快速接近stop0(logical0原点设置)，找到stop0;
'     第二步:慢速反向离开stop0，反向移动指定原点范围脉冲数;
'     第三步:再次慢速接近stop0;
'     第四步:慢速接近stop1(logical1编码器Z相).
' (2) 第四步可以选择是否执行,通过logical1来选择.
' (3) 若需多轴回原点,必须等待上一轴回原点结束后，才能执行下一轴的回原点动作.
'*****************************************************

Declare Function inp_arc Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal cood As Long) As Integer
'*************************功能:两轴圆弧插补**************************
'功能：     任意两轴圆弧插补运动 ，本函数用两轴插补指令封装，通过普通插补实现
'参数:
'    cardno 卡号
'    axis1 axis2 轴号(1 - 4)
'    dir                 画圆方向    0:顺时针圆 ;1：逆时针圆
'    cood[]              圆弧上三点的坐标(起点,中间点,终点)共含六个元素
'
'  返回值：  -3:三点不能构成圆弧， -2:限位信号停止；-1:参数错误；    0:成功；  1:圆弧插补中止.
'  注意：默认参与圆弧插补的两个轴脉冲当量相同;
'  如果插补轨迹为整圆，中间点需设置成与起点关于圆心对称的点.
'********************************************************************
Declare Function fifo_arc Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal speed As Integer, ByVal ccood As Long) As Integer
'*************************功能:两轴圆弧插补缓存实现**************************
'功能：     任意两轴圆弧插补运动，本函数用硬件缓存插补指令封装，通过缓存实现。
'参数:
'    cardno 卡号
'    axis1 axis2 轴号(1 - 4)
'    speed 插补速度
'    cood[]              圆弧上三点的坐标(起点,中间点,终点)共含六个元素
'
'返回值：  -3:三点不能构成圆弧;  -2-限位信号停止;    -1:参数错误;    0:成功;     1:圆弧插补中止.
'注意：默认参与圆弧插补的两个轴脉冲当量相同;
'      如果插补轨迹为整圆，中间点需设置成与起点关于圆心对称的点.
'
'********************************************************************
Declare Function continue_move1 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis As Integer, ByVal dir As Integer) As Integer
'*************************功能:单轴连续运动**************************
'功能:      单轴连续运动
'参数:
'    cardno 卡号
'    axis 轴号(1 - 4)
'    dir                 0:正向 ;1：负向
'
'返回值：   -1:限位信号停止; 1:错误;     0:正确.
'注意:写入驱动命令前,一定要正确地设定速度参数.
'********************************************************************

Declare Function continue_move2 Lib "8940A1.dll" (ByVal cardno As Integer, ByVal axis1 As Integer, ByVal axis2 As Integer, ByVal dir1 As Integer, ByVal dir2 As Integer) As Integer
'*************************功能:两轴连续运动**************************
'功能:      两轴连续运动
'参数:
'    cardno 卡号
'    axis1 轴号(1 - 4)
'    axis2 轴号(1 - 4)
'    dir1                0:正向; 1：负向
'    dir2                0:正向; 1：负向
'
'返回值：   -1:限位信号停止; 1:错误;    0:正确.
'注意:写入驱动命令前,一定要正确地设定速度参数.
'********************************************************************


Public Sub MyProc()

    DoEvents

End Sub

