Attribute VB_Name = "Module1"
'#define  IPOL_STATE_Sleeping            0        //停止状态
'#define  IPOL_STATE_Stoped          1        //被轴停止命令结束
'#define  IPOL_STATE_LineEndStoped   2        //被行结束停止命令停止
'#define  IPOL_STATE_Ended               3        //被End结束
'#define  IPOL_STATE_Awaiting            4        //插补缓冲区空
'#define  IPOL_STATE_FRateZero           5        //被进给倍率为0
'#define  IPOL_STATE_Suspended           6        //被进给暂停挂起
'#define  IPOL_STATE_Running             7        //插补正在进行

'#define IPOL_DATA_BUFF_NUM    256


'#define NULLITY                         0               //0=轴无效;
'#define OPEN_LOOP_PULSE         1               //1=开环脉冲模式;
'#define CLOSE_LOOP_PULSE        2               //2=位置闭环脉冲输出模式;
'#define CLOSE_LOOP_DAV          3               //3=位置闭环模拟量输出模式;
'#define SIMPLE_DV_OUT           4               //4=单纯电压输出模式;
'#define SIMPLE_PWM_OUT          5               //5=单纯PWM脉冲输出模式;  9030固件
'#define AUTO_USER_SET           6               //6=自动用户设定模式;           最后只能是1或3


Declare Function InitCard_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal Axis0 As Integer, ByVal Axis1 As Integer, ByVal Axis2 As Integer, ByVal Axis3 As Integer, _
ByVal PWM_DA_Mode As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : ExitCard_9030
' ' 函数编号 : 23

' ' 描述     :退出9030卡,并释放所占资源
' ' 参数:
        
'          Board_NO: 0-3,板号
          

' ' 返回值:  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ExitCard_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisWorkMode_9030
' ' 函数编号 : 98

' ' 描述     : 设置轴工作模式

' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'WorkMode:            设置轴工作模式:
'                                                                        0=轴无效;
'                                                                        1=开环脉冲模式;
'                                                                        2=位置闭环脉冲输出模式;
'                                                                        3=位置闭环模拟量输出模式;
'                                                                        4=单纯电压输出模式;
'                                                                        5=单纯PWM脉冲输出模式;  9030固件
'
' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisWorkMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal WorkMode As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisKP_9030
' ' 函数编号 : 99
'
' ' 描述     : 设置轴PID调节比例系数

' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          Kp            : PID调节比例系数; 范围: 整型数 Lib "dfjzh9030dll.dll" (0-65535) As Integer  单位: 无。

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisKP_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal Axis_No As Integer, ByVal Kp As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisKI_9030
' ' 函数编号 : 100

' ' 描述     : 设置轴PID调节积分系数

' ' 参数     :
'
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          Ki            : 轴PID调节积分系数; 范围: 整型数 Lib "dfjzh9030dll.dll" (0-65535) As Integer  单位: 无。

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisKI_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal Ki As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisKD_9030
' ' 函数编号 : 101

' ' 描述     : 设置轴PID调节微分系数
'
' ' 参数     :
'
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
'
'          Kd            : 轴PID调节微分系数; 范围: 整型数 Lib "dfjzh9030dll.dll" (0-65535) As Integer  单位: 无。

 '' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisKD_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal Kd As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisIL_9030
' ' 函数编号 : 102

' ' 描述     : 设置轴PID调节积分限

' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          IL            : 轴PID调节积分限; 范围: 整型数 Lib "dfjzh9030dll.dll" (0-65535) As Integer  单位: 无。

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisIL_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal IL As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisFVRate_9030
' ' 函数编号 : 103

' ' 描述     : 设置轴PID调节速度前馈系数

' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'EcLine:                   轴电机编码器反馈线数  Lib "dfjzh9030dll.dll" (四倍频前)
'          MaxSpeed      : 轴电机最高转速;  单位: RPM  Lib "dfjzh9030dll.dll" (转/分钟)。
'
' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFVRate_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal EcLine As Long, ByVal MaxSpeed As Double) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisFV_9030
' ' 函数编号 : 103

' ' 描述     : 设置轴PID调节速度前馈

' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          FV            : 轴PID调节速度前馈; 范围: 整型数 Lib "dfjzh9030dll.dll" (0-65535) As Integer  单位: 无。

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFV_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal FV As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisPEL_9030
' ' 函数编号 : 104

' ' 描述     : 设置轴位置闭环误差极限

' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          PosErrL  : 设置轴位置闭环误差极限; 范围: 整型数 Lib "dfjzh9030dll.dll" (0-65535) As Integer  单位: 无。

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisPEL_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal PosErrL As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisDAOut_9030
' ' 函数编号 : 105
'
' ' 描述     : 设置轴输出电压

' ' 参数     :
'
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
'
'          DA_Avlue :   -10 - +10 V; 精度: 1/6000; 即: 精度大于12位 Lib "dfjzh9030dll.dll" (4096),不到13位 Lib "dfjzh9030dll.dll" (8192).
'                       该DA输出与PWM输出占用同一个硬件资源 , 通过9030端子板上的跳线选择是DA输出
'                                   还是PWM输出.

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisDAOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal DA_Avlue As Double) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisPWMOut_9030
' ' 函数编号 : 106

' ' 描述     : 设置轴PWM脉冲输出
' ' 参数     :
        
'          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          frequency:   18-1500000;    PWM输出频率; 单位: 赫兹 Lib "dfjzh9030dll.dll" (Hz)
'          Pulse_Highf: 占空比,范围:0.0-1.0;  精度:不小于1%

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
'Declare Function SetAxisPWMOut_9030 Lib "dfjzh9030dll.dll" (ByVal    As Integer Board_NO,ByVal    As Integer Axis_No,ByVal     As Long frequency,ByVal    As Single Pulse_Highf) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : AxisPWMStop_9030
' ' 函数编号 : 107

' ' 描述   : 轴PWM频率输出停止

 '' 参数   :
        
 '         Board_NO :   0－3, 板号
 '          Axis_No  : 0－3, 轴号

 '' 返回值 :  0或-1,-1=不成功,0=成功
 ''
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function AxisPWMStop_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : Home_9030
' ' 函数编号 : 3

' ' 描述     : 轴位置清零

' ' 参数:
        
'          Board_NO: 0－3, 板号
'          Axis_No : 0－3, 轴号
'
' ' 返回值:  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Home_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : HomeFB_9030
' ' 函数编号 : 108

' ' 描述     :  轴位置编码器清零

' ' 参数:
'
'          Board_NO: 0－3, 板号
'          Axis_No : 0－3, 轴号

' ' 返回值:  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function HomeFB_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisIO_9030
' ' 函数编号 : 15

' ' 描述     :  设置轴IO是否有效
'
' ' 参数     :
        
'          Board_NO: 0－3, 板号
'          Axis_No : 0－3, 轴号
'          PHN_flag: 正限位、Home点、负限位：1=正限位；2=Home点；3=负限位
'          Mode    : 0、1、2或3, 0=该点无效；3=自由挂接模式；1、2=对正负限位复用固定点中断模式；1、2=对Home点固定点模式
'          H_L_Act : 0或1, 0=该点低电平有效,1=该点高电平有效
'          IO_index: 1-16, 在Mode=3时 Lib "dfjzh9030dll.dll" (自由挂接模式) As Integer该点挂接到通用输入点的那一点

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisIO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
                                                          ByVal PHN_flag As Integer, ByVal Mode As Integer, _
                                                          ByVal H_L_Act As Integer, ByVal IO_Index As Integer) As Integer


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : SetAxisIOHL_9030
' ' 函数编号 : 85

' ' 描述     :  设置轴IO是高电平有效还是低电平有效

' ' 参数     :
        
'          Board_NO: 0－3, 板号
'          Axis_No : 0－3, 轴号
'          PLimit  : 0或1, 0=该轴正限位低电平有效,1=该轴正限位高电平有效
'          NLimit  : 0或1, 0=该轴负限位低电平有效,1=该轴负限位高电平有效
'          Home    : 0或1, 0=该轴Home点低电平有效,1=该轴Home点高电平有效

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
'//Declare Function SetAxisIOHL_9030 Lib "dfjzh9030dll.dll" (ByVal    As Integer Board_NO,ByVal    As Integer Axis_No,ByVal    As Integer PLimit,ByVal    As Integer NLimit,ByVal    As Integer Home) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : GoHome_9030
' ' 函数编号 : 24
'
' ' 描述   : 轴回Home点

' ' 参数   :
        
          '          Board_NO : 0－3, 板号
'          Axis_No  : 0－3, 轴号
          
'          goHomeVel    : 轴回Home点速度; 范围: 长整型数;  单位: Hz，轴的脉冲频率。
'          LeaveHomeVel : 轴离开Home点速度; 范围: 长整型数;    单位: Hz，轴的脉冲频率。
'          LeaveHomePos : 轴离开Home点距离; 范围: 长整型数,必须是正值;  单位: 脉冲个数。
'          LookZIndexVel: 轴找一转脉冲速度; 范围: 长整型数;    单位: Hz，轴的脉冲频率。
'PulseNum:                    轴编码器反馈每转脉冲数  Lib "dfjzh9030dll.dll" (或一转脉冲的间隔脉冲个数 Lib "dfjzh9030dll.dll" (光栅尺反馈))
'          Z_IndexFlag  : 0或1;当轴为位置闭环模式 Lib "dfjzh9030dll.dll" (2和3模式)时,可设回零找一转脉冲模式: 0=回零不找一转脉冲;1=回零找一转脉冲

 '' 返回值 :  0或-1,-1=不成功,0=成功
 ''
 ''
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GoHome_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal goHomeVel As Long, ByVal LeaveHomeVel As Long, ByVal LeaveHomePos As Long, ByVal LookZIndexVel As Long, ByVal PulseNum As Long, ByVal Z_IndexFlag As Integer) As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '
' ' 函数名   : LookZIndex_9030
' ' 函数编号 : 114
'
' ' 描述   : 找一转脉冲 Lib "dfjzh9030dll.dll" (轴回Home点)

' ' 参数   :
'
'          Board_NO : 0－3, 板号
 '         Axis_No  : 0－3, 轴号
'
'          LookZIndexVel: 轴找一转脉冲速度; 范围: 长整型数;    单位: Hz，轴的脉冲频率。
'PulseNum:                    轴编码器反馈每转脉冲数  Lib "dfjzh9030dll.dll" (或一转脉冲的间隔脉冲个数 Lib "dfjzh9030dll.dll" (光栅尺反馈))

' ' 返回值 :  0或-1,-1=不成功,0=成功
' '
' '
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LookZIndex_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal LookZIndexVel As Long, ByVal PulseNum As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Set_Emergency_Stop_9030
 ' 函数编号 : 29

 ' 描述   : 设置IO急停信号

 ' 参数   :
        
 '         Board_NO : 0－3, 板号
 '         Mask     : 0－8, 急停信号与通用输入点I1-I8相联,0表示无效,缺省为无效
 '         Mode     :  Lib "dfjzh9030dll.dll" (保留)

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_Emergency_Stop_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Mask As Integer, ByVal Mode As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisPos_9030
 ' 函数编号 : 4

 ' 描述     : 设置轴的位置

 ' 参数     :
        
  '        Board_NO : 0－3, 板号
   '       Axis_No  : 0－3, 轴号
  '
   '       position : 轴的位置; 范围: 长整型数;  单位: 脉冲个数。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal position As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisVel_9030
 ' 函数编号 : 5

 ' 描述     : 设置轴的速度

 ' 参数     :
        
          'Board_NO : 0－3, 板号
          'Axis_No  : 0－3, 轴号
          
         ' velocity : 轴的速度; 范围: 长整型数;    单位: Hz，轴的脉冲频率。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisStartVel_9030
 ' 函数编号 : 88

 ' 描述   : 设轴的起跳速度

 ' 参数   :
        
       '   Board_NO : 0－3, 板号
       '   Axis_No  : 0－3, 轴号
          
       '   velocity : 轴的起跳速度; 范围: 正长整型数;    单位: Hz，轴的脉冲频率。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisStartVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisStopVel_9030
 ' 函数编号 : 95

 ' 描述   : 设轴的停止速度,固件3.0版

 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号
          
        '  velocity : 轴的停止速度; 范围: 正长整型数;    单位: Hz，轴的脉冲频率。缺省值: 16

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisStopVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisDec_9030
 ' 函数编号 : 96

 ' 描述   : 设轴的减速度,固件3.0版

 ' 参数   :
        
        '  Board_NO : 0－3, 板号
       '   Axis_No  : 0－3, 轴号
          
        '  deceleration: 轴的减速度; 范围: 长整型数,必须是正值;  单位: 脉冲数 / 秒平方。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal deceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisStopDec_9030
 ' 函数编号 : 122

 ' 描述   : 设轴的Stop命令的减速度,固件5.1版

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          
         ' deceleration: 轴的减速度; 范围: 长整型数,必须是正值;  单位: 脉冲数 / 秒平方。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisStopDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal deceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisAcc_9030
 ' 函数编号 : 6

 ' 描述     : 设置轴的加速度

 ' 参数     :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号
          
       '   acceleration: 轴的加速度; 范围: 长整型数,必须是正值;  单位: 脉冲数 / 秒平方。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisAcc_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal acceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : StartAxis_9030
 ' 函数编号 : 7

 ' 描述     : 轴开始运行,位置模式

 ' 参数     :
        
        '  Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
'

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function StartAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名 : StopAxis_9030
 ' 函数编号 : 9

 ' 描述   : 轴停止,按所设加速度减速停止

 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function StopAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : AbortAxis_9030
 ' 函数编号 : 19

 ' 描述   : 轴停止,在10毫秒之内减速停止
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
         '

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function AbortAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : CeaseAxis_9030
 ' 函数编号 : 20

 ' 描述   : 轴停止,立即停止,无减速过程

 ' 参数   :
        
        '  Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function CeaseAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : StartAxisVel_9030
 ' 函数编号 : 8

 ' 描述   : 设置轴速度模式,并按所设速度开始运行

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          
         ' velocity : 轴的速度; 范围: 长整型数;  单位: Hz，轴的脉冲频率。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function StartAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal Axis_No As Integer, ByVal velocity As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisOffset_9030
 ' 函数编号 : 84

 ' 描述   : 设置轴位置偏移值

 ' 参数   :
        
          'Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          
         ' offset   : 轴位置的偏移值; 范围: 长整型数;  单位: 脉冲个数。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisOffset_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal offset As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisFBOffset_9030
 ' 函数编号 : 109

 ' 描述     :   设轴位置编码器偏移值

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          
         ' offset   : 轴位置的偏移值; 范围: 长整型数;  单位: 脉冲个数。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFBOffset_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal offset As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisOPC_9030
 ' 函数编号 : 112

 ' 描述     :   设置轴电压输出0点补偿

 ' 参数   :
        
          'Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          
         ' OPC_value: 轴电压输出0点补偿;  范围: -1000 - +1000 。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisOPC_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal OPC_value As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisMotorOnOff_9030
 ' 函数编号 : 110

 ' 描述     :   设轴电机On或Off，使能

 ' 参数   :
        
       '   Board_NO : 0－3, 板号
       '   Axis_No  : 0－3, 轴号
       '
       '   OnOff    : 0,1;       0=Off,1=On。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisMotorOnOff_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal OnOff As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisOutMode_9030
 ' 函数编号 : 17

 ' 描述   : 设置轴的输出模式
 ' 参数   :
        
       '   Board_NO : 0－3, 板号
       '   Axis_No  : 0－3, 轴号

       '   Mode_A   : 0或1; 0=共阳极输出,1=共阴极输出
       '   Mode_B   : 0或1; 0=脉冲-方向模式输出,1=脉冲-脉冲模式输出
       '   Mode_C   : 0或1; 0=轴方向正常,1=轴方向反转
  

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisOutMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal Mode_A As Integer, ByVal Mode_B As Integer, ByVal Mode_C As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisSAcce_9030
 ' 函数编号 : 72

 ' 描述   : 设置轴S型加速度
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号

       '   PowerFlag: 1-4; S型加速度 Lib "dfjzh9030dll.dll" (1,2,3,4)指数
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisSAcce_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal PowerFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisTEC_9030
 ' 函数编号 : 73

 ' 描述   : 设置轴螺距反向间隙补偿
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
        '  ErrorV   : 轴螺距反向间隙补偿值; 范围: 0-32767;  单位: 脉冲个数。
        '  TimeNum  : 补偿所用时间; 范围: 1-20;  单位: 毫秒。
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisTEC_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal ErrorV As Long, ByVal TimeNum As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisTECData_9030
 ' 函数编号 : 111

 ' 描述   : 设置轴螺距误差补偿数据
 ' 参数   :
        
         ' Board_NO :    0－3, 板号
       '   Axis_No  :    0－3, 轴号
       '   Mode          :       2或3            2=反向间隙+螺距误差补偿;3=双向螺距误差补偿
        '  EffectNum :   Mode=2时，1-512；       Mode=3时，0-256；       有效数据数
        '  BasePoint :   基点位置， 数组TEData[0]对应的位置,单位: 脉冲个数。
        '  NodeLen       :       数据间距离，单位: 脉冲个数。
        '  Direc         :       1 或 -1；方向: Mode=2时， 反向间隙补偿方向; Mode=3时, 前 Lib "dfjzh9030dll.dll" (双向)一组数据方向
        '  ReverseGap:   反向间隙补偿数据，范围: 0-32767;  单位: 脉冲个数。

        '  TEData        :       螺距误差补偿数据数组，数组个数固定为512个，每个数据范围: -32768 - +32767;  单位: 脉冲个数。
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisTECData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal Mode As Integer, ByVal EffectNum As Integer, ByVal BasePoint As Long, ByVal NodeLen As Long, ByVal Direc As Integer, ByVal ReverseGap As Integer, ByVal TEData As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisTECWork_9030
 ' 函数编号 : 111

 ' 描述   : 启动/停止 轴螺距误差补偿
 ' 参数   :
        
         ' Board_NO :    0－3, 板号
         ' Axis_No  :    0－3, 轴号
         ' Work     :    0或1            0=停止螺距误差补偿;1=启动螺距误差补偿
           

 ' 返回值 :  0或-1、-2：        -1=不成功,-2=轴螺距误差补偿数据失效，0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisTECWork_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal Work As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisFE_9030
 ' 函数编号 : 74

 ' 描述   : 设置轴跟随编码器运动
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
       '   Axis_No  : 0－3, 轴号
       '   Rate     : 轴跟随比率,可正负; 范围: 绝对值0.001-1000;  单位: 无;  分辨率: 0.001
       '   Kp       : 跟随PID调解系数; 范围: 0.001-1000;  单位: 无;  分辨率: 0.001
       '   Mode1    : 0=速度模式,1=位置模式
        '  Mode2    : 0= 不自动清零， 1=自动清零
        '  Mode3    : 0=停止状态跟随，1=可运动状态跟随
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisFE_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal Rate As Double, ByVal Kp As Double, ByVal Mode1 As Integer, ByVal Mode2 As Integer, ByVal Mode3 As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetAxisEGear_9030
 ' 函数编号 : 75

 ' 描述   : 设置轴电子齿轮运动
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号
       '   F_Axis_No: 0－3, 主动轴号

       '   Rate     : 轴跟随比率,可正负; 范围: 绝对值0.001-1000;  单位: 无;  分辨率: 0.001
       '   Kp       : 跟随PID调解系数; 范围: 0.001-1000;  单位: 无;  分辨率: 0.001
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetAxisEGear_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, _
ByVal F_Axis_No As Integer, ByVal Rate As Double, ByVal Kp As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : CancelAxisFEG_9030
 ' 函数编号 : 76

 ' 描述   : 取消轴跟随编码器运动
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号


 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function CancelAxisFEG_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ResetEn0Flag_9030
 ' 函数编号 : 79

 ' 描述   : 编码器自动清零标志置0
 ' 参数   :
        
         ' Board_NO : 0－3, 板号


 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ResetEn0Flag_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : GetEnAuto0Flag_9030
 ' 函数编号 : 80

 ' 描述   : 获得编码器自动清零标志
                        
 ' 参数   :
        
 '         Board_NO : 0－3, 板号
          
 ' 返回值 :
'0:                                编码器还没有被自动清零
'1:                                编码器自动清零
'255:                          命令不成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetEnAuto0Flag_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadAxisPos_9030
 ' 函数编号 : 2

 ' 描述   : 读取轴的当前位置
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
  

 ' 返回值 : 轴的位置; 范围: 长整型数;  单位: 脉冲个数。

         '   当返回轴位置=-2147483648 时,表示不成功,有错误产生。
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadAxisTheoryPos_9030
 ' 函数编号 : 2

 ' 描述   : 读取轴的当前理论位置
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
  

 ' 返回值 : 轴的位置; 范围: 长整型数;  单位: 脉冲个数。

            '当返回轴位置=-2147483648 时,表示不成功,有错误产生。
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisTheoryPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadAxisEncodePos_9030
 ' 函数编号 : 2

 ' 描述   : 读取轴的编码器位置
 ' 参数   :
        
          'Board_NO : 0－3, 板号
          'Axis_No  : 0－3, 轴号
  

 ' 返回值 : 轴的位置; 范围: 长整型数;  单位: 脉冲个数。

           ' 当返回轴位置=-2147483648 时,表示不成功,有错误产生。
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisEncodePos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadAxisTECV_9030
 ' 函数编号 : 2

 ' 描述   : 读取轴螺距误差补偿数据
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
  

 ' 返回值 : 轴的螺距误差补偿数据; 范围: 整型数;  单位: 脉冲个数。

           ' 当返回数据=-32768 时,表示不成功,有错误产生。
 ''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisTECV_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadAxisVel_9030
 ' 函数编号 : 26

 ' 描述   : 读取轴的当前速度
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号
          

 ' 返回值 : 轴的当前速度; 范围: 长整型数;    单位: Hz，轴的脉冲频率。

           ' 当返回=-2147483648 时,表示不成功,有错误产生。
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadAxisState_9030
 ' 函数编号 : 27

 ' 描述   : 读取轴的状态
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号
          

 ' 返回值 :
'1:                  轴在运动中
'0:                  轴在停止状态 , 轴位置到达, 或被清零 Lib "dfjzh9030dll.dll" (Home_9030命令)
'                -1: 轴GoHome OK
''                -2: 轴在GoHome中暂时停止
 '               -3: 轴被StopAxis_9030命令停止
'                -4: 轴被AbortAxis_9030命令停止
'                -5: 轴被CeaseAxis_9030命令停止
'                -6: 轴被 插补 命令停止
'                -7: 轴被正限位停止
'                -8: 轴被负限位停止
 '               -9: 轴位置寄存器溢出被停止
'           -10: 轴被外部IO急停停止
 '          -11: 轴在跟随模式下速度为0
'           -12: 轴编码器位置寄存器溢出
'           -13: 轴跟随误差超限
' '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadAxisState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadEncoderPos_9030
 ' 函数编号 : 10

 ' 描述   : 读取编码器位置
 ' 参数   :
        
         ' Board_NO : 0－3, 板号

 ' 返回值 : 编码器位置; 范围: 长整型数;  单位: 脉冲个数。
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadEncoderPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : HomeEncode_9030
 ' 函数编号 : 11

 ' 描述   : 复位编码器,编码器位置清零
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
            

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function HomeEncode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetEncodeCount_9030
 ' 函数编号 : 113

 ' 描述     :   设附加编码器初值

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
          
          'offset   : 编码器初值; 范围: 长整型数;  单位: 脉冲个数。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetEncodeCount_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal offset As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadFirmwareVersion_9030
 ' 函数编号 : 28

 ' 描述   : 读取9030控制卡固件版本号
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
            

 ' 返回值 :  0或 大于1, 0=不成功, 大于1=版本号, 比如:  10=1.0版本
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadFirmwareVersion_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : GetHWID_9030
 ' 函数编号 : 1020

 ' 描述   : 读取9030控制卡硬件ID号
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
            

 ' 返回值 :  0或 大于1, 0=不成功, 大于1=ID号, 57=9030,58=9011
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetHWID_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadDllVersion_9030
 ' 函数编号 : 67

 ' 描述   : 读取9030动态链接库版本号
 ' 参数   :
        
  '                       Lib "dfjzh9030dll.dll" (无)

 ' 返回值 :  0或 大于1, 0=不成功, 大于1=版本号, 比如:  10=1.0版本
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadDllVersion_9030 Lib "dfjzh9030dll.dll" () As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : GetDriverVersion_9030
 ' 函数编号 : 70

 ' 描述   : 读取9030卡的驱动版本号
 ' 参数   :
        
         '                Lib "dfjzh9030dll.dll" (无)
'
 ' 返回值 :  0或 大于1, 0=不成功, 大于1=版本号, 比如:  100=1.00版本,比如:  110=1.10版本
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetDriverVersion_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadIO_9030
 ' 函数编号 : 14

 ' 描述   : 读取通用输入点I1-I20状态
 ' 参数   :
        
       '   Board_NO : 0－3, 板号
            

 ' 返回值 :  0-19位有效,对应输入点I1-I20状态
 '                        -1=不成功,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadIO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadIOBit_9030
 ' 函数编号 : 14

 ' 描述   : 按位读取通用输入点I1-I20状态
 ' 参数   :
        
       '   Board_NO : 0－ 3,  板号
       '   Index    : 1 - 20, 输入点索引号

 ' 返回值 :  0-1,输入点状态
 '                        -1=不成功,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadIOBit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : WriteIo_9030
 ' 函数编号 : 16

 ' 描述   : 设置通用输出点O1-O8状态
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' IO_V     : 输出点值,低8位有效,对应O1-O8,8个输出点状态

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function WriteIo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_V As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : WriteIoBit_9030
 ' 函数编号 : 81

 ' 描述   : 按位设置通用输出点O1-O8状态
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
        '  IO_V     : 0 - 1，输出点值
        '  Index    : 1 - 8, 输出点索引号

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function WriteIoBit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_V As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadOs_9030
 ' 函数编号 : 82

 ' 描述   : 读取输出点O1-O8状态
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
            

 ' 返回值 :  0-7位有效,对应输入点O1-O8状态

         '                -1=不成功,
 ''
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadOs_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadOsBit_9030
 ' 函数编号 : 83

 ' 描述   : 按位读取输出点O1-O8状态
 ' 参数   :
        
       '   Board_NO : 0－3, 板号
       '   Index    : 1 - 8, 输出点索引号

 ' 返回值 :  0-1,输出点状态
        '                 -1=不成功,
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadOsBit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadMPGIO_9030
 ' 函数编号 : 118

 ' 描述   : 读手轮IO输入点MPG_I1-MPG_I7状态
 ' 参数   :
        
       '   Board_NO : 0－3, 板号
       '   Index    : 0=全读，返回值的bit0-bit6 对应MPG_I1-MPG_I7状态，1-7=按位读取，返回对应位状态

 ' 返回值 :  Index=0时 0-6位有效,对应输入点MPG_I1-MPG_I7状态，Index=1-7时，返回对应位状态 Lib "dfjzh9030dll.dll" (0或1）
       '                  -1=不成功,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadMPGIO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Index As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : GetAxisMode_9030
 ' 函数编号 : 119

 ' 描述   : 读轴用户设定模式
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号

 ' 返回值 :  0或1,对应轴的用户设定模式
         '                -1=不成功,
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetAxisMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : PwmOut_9030
 ' 函数编号 : 12

 ' 描述   : PWM脉冲输出
 
 ' 参数   :
        
        ' Board_NO :   0－3, 板号
        '  frequency:   18-1500000;    PWM输出频率; 单位: 赫兹 Lib "dfjzh9030dll.dll" (Hz)
        '  Pulse_Highf: 占空比,范围:0.0-1.0;  精度:不小于1%

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function PwmOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal frequency As Long, ByVal Pulse_Highf As Single) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   :PwmOut2_9030
 ' 函数编号 : 12

 ' 描述   : PWM脉冲输出
 
 ' 参数   :
        
        '  Board_NO   :   0－3, 板号
       '   frequency  :   18-1500000;    PWM输出频率; 单位: 赫兹 Lib "dfjzh9030dll.dll" (Hz)
        'Pulse_Highf:             高电平脉宽  Lib "dfjzh9030dll.dll" (ms)

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function PwmOut2_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal frequency As Long, ByVal Pulse_Highf As Single) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : PwmStop_9030
 ' 函数编号 : 13

 ' 描述   : PWM 脉冲停止输出

 ' 参数   :
        
          'Board_NO :   0－3, 板号

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function PwmStop_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : DAOut_9030
 ' 函数编号 : 68

 ' 描述   : DA Lib "dfjzh9030dll.dll" (数模转换)模拟量输出
 
 ' 参数   :
        
         ' Board_NO :   0－3, 板号
         ' DA_Avlue :   -10 - +10 V; 精度: 1/6000; 即: 精度大于12位 Lib "dfjzh9030dll.dll" (4096),不到13位 Lib "dfjzh9030dll.dll" (8192).
         '              该DA输出与PWM输出占用同一个硬件资源 , 通过9030端子板上的跳线选择是DA输出
         '                          还是PWM输出.

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function DAOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal DA_Avlue As Double) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetXAxis_9030
 ' 函数编号 : 32

 ' 描述   : 将实际轴与插补引擎的X轴相匹配

 ' 参数   :
        
          'Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号

'factor_c_t:             轴的脉冲当量?
'delta:                  该轴插补定位误差检查值?单位为用户单位?

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetXAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_OffXAxis_9030
 ' 函数编号 : 33

 ' 描述   : 撤消插补引擎的X轴匹配
 ' 参数   :
        
          'Board_NO : 0－3, 板号
          

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffXAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetYAxis_9030
 ' 函数编号 : 34

 ' 描述   : 将实际轴与插补引擎的Y轴相匹配

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号

'factor_c_t:             轴的脉冲当量?
'delta:                  该轴插补定位误差检查值?单位为用户单位?

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetYAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_OffYAxis_9030
 ' 函数编号 : 35

 ' 描述   : 撤消插补引擎的Y轴匹配
 ' 参数   :
        
         ' Board_NO : 0－3, 板号
          

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffYAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetZAxis_9030
 ' 函数编号 : 36

 ' 描述   : 将实际轴与插补引擎的Z轴相匹配

 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        '  Axis_No  : 0－3, 轴号

'factor_c_t:             轴的脉冲当量?
'delta:                  该轴插补定位误差检查值?单位为用户单位?

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetZAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_OffZAxis_9030
 ' 函数编号 : 37

 ' 描述   : 撤消插补引擎的Z轴匹配
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
          

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffZAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetWAxis_9030
 ' 函数编号 : 38

 ' 描述   : 将实际轴与插补引擎的W轴相匹配

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 0－3, 轴号

'factor_c_t:             轴的脉冲当量?
'delta:                  该轴插补定位误差检查值?单位为用户单位?

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetWAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal factor_c_t As Double, ByVal delta As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_OffWAxis_9030
 ' 函数编号 : 39

 ' 描述   : 撤消插补引擎的W轴匹配
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
          

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_OffWAxis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetAxisMaxErrLtd_9030
 ' 函数编号 : 87

 ' 描述   : 设轴最大插补位置误差限制

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
         ' Axis_No  : 1－4, 轴标志号;1=X轴,2=Y轴,3=Z轴,4=W轴
          
         ' ErrLid   :    该轴最大插补定位误差限制值。单位为用户单位。如果超过该值,系统将报警.

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetAxisMaxErrLtd_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal ErrLid As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_MicroAdjustPos_9030
 ' 函数编号 : 89

 ' 描述   :   插补暂停后微调轴位置

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
        '  Axis_No  : 1－4, 轴标志号;1=X轴,2=Y轴,3=Z轴,4=W轴
          
'MA_Pos:               该轴位置微调值?相对值 , 单位为用户单位?

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_MicroAdjustPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal MA_Pos As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetDecMagnifyCoeff_9030
 ' 函数编号 : 86

 ' 描述   : 设插补减速度放大系数

 ' 参数   :
        
         ' Board_NO : 0－3, 板号
          
         ' MagnifyCoeff: 1.0 - 2.0,  为了改善插补减速时的冲击。

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetDecMagnifyCoeff_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MagnifyCoeff As Double) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetACCDec_9030
 ' 函数编号 : 40

 ' 描述   : 设置插补加速度和减速度

 ' 参数   :
        
        '  Board_NO : 0－3, 板号

        ' acceleration:  插补加速度，范围：1-10000，单位：用户单位 / 秒/ 秒。缺省值: 500
        ' deceleration:  插补减速度，范围：1-10000，单位：用户单位 / 秒/ 秒。缺省值: 500

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetACCDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal acceleration As Long, ByVal deceleration As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Start_9030
 ' 函数编号 : 41

 ' 描述   : 插补开始

 ' 参数   :
        
         ' Board_NO : 0－3, 板号

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Start_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ObligeFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetXStartPos_9030
 ' 函数编号 : 54

 ' 描述   : 获得插补轴当前位置,也是轴插补的开始位置, 在发送插补命令 Lib "dfjzh9030dll.dll" (LM_Line_9030,IpolArc_6030)开始前调用

 ' 参数   :
        
        '  Board_NO : 0－3, 板号
        ' Pos:                 轴插补的开始位置 , 单位为用户单位?指针类参数, 可以为NULL Lib "dfjzh9030dll.dll" (空指针)

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetXStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer
Declare Function LM_GetYStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer
Declare Function LM_GetZStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer
Declare Function LM_GetWStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Pos As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetAxisStartPos_9030
 ' 函数编号 : 54

 ' 描述   : 获得插补轴当前位置,也是轴插补的开始位置,                     在发送插补命令 Lib "dfjzh9030dll.dll" (LM_Line_9030,IpolArc_6030)开始前调用

 ' 参数   :
        
          'Board_NO : 0－3, 板号
          'AxisFlag : 1- 4, 指示是哪个轴；1：X轴， 2：Y轴，3：Z轴，4：W轴。

 ' 返回值 :  轴的当前位置。用户单位
           '              当返回轴位置=-2147483648 时,表示不成功,有错误产生。
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetAxisStartPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal AxisFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Line_9030
 ' 函数编号 : 42

 ' 描述   : 直线插补
        '                该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行
'
 ' 参数   :
        
      '    Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
 '         Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Line_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_LineMaxV_9030
 ' 函数编号 : 69

 ' 描述   : 直线插补
  '                      该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
  '        Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'          MaxSpeed : 本行插补最大速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineMaxV_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal Speed As Double, _
ByVal MaxSpeed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_LineMeasure_9030
 ' 函数编号 : 65

 ' 描述   : 直线插补及测量IO点输入
   '                     该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          IO_Index : 1-8;  通用IO点输入点序号,
'          Mode     : 1,2,3,4，模式
'                                                                1,2:通用IO点输入点为1， 1=立即停止，2=10ms减速停止
'                                                                3,4:通用IO点输入点为0， 3=立即停止，4=10ms减速停止
          
'          Speed    : 插补速度,也是本行插补最大速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineMeasure_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal IO_Index As Integer, ByVal Mode As Integer, _
ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_LineFE_9030
 ' 函数编号 : 77

 ' 描述   : 直线插补 跟随编码器位置
 '                       该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
  '        Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'
'          Mode     : 0,1; 0=到达目标,立即停止跟随; 1=到达目标,不停止跟随,有后续命令决定;
'          Rate     : 轴跟随比率,可正负;
                                 
'                                 范围: 由跟随轴的脉冲当量决定, 绝对值范围: 0.001'跟随轴的脉冲当量 - 1000 '跟随轴的脉冲当量 ;
'                                 单位: 无;  分辨率: 0.001'跟随轴的脉冲当量
'
'          Kp       : 跟随PID调解系数; 范围: 0.001-1000;  单位: 无;  分辨率: 0.001
'          Speed    : 插补预期速度,也是本行插补最大速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineFE_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal Rate As Double, ByVal Kp As Double, ByVal Mode As Integer, ByVal Speed As Double, ByVal LineNO As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_ArcCW_9030
 ' 函数编号 : 43

 ' 描述   : 圆弧插补,顺圆
'                        该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'xcPos:               x轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'ycPos:               y轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定

'                                当z轴或w轴的起点与终点坐标不重合时,z轴或w轴随xy轴做螺旋线运动,Z和w可同时做螺旋线运动.
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCW_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal xcPos As Double, ByVal ycPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_ArcCW_ZX_9030
 ' 函数编号 : 43

 ' 描述   : 圆弧插补,顺圆,在ZX平面
  '                      该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zcPos:               z轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'xcPos:               x轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定

'                                当y轴或w轴的起点与终点坐标不重合时,y轴或w轴随zx轴做螺旋线运动,y和w可同时做螺旋线运动.
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCW_ZX_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal zcPos As Double, ByVal xcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_ArcCW_YZ_9030
 ' 函数编号 : 43

 ' 描述   : 圆弧插补,顺圆,在YZ平面
   '                     该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'ycPos:               y轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zcPos:               z轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定

'                                当x轴或w轴的起点与终点坐标不重合时,x轴或w轴随yz轴做螺旋线运动,x和w可同时做螺旋线运动.
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCW_YZ_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal ycPos As Double, ByVal zcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_ArcCCW_9030
 ' 函数编号 : 44

 ' 描述   : 圆弧插补,逆圆
 '                       该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'xcPos:               x轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'ycPos:               y轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定

 '                               当z轴或w轴的起点与终点坐标不重合时,z轴或w轴随xy轴做螺旋线运动,Z和w可同时做螺旋线运动.
'
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCCW_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal xcPos As Double, ByVal ycPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_ArcCCW_ZX_9030
 ' 函数编号 : 44

 ' 描述   : 圆弧插补,逆圆,在ZX平面
 '                       该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zcPos:               z轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'xcPos:               x轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定

'                                当y轴或w轴的起点与终点坐标不重合时,y轴或w轴随xy轴做螺旋线运动,y和w可同时做螺旋线运动.
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCCW_ZX_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, _
ByVal zcPos As Double, ByVal xcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_ArcCCW_YZ_9030
 ' 函数编号 : 44

 ' 描述   : 圆弧插补,逆圆,在YZ平面
 '                       该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行

 ' 参数   :
 '
 '         Board_NO : 0－3, 板号

'xPos:                x轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'ycPos:               y轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zcPos:               z轴的圆心位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
 '         Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定

'                                当x轴或w轴的起点与终点坐标不重合时,x轴或w轴随xy轴做螺旋线运动,x和w可同时做螺旋线运动.
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_ArcCCW_YZ_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, _
ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double, ByVal wPos As Double, ByVal ycPos As Double, ByVal zcPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_End_9030
 ' 函数编号 : 45

 ' 描述   : 插补结束
           '             该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行
          '              当插补引擎运行到该行时将停止运行 , 如要重新开始, 则再送入插补数据命令
          '               Lib "dfjzh9030dll.dll" (LM_Line_9030,LM_ArcCW_9030,LM_ArcCCW_9030),并发LM_Start_9030命令

 ' 参数   :
        
       '   Board_NO : 0－3, 板号
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_End_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Wait_9030
 ' 函数编号 : 60

 ' 描述   : 插补等待
           '             该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行
           '             当插补引擎运行到该行时将暂停运行,在等待用户所设时间后,再重新开始.

 ' 参数   :
        
      '    Board_NO    : 0－3, 板号
'Millisecond:            等待时间 , 毫秒
'LineNO:                 插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Wait_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Millisecond As Long, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_PWM_9030
 ' 函数编号 : 61

 ' 描述   : 插补模式的PWM输出占空比随速度变化
  '                      该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行
  '                      当插补引擎运行到该行时将根据参数决定PWM输出.

 ' 参数   :
        
 '         Board_NO    : 0－3, 板号
  '        frequency   : 18-1500000;    PWM输出频率; 单位: 赫兹 Lib "dfjzh9030dll.dll" (Hz)
 '                       当为0时 , 停止PWM输出
 '         Pulse_Highf : 占空比,范围:0.0-1.0;  精度:不小于1%
  '                      该值对应插补速度达到Speed所设速度时的最大占空比 , 当插补速度大于
 '                                       Speed所设速度时,也按该参数输出.
'
 '         Speed       : 占空比随动的最大插补速度,单位: 用户单位/分钟
  '                      当该值为0时 , 表明不是随动模式, 当插补引擎运行到该行时
  '                                      将根据frequency,Pulse_Highf参数值立即PWM输出.
'LineNO:                 插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_PWM_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal frequency As Long, ByVal Pulse_Highf As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_IOOut_9030
 ' 函数编号 : 62

 ' 描述   : 插补模式的IO点输出
     '                   该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行
   '                     当插补引擎运行到该行时将根据参数输出IO点.

 ' 参数   :
        
   '       Board_NO    : 0－3, 板号
    '      IO_Index    : 1-8;  通用IO点输出点序号,
     '     IO_Value    : 0 - 1; 输出点值
'LineNO:                 插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_IOOut_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_Index As Integer, ByVal IO_Value As Integer, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Wait_I_9030
 ' 函数编号 : 63

 ' 描述   : 插补等待通用IO点输入
   '                     该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有16行
  '                      当插补引擎运行到该行时将等待直到输入点为1.

 ' 参数   :
        
'          Board_NO    : 0－3, 板号
'          IO_Index    : 1-8;  通用IO点输入点序号,
'LineNO:                 插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Wait_I_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal IO_Index As Integer, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_CleanBuff_9030
 ' 函数编号 : 46

 ' 描述   : 清除插补缓存
                        
 ' 参数   :
        
 '         Board_NO : 0－3, 板号
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_CleanBuff_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetMinVel_9030
 ' 函数编号 : 47

 ' 描述   : 设置插补最小速度
                        
 ' 参数   :
        
 '         Board_NO  : 0－3, 板号
 '         MinLineVel: 直线最小插补速度, 单位: 用户单位/分钟, 缺省值: 30
 '         MinArcVel : 圆弧最小插补速度, 单位: 用户单位/分钟, 缺省值: 30
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetMinVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MinLineVel As Double, ByVal MinArcVel As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetMaxVel_9030
 ' 函数编号 : 57

 ' 描述   : 设置插补最大速度
                        
 ' 参数   :
        
       '   Board_NO  : 0－3, 板号
       '   MaxLineVel: 直线最大插补速度, 单位: 用户单位/分钟, 缺省值: 4000
       '   MaxArcVel : 圆弧最大插补速度, 单位: 用户单位/分钟, 缺省值: 4000
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetMaxVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MaxLineVel As Double, ByVal MaxArcVel As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetSpeedPri_9030
 ' 函数编号 : 58

 ' 描述   : 设置插补速度速度优先还是精度优先
                        
 ' 参数   :
        
'          Board_NO    : 0－3, 板号
 '         PriorityFlag: 0-1; 1:插补速度优先, 0:插补精度优先。缺省值: 1 插补速度优先
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSpeedPri_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal PriorityFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetSAccePower_9030
 ' 函数编号 : 71

 ' 描述   : 设置插补S型加速度 Lib "dfjzh9030dll.dll" (1,2,3,4)指数
                        
 ' 参数   :
        
'          Board_NO    : 0－3, 板号
  '        PowerFlag   : 1-4;  S型加减速 Lib "dfjzh9030dll.dll" (指数曲线)的指数
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSAccePower_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal PowerFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetParaAngle_9030
 ' 函数编号 : 48

 ' 描述   : 设插补参数,角度
                        
 ' 参数   :
        
  ''        Board_NO : 0－3, 板号
'angle1:
'angle2:
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetParaAngle_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal angle1 As Integer, ByVal angle2 As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetSysPara_9030
 ' 函数编号 : 59

 ' 描述   : 设插补系统参数
                        
 ' 参数   :
  '
  '        Board_NO : 0－3, 板号
 '         MinLength: 系统最小长度,单位: 用户单位, 缺省值: 0.001;在直线或圆弧插补时,直线或圆弧长度
 '                    不能小于该设定值
 '         MinSpeed : 系统最小速度,单位: 用户单位/分钟, 缺省值: 0.001;在直线或圆弧插补时,直线或圆弧
 '                                插补速度不能小于该设定值
  '        ArcError : 系统最大圆弧误差,单位: 用户单位, 缺省值: 0.2;在圆弧插补时,圆弧起点和终点半径
 '                    误差不能大于该设定值
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSysPara_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal MinLength As Double, ByVal MinSpeed As Double, ByVal ArcError As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetBuffLen_9030
 ' 函数编号 : 55

 ' 描述   : 获得插补缓存剩余长度
                        
 ' 参数   :
        
 '         Board_NO : 0－3, 板号
          
 ' 返回值 :  0-32,插补缓存长度
 '            -1=不成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetBuffLen_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetState_9030
 ' 函数编号 : 56

 ' 描述   : 获得插补状态
                        
 ' 参数   :
        
 '         Board_NO : 0－3, 板号
          
 ' 返回值 :
'0:                                停止状态
'1:                                被轴停止命令结束  Lib "dfjzh9030dll.dll" (建议用Abort_9030命令)
'2:                                被行结束停止命令停止
'3:                                被LM_End_9030结束
'4:                                插补缓冲区空
'5:                                被进给倍率为0暂停
'6:                                被进给暂停挂起
'7:                                插补正在进行
'255:                  命令不成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetMeasureState_9030
 ' 函数编号 : 66

 ' 描述   : 获得插补测量状态
                        
 ' 参数   :
        
'          Board_NO : 0－3, 板号
          
 ' 返回值 :
'0:                                没有检测到测量信号
'1:                                检测到测量信号
'255:                          命令不成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetMeasureState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetFEnState_9030
 ' 函数编号 : 78

 ' 描述   : 获得轴编码器跟随已达到目标状态
                        
 ' 参数   :
        
'          Board_NO : 0－3, 板号
          
 ' 返回值 :
'0:                                跟随还没有达到目标
'1:                                跟随已达到目标
'255:                          命令不成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetFEnState_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetLineNO_9030
 ' 函数编号 : 49

 ' 描述   : 获得插补当前行号
 ' 参数   :
        
 '         Board_NO : 0－3, 板号

 ' 返回值 : 当前行号; 范围: 长整型数;  用户设定值。

 '           当返回值=-2147483648 时,表示不成功,有错误产生。
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetLineNO_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Long


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Pause_9030
 ' 函数编号 : 50

 ' 描述   :  插补暂停,插补各轴在10毫秒之内减速停止
                        
 ' 参数   :
        
'          Board_NO : 0－3, 板号
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Pause_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Resume_9030
 ' 函数编号 : 51

 ' 描述   :  恢复插补暂停,插补各轴以插补加速度恢复运行
                        
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Resume_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetSpeedRate_9030
 ' 函数编号 : 52

 ' 描述   :  设插补速率
                        
 ' 参数   :
        
      '    Board_NO : 0－3,  板号
      '    Rate     : 1-160, 100=100%,即:按原设定速度执行;10=10%,即:按原设定速度的百分之十执行,
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetSpeedRate_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Rate As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_LineEnd_9030
 ' 函数编号 : 53

 ' 描述   : 插补运行完当前行停止
       '                 该命令不进入插补缓存区 , 可在插补运行时, 随时执行
     '                   当插补引擎运行完当前行停止时 , 如要重新开始, 则再执行LM_Start_9030命令

 ' 参数   :
        
   '      Board_NO : 0－3, 板号
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_LineEnd_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetForceCtrl_9030
 ' 函数编号 : 94

 ' 描述   : 设置插补力控制模式,固件3.0版
                        
 ' 参数   :
        
       '   Board_NO    : 0－3, 板号
      '    ForceFlag   : 0-1;  0:插补力控制无效,1=插补力控制有效
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetForceCtrl_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ForceFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetNurbsScanMode_9030
 ' 函数编号 : 1000

 ' 描述   : 设置Nurbs曲线插补预处理的扫描模式,固件3.0版

 ' 参数   :
        
        '  Board_NO              : 0－3, 板号

       '   Mode                  :       Nurbs曲线插补预先扫描速度模式: 0=给定速度扫描,1=按程序速度百分比扫描
'ScanSpeed:                              给定的扫描速度
      '    ScanSpeedRate :       按程序速度百分比的比值: 范围: 1-100
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsScanMode_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Mode As Integer, ByVal ScanSpeed As Double, ByVal ScanSpeedRate As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetNurbsVelCtrl_9030
 ' 函数编号 : 1001

 ' 描述   : 设置Nurbs曲线插补时的速度控制,固件3.0版

 ' 参数   :
        
      '    Board_NO              : 0－3, 板号

     '     BSErrEnable   :       0=控制弦高误差无效,1=控制弦高误差有效
     '     BSErrV                :       最大弦高误差值,范围: 0.0001-10.0, 单位: 用户单位
  
      '    RAccEnable    :       0=控制法向加速度无效,1=控制法向加速度有效
     '     RAccV                 :       最大法向加速度值,范围: 1-1000000,       单位: 用户单位/秒/秒
 

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsVelCtrl_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal BSErrEnable As Integer, ByVal BSErrV As Double, ByVal RAccEnable As Integer, ByVal RAccV As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetNurbsAccDec_9030
 ' 函数编号 : 93

 ' 描述   : 设置Nurbs曲线插补加速度和减速度,固件3.0版

 ' 参数   :
        
       '   Board_NO : 0－3, 板号

     '     Nurbs_Acc:    插补加速度，范围：1-10000，单位：用户单位 / 秒/ 秒。缺省值: 500
      '    Nurbs_Dec:    插补减速度，范围：1-10000，单位：用户单位 / 秒/ 秒。缺省值: 500

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsAccDec_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Nurbs_Acc As Double, ByVal Nurbs_Dec As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetNurbsCompCoef_9030
 ' 函数编号 : 97

 ' 描述   : 设Nurbs曲线插补误差补偿系数  固件3.0版

 ' 参数   :
        
      '    Board_NO : 0－3, 板号

     '     Coef: 设Nurbs曲线插补误差补偿系数，范围：0.0-0.05，单位：无。缺省值: 0.01

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetNurbsCompCoef_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Coef As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Nurbs_9030
 ' 函数编号 : 92

 ' 描述   : Nurbs曲线插补,固件3.0版
  '                      该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有64行

 ' 参数   :
        
     '     Board_NO : 0－3, 板号

'knot1:               节点值1?
'knot2:               节点值2?
'knot3:               节点值3?
'knot4:               节点值4?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Nurbs_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal knot1 As Double, ByVal knot2 As Double, ByVal knot3 As Double, ByVal knot4 As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_Nurbs4Axis_9030
 ' 函数编号 : 92

 ' 描述   : 4轴Nurbs曲线插补,固件3.0版
      '                 该命令将把插补数据送入9030卡的插补缓存器中,9030卡的插补缓存器总共有64行

 ' 参数   :
        
 '         Board_NO : 0－3, 板号

'knot1:               节点值1?
'knot2:               节点值2?
'knot3:               节点值3?
'knot4:               节点值4?
'wPos:                w轴插补的终点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'          Speed    : 插补速度,单位: 用户单位/分钟
'LineNO:              插补行号 , 用户自由设定
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_Nurbs4Axis_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal knot1 As Double, ByVal knot2 As Double, ByVal knot3 As Double, ByVal knot4 As Double, ByVal wPos As Double, ByVal Speed As Double, ByVal LineNO As Long) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_NurbsData_9030
 ' 函数编号 : 91

 ' 描述   : 设Nurbs曲线插补数据,固件3.0版

 ' 参数   :
        
'          Board_NO : 0－3, 板号

'xPos:                X轴的控制点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴的控制点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴的控制点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'knot:                节点值?
'weight:              权值?
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_NurbsData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal knot As Double, ByVal weight As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_NurbsInit_9030
 ' 函数编号 : 90

 ' 描述     : 初始化Nurbs曲线,为Nurbs曲线插补做准备

 ' 参数   :
        
       '   Board_NO : 0－3, 板号

       '   _deg     :  Lib "dfjzh9030dll.dll" (保留参数,恒等于3) 3=三次Nurbs曲线
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_NurbsInit_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal deg As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SendNurbsData_9030
 ' 函数编号 : 1013

 ' 描述   :   向9030卡转输NURBS曲线数据,固件3.0版
                        
 ' 参数   :
        
 '         Board_NO    : 0－3, 板号
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SendNurbsData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetNubrsExecPara_9030
 ' 函数编号 : 1012

 ' 描述   : 读取9030卡NURBS曲线插补运行参数值
 ' 参数   :
        
       '   Board_NO : 0－3, 板号
            

 ' 返回值 :  当前9030卡NURBS曲线插补运行参数值
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetNubrsExecPara_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Single

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetFactVel_9030
 ' 函数编号 : 1002

 ' 描述   :       获得实际插补速度                      动态连接库3.0版
 ' 参数   :
        
        '  Board_NO : 0－3, 板号
            

 ' 返回值 :  实际插补速度       单位: 用户单位/分钟
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetFactVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Single

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetNurbsInBuffLen_9030
 ' 函数编号 : 1003

 ' 描述   : 获得Nurbs曲线在插补缓冲区的数               动态连接库3.0版
                        
 ' 参数   :
        
       ' Board_NO : 0－3, 板号
       '
 ' 返回值 :  0-64,Nurbs曲线在插补缓冲区的数
       '      -1=不成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetNurbsInBuffLen_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Set_NurbsInit_9030
 ' 函数编号 : 1004

 ' 描述     : 初始化Nurbs曲线,为Nurbs曲线计算做准备

 ' 参数   :
        
        '  Nurbs_NO : 0－7, Nurbs曲线号

       '   _deg     :  Lib "dfjzh9030dll.dll" (保留参数,恒等于3) 3=三次Nurbs曲线
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_NurbsInit_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal deg As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Set_NurbsData_9030
 ' 函数编号 : 1005

 ' 描述   : 设Nurbs曲线计算数据,动态连接库3.0版

 ' 参数   :
        
     '     Nurbs_NO : 0－7, Nurbs曲线号

'xPos:                X轴的控制点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'yPos:                y轴的控制点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'zPos:                z轴的控制点位置  Lib "dfjzh9030dll.dll" (绝对值), 单位为用户单位?
'knot:                节点值?
'weight:              权值?
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_NurbsData_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal xPos As Double, _
ByVal yPos As Double, ByVal zPos As Double, ByVal knot As Double, ByVal weight As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Set_NurbsEnd_9030
 ' 函数编号 : 1006

 ' 描述   : 设Nurbs曲线数据结束,动态连接库3.0版

 ' 参数   :
        
    '     Nurbs_NO : 0－7, Nurbs曲线号

'knot1:               节点值1?
'knot2:               节点值2?
'knot3:               节点值3?
'knot4:               节点值4?
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Set_NurbsEnd_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, _
ByVal knot1 As Double, ByVal knot2 As Double, ByVal knot3 As Double, ByVal knot4 As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Get_NurbsPos_9030
 ' 函数编号 : 1007

 ' 描述   : 计算Nurbs曲线轴位置 Lib "dfjzh9030dll.dll" (型值点)坐标,动态连接库3.0版

 ' 参数   :
        
    '      Nurbs_NO : 0－7, Nurbs曲线号

    '      Up       : Nurbs曲线参数变量,范围: 0-  Lib "dfjzh9030dll.dll" (控制点数-3)。
'xPos:                X轴位置指针?
'yPos:                Y轴位置指针?
'zPos:                Z轴位置指针?
          
 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsPos_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal Up As Double, ByVal xPos As Double, ByVal yPos As Double, ByVal zPos As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Get_NurbsPosVB_9030
 ' 函数编号 : 1007

 ' 描述   : 计算Nurbs曲线轴位置 Lib "dfjzh9030dll.dll" (型值点)坐标,动态连接库3.0版

 ' 参数   :
        
      '    Nurbs_NO : 0－7, Nurbs曲线号

    '      Up       : Nurbs曲线参数变量,范围: 0-  Lib "dfjzh9030dll.dll" (控制点数-3)。
'xflag:               为1时 , 指示返回X轴位置?
'yflag:               为1时 , 指示返回Y轴位置?
'zflag:               为1时 , 指示返回Z轴位置?
          
 ' 返回值 :  Nurbs曲线在参数变量为Up的轴位置
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsPosVB_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal Up As Double, ByVal xflag As Integer, ByVal yflag As Integer, ByVal zflag As Integer) As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Get_NurbsLen_9030
 ' 函数编号 : 1009

 ' 描述   : 计算Nurbs曲线长度,动态连接库3.0版

 ' 参数   :
        
      '    Nurbs_NO : 0－7, Nurbs曲线号

       '   Up       : Nurbs曲线参数变量,范围: 0-  Lib "dfjzh9030dll.dll" (控制点数-3)。

          
 ' 返回值 :  Nurbs曲线在参数变量为Up的长度
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsLen_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer, ByVal Up As Double) As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : Get_NurbsErrorNo_9030
 ' 函数编号 : 1011

 ' 描述   : 获得产生Nurbs曲线 Lib "dfjzh9030dll.dll" (计算)的错误号             动态连接库3.0版
                        
 ' 参数   :
        
    '      Nurbs_NO : 0－7, Nurbs曲线号
          
 ' 返回值 :
'Nurbs错误号:
'
'                        1=内存不够
'                        2=超出计算范围
'                        3=计算逻辑错误
'                        4=扫描速度过快
''                        5=数据不完整
'                        6=节点矢量的数据超过数学公式的定义
'                        7=控制点数不能少于4
'                        8=参数变量超过范围
 '                       9=电机控制卡初始化不成功
'                        -1=不成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function Get_NurbsErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Nurbs_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_GetNurbsErrorNo_9030
 ' 函数编号 : 1010

 ' 描述   : 获得产生Nurbs曲线 Lib "dfjzh9030dll.dll" (插补)的错误号             动态连接库3.0版
                        
 ' 参数   :
        
 '         Board_NO : 0－3, 板号
          
 ' 返回值 :
'Nurbs错误号:

'                        1=内存不够
'                        2=超出计算范围
'                        3=计算逻辑错误
'                        4=扫描速度过快
'                        5=数据不完整
 '                       6=节点矢量的数据超过数学公式的定义
''                        7=控制点数不能少于4
 '                       8=参数变量超过范围
'                        -1=不成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_GetNurbsErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : GetErrorNo_9030
 ' 函数编号 : 100

 ' 描述   : 获得错误号
 '                       当调用函数的返回值 为"不成功"时,或电机卡不正常时,
 '                       可调用该函数获得错误信息
''
 ' 参数   :
        
'          Board_NO  : 0－3, 板号
'          CleanFlag : 0或1;  0=不清除错误;  1=清除错误;
 '         'ErrorNo  : 返回错误编号                指针类参数,可以为NULL Lib "dfjzh9030dll.dll" (空指针)
 '         'FuncNo   : 返回产生错误的函数编号      指针类参数,可以为NULL Lib "dfjzh9030dll.dll" (空指针)
          
 ' 返回值 :  0或1, 0=无错误,1=有错误
 '
 ' 错误编号表:

'1                   轴位置寄存器超限
'2                   轴在插补运动中
'3                   插补误差超限
'4                   轴在运动中
'5                   逻辑错误
'6                   插补轴压在限位上

'7                   内存不够
'8                   参数错误
'9                   插补缓存器满
'10              与电机卡通讯失败
'11              与电机卡通讯超时
'12              插补数据小于系统最小速度
'13              插补数据小于系统最小长度
'14              插补数据大于系统圆弧半径误差
'15              插补数据圆弧计算错误
'           16   大于系统最大值或小于系统最小值:  系统值范围:  轴位置:       -2147483648 至 2147483647
'                                                                                                                  插补向量长度: <1073741823      Lib "dfjzh9030dll.dll" (直线长度和圆弧长度)
'                                                                                                                  圆弧最大半径: <1048575
'                                                                                                                  插补圆心位置: -2147483648 至 2147483647
'           17   9030动态链接库版本与9030卡的固件版本不一致
'18              有插补测量代码 , 后续不能再加入插补命令
'19              轴在跟随运动
'20              轴在停止过程
'21                  Nurbs曲线计算所需内存不够
'22                  Nurbs曲线超出计算范围
'23                  Nurbs曲线计算逻辑错误
'24                  Nurbs曲线扫描速度过快
'25                  Nurbs曲线数据不完整
'26                  Nurbs曲线的起点不连续
'27                  Nurbs曲线的数据没有及时传入卡内
'
'
'28                  多块90300卡的固件版本不一致
'                29  9030卡固件版本太低

'31                              接收命令错误
'32                              接收插补数据错误
'33                              接收NURBS数据错误
'34                              当轴是闭环模式时 , 不能单独操作位置编码器
'35                              轴编码器反馈位置寄存器超限
'36                              轴位置跟随误差超限
'37                              轴找一转脉冲失败
'38                              轴Home点无效 , GoHome失效


 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal CleanFlag As Integer, ByVal ErrorNo As Integer, ByVal FuncNo As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : DM_SetAxisVel_9030
 ' 函数编号 : 21

 ' 描述 :轴直接速度输出模式
 '  input       :
 '                               num          轴号1－4
 '                               velocity      轴速度模式
 '
 '  output      :none
 '  return      :none
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function DM_SetAxisVel_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal position As Long, ByVal velocity As Long, ByVal direction As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : DM_SetAxisPos_9030
 ' 函数编号 : 22

 ' 描述 :轴直接速度输出模式
 '  input       :
 '                               num          轴号1－4
 '                               velocity      轴速度模式
 '
 '  output      :none
 '  return      :none
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function DM_SetAxisPos_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer, ByVal position As Long, ByVal velocity As Long, ByVal direction As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : RegCANExp_9030
 ' 函数编号 : 2000

 ' 描述     :登记CAN总线扩展卡
 ' 参数:
        
 '         Board_NO: 0-3,板号
'ID:                 CAN总线扩展卡ID号
'CardType:           扩展卡型号


 ' 返回值:  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function RegCANExp_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long, ByVal CardType As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : EnableCANExp_9030
 ' 函数编号 : 116

 ' 描述     : CAN总线扩展卡使能 Lib "dfjzh9030dll.dll" (初始化邮箱)
 ' 参数:
        
      '    Board_NO: 0-3,板号

 ' 返回值:  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function EnableCANExp_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SendCANData_9030
 ' 函数编号 : 117

 ' 描述     : CAN总线扩展卡发送数据  Lib "dfjzh9030dll.dll" (邮箱发送消息)
 ' 参数:
        
'          Board_NO: 0-3,板号
'ID:                 扩展卡的ID号
'CardType:           扩展卡型号
'D1234:              低4字节
'D5678:              高4字节

 ' 返回值:  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SendCANData_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long, ByVal CardType As Integer, ByVal D1234 As Long, ByVal D5678 As Long) As Integer


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadCANL_9030
 ' 函数编号 : 2001

 ' 描述   : 读取CAN总线扩展卡数据 低4位
 ' 参数   :
        
  '        Board_NO : 0－3, 板号
'ID:                 扩展卡的ID号
            

 ' 返回值 :  0-7、8-15、16-23、24-31位,分别对应第1、2、3、4字节
                         
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadCANL_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : ReadCANH_9030
 ' 函数编号 : 2001

 ' 描述   : 读取CAN总线扩展卡数据 高4位
 ' 参数   :
        
  '        Board_NO : 0－3, 板号
'ID:                 扩展卡的ID号
            

 ' 返回值 :  0-7、8-15、16-23、24-31位,分别对应第5、6、7、8字节
                         
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function ReadCANH_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal ID As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : GetCANErrorNo_9030
 ' 函数编号 : 2002

 ' 描述   : 获得CAN总线错误号
  '                      当调用函数的返回值 为"不成功"时,或电机卡不正常时,
  '                      可调用该函数获得错误信息

 ' 参数   :
        
  '        Board_NO  : 0－3, 板号
  '        CleanFlag : 0或1;  0=不清除错误;  1=清除错误;

          
 ' 返回值 :  0-9, 0=无错误,错误号
 '
 ' 错误编号表:

'1                                       邮箱号越界
'2                                       数据长度越界
'3                                       邮箱已经占用
'4                                       邮箱的ID号重复
'5                                       CAN总线只能使能一次
'6                                       邮箱不匹配
'                7                       CAN总线接受数据超过限制  Lib "dfjzh9030dll.dll"  Lib "dfjzh9030dll.dll" (>64字节)
'8                                       邮箱已满
'9                                       CAN 总线接收数据无效

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function GetCANErrorNo_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal CleanFlag As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetLineArcCompCoef_9030
 ' 函数编号 : 120

 ' 描述   : 设直线/圆弧插补误差补偿系数  固件3.7版以上

 ' 参数   :
        
   '       Board_NO : 0－3, 板号
   '       flag     : 0-1   0=直线插补,1=圆弧插补
   '       Coef: 设直线/圆弧插补误差补偿系数，直线插补范围：0.0,0.03-0.08，单位：无。缺省值: 0.00
    '                                                                             圆弧插补范围：0.0,0.03，         单位：无。缺省值: 0.00

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function LM_SetLineArcCompCoef_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal flag As Integer, ByVal Coef As Double) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : SetActiveEncoder_9030
 ' 函数编号 : 121

 ' 描述   : 在轴跟随编码器运动之前,设主动编码器号
 ' 参数   :
        
 '         Board_NO : 0－3, 板号
 '         Axis_No  : 0－4, 0-3对应轴号,4=附加编码器,缺省为4
           

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
Declare Function SetActiveEncoder_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal Axis_No As Integer) As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 '
 ' 函数名   : LM_SetIpolCompCoef_9030
 ' 函数编号 : 120

 ' 描述   : 设插补误差补偿系数  固件3.6版以上

 ' 参数   :
    
 '     Board_NO : 0－3, 板号
 '     flag1    : 0-3   0=直线插补,1=直线插补,2=圆弧插补,3=Nurbs曲线插补
 '     flag2    : 0-2   0=直线插补,1=直线插补,2=圆弧插补,3=Nurbs曲线插补
 '     Coef: 设直线/圆弧插补误差补偿系数，直线插补范围：0.0,0.03-0.08，单位：无。缺省值: 0.00
 '                                        圆弧插补范围：0.0,0.03，     单位：无。缺省值: 0.00

 ' 返回值 :  0或-1,-1=不成功,0=成功
 '
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/

Declare Function LM_SetIpolCompCoef_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal flag1 As Integer, ByVal flag2 As Integer, ByVal Coef As Double) As Integer

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/

' 函数名   : LM_SetSysDelay_9030
' 函数编号 :

' 描述   : 设插补系统参数
            
' 参数   :
    
'      Board_NO : 0－3, 板号
'      DelayFlag: 0,1,  0=不延时; 1=延时1毫秒  读插补缓存剩余长度时延时标志.  缺省值: 1;
      
      
' 返回值 :  0或-1,-1=不成功,0=成功
'
 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Declare Function LM_SetSysDelay_9030 Lib "dfjzh9030dll.dll" (ByVal Board_NO As Integer, ByVal DelayFlag As Integer) As Integer

 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''/
' * 函数名   : UnlockFlash_9030
' * 函数编号 : 126

' * 描述   : 9030卡上 用户flash 解锁
' * 参数   :
    
'      Board_NO : 0－3, 板号
'      password1: 长整型,解锁密码，缺省为0
'      password2: 长整型,解锁密码，缺省为0
       

' * 返回值 :  1，0或-1,0=不成功,1=成功，-1=操作失败
' *
' *
' *****************************************************************************/
'short APIENTRY UnlockFlash_9030(unsigned short Board_NO,unsigned long password1,unsigned long password2);

'/******************************************************************************
' *
' * 函数名   : LockFlash_9030
' * 函数编号 : 1023

' * 描述   : 9030卡上 用户flash 锁定
' * 参数   :
    
'      Board_NO : 0－3, 板号
'password1:       长整型 , 锁定密码
'password2:       长整型 , 锁定密码
       

' * 返回值 :  1，0或-1,0=不成功,1=成功，-1=操作失败
' *
' *
' *****************************************************************************/
'short APIENTRY LockFlash_9030(unsigned short Board_NO,unsigned long password1,unsigned long password2);

'/******************************************************************************
' *
' * 函数名   : WriteFlash_9030
' * 函数编号 : 128

' * 描述   : 写 9030卡上 用户flash
' * 参数   :
    
'      Board_NO : 0－3, 板号
'      offset    : 0-199 偏移值
'      len       : 1-8
'word1:       长整型 , 锁定密码
'word2:       长整型 , 锁定密码
       

' * 返回值 :  1，0或-1,0=不成功,1=成功，-1=操作失败
' *
' *
' *****************************************************************************/
'short APIENTRY WriteFlash_9030(unsigned short Board_NO,unsigned short offset,unsigned short len,unsigned long word1,unsigned long word2);

'/******************************************************************************
' *
' * 函数名   : WriteFlashChar_9030
' * 函数编号 : 128

' * 描述   : 写 9030卡上 用户flash
' * 参数   :
    
'      Board_NO : 0－3, 板号
'      offset    : 0-199 偏移值
'      len       : 1-8
'Data:             数据指针
       

' * 返回值 :  1，0或-1,0=不成功,1=成功，-1=操作失败
' *
' *
' *****************************************************************************/
'short APIENTRY WriteFlashChar_9030(unsigned short Board_NO,unsigned short offset,unsigned short len,unsigned char *Data);

'/******************************************************************************
' *
' * 函数名   : UpdateFlash_9030
' * 函数编号 : 129

' * 描述   : 更新 9030卡上 用户flash 内容
' * 参数   :
    
'      Board_NO : 0－3, 板号
     
       

' * 返回值 :  0或-1,-1=不成功,0=成功
' *
' *
' *****************************************************************************/
'short APIENTRY UpDateFlash_9030(unsigned short Board_NO);


'/******************************************************************************
' *
' * 函数名   : ReadFlash_9030
' * 函数编号 : 128

' * 描述   :  读 9030卡上 用户flash
' * 参数   :
    
'      Board_NO : 0－3, 板号
'      offset    : 0-199 偏移值
'      len       : 1-4
       

' * 返回值 :
' *
' *
' *****************************************************************************/
'unsigned long APIENTRY ReadFlash_9030(unsigned short Board_NO,unsigned short offset,unsigned short len);

'/******************************************************************************
' *
' * 函数名   : ReadFlashChar_9030
' * 函数编号 : 128

' * 描述   :  读 9030卡上 用户flash
' * 参数   :
    
'      Board_NO : 0－3, 板号
'      offset    : 0-199 偏移值
'      len       : 1-4
'Data:             数据指针
       

' * 返回值 :  0或-1,-1=不成功,0=成功
' *
' *
' *****************************************************************************/
'short APIENTRY ReadFlashChar_9030(unsigned short Board_NO,unsigned short offset,unsigned short len,unsigned char *Data);


