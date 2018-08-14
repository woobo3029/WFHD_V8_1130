Attribute VB_Name = "MdlAI"
Option Explicit

Declare Function GetCruve Lib "Curve.dll" (ByVal PathName As String, ByRef ControlParam As Double) As Long

Type m_p
    X As Double
    Y As Double
End Type

Dim m_point(10000) As m_p
'Dim point_Number As Long
Dim lamda As Double '//控制点的密集程度

'std::vector<double> med_points_x;
'std::vector<double> med_points_y;

Dim med_points_x(100000) As Double
Dim med_points_y(100000) As Double
Dim med_points_count As Long
Dim buchang_x(10000) As Double
Dim buchang_y(10000) As Double



'int splite(CString strsrc,CString strsplite,CString *Array)
'{
'    int k = 0;
'    While (strsrc.GetLength() > 0)
'    {
'        int pos = strsrc.Find(strsplite,0);//定位分割符
'        CString strleft;
'        if(pos!=-1)
'        {
'            //定位成功
'            strleft = strsrc.Left(pos);//前面的字符串为新的分割单元
'            Array[k] = strleft;
'            k++;
'            strsrc = strsrc.Right(strsrc.GetLength()-pos-strsplite.GetLength());//指定新的分割对象
'        }
'        Else
'        {
'            //定位不成功
'            strleft = strsrc;//将原字符串作为分割单元入目标数组
'            Array[k] = strleft;
'            k++;
'            strsrc.Empty();//清空原字符串
'        }
'    }
'    return k;
'}
    
'///////////////////////////////////////////////////////////////////////////////////////
'//
'// 函数名称：DrawBezier()
'//
'// 输入参数:
'//          PathName:AI文件路径名
'//          x,y:中间点的坐标
'//          controlParam:控制点的密集程度，取值在0--1之间，越小点越多，现在做的取值为0.0001
'//
'// 返回值: 过程点的总数目
'//
'// 说明:
'//       该函数函数得到曲线过程的中间点
'//
'//////////////////////////////////////////////////////////////////////////////////////
 
Function GetDataFromAI(ByVal PathName As String, ByVal ControlParam As Double) As Long
    lamda = ControlParam '//控制点的密集程度
    Dim fp As Long
    Dim str As String

    fp = FreeFile
    Open PathName For Input As fp
    
    Dim k As Long, I As Long
    
    'point_Number = 0
    med_points_count = 0

    Dim CurrentPointx As Double, CurrentPointy As Double, endx As Double, endy As Double

    Dim VArray As Variant '//每行以空格分割后的字符数组
    Dim VArrayLength As Long
    Dim db(1000) As Double

    Dim SaveDate As Boolean, IfSplite As Boolean
    
    SaveDate = False
    IfSplite = False

    Do While Not EOF(fp)
        Line Input #fp, str
        
        If str = "%%EndSetup" Then '
            IfSplite = True
        ElseIf str = "%%PageTrailer" Then
            IfSplite = False
        End If
        
        If IfSplite = True And str <> "%%EndSetup" Then
            VArray = Split(str, " ") '//以空格划分字符串
            VArrayLength = UBound(VArray)
            
            'DPoint *m_point = new DPoint[ArrayLength];
            If VArray(VArrayLength) = "m" Then '//找到起点行，开始保存数据
                SaveDate = True
            End If

            If SaveDate = True And VArrayLength >= 2 Then
                For I = 0 To VArrayLength - 1
                    db(I) = Val(VArray(I)) '//将CString转换为double
                    db(I) = db(I) / 72 * 25.400045
                Next
                
                Dim a As String
                a = VArray(VArrayLength)

'Debug.Print "line="; str
                Select Case a
                Case "m"
                    For I = 0 To VArrayLength - 1 Step 2
                        CurrentPointx = db(I)
                        CurrentPointy = db(I + 1)
'Debug.Print "m x,y="; CurrentPointx, CurrentPointy
                    Next
                    
                    med_points_x(med_points_count) = 999999
                    med_points_y(med_points_count) = 999999
                    
                    med_points_count = med_points_count + 1
                    
                    'point_Number = point_Number + 1
                    
                Case "L":
                    For I = 0 To VArrayLength - 1 Step 2
                        endx = db(I)
                        endy = db(I + 1)
'Debug.Print "L x,y="; endx, endy
                    Next

                    makeBeeline CurrentPointx, CurrentPointy, endx, endy

                    CurrentPointx = endx
                    CurrentPointy = endy
                
                Case "l":
                    For I = 0 To VArrayLength - 1 Step 2
                        endx = db(I)
                        endy = db(I + 1)
'Debug.Print "l x,y="; endx, endy
                    Next

                    makeBeeline CurrentPointx, CurrentPointy, endx, endy
                    
                    CurrentPointx = endx
                    CurrentPointy = endy
                
                Case "c":
                    k = 1
                    m_point(0).X = CurrentPointx
                    m_point(0).Y = CurrentPointy
                    For I = 0 To VArrayLength - 1 Step 2
                        m_point(k).X = db(I)
                        m_point(k).Y = db(I + 1)
'Debug.Print "c x,y="; m_point(k).X, m_point(k).Y
                        k = k + 1
                    Next

                    makeCurve k ', m_point
                    CurrentPointx = m_point(k - 1).X
                    CurrentPointy = m_point(k - 1).Y
                    
                Case "C":
                    k = 1
                    m_point(0).X = CurrentPointx
                    m_point(0).Y = CurrentPointy
                    For I = 0 To VArrayLength - 1 Step 2
                        m_point(k).X = db(I)
                        m_point(k).Y = db(I + 1)
'Debug.Print "C X,Y="; m_point(k).X, m_point(k).Y
                        k = k + 1
                    Next

                    makeCurve k ', m_point
                    CurrentPointx = m_point(k - 1).X
                    CurrentPointy = m_point(k - 1).Y
                
                Case "y":
                    k = 1
                    m_point(0).X = CurrentPointx
                    m_point(0).Y = CurrentPointy
                    For I = 0 To VArrayLength - 1 Step 2
                        m_point(k).X = db(I)
                        m_point(k).Y = db(I + 1)
'Debug.Print "y x,y="; m_point(k).X, m_point(k).Y
                        k = k + 1
                    Next
                    m_point(k).X = m_point(k - 1).X
                    m_point(k).Y = m_point(k - 1).Y
'Debug.Print "y x,y="; m_point(k).X, m_point(k).Y
                    k = k + 1

                    makeCurve k ', m_point
                    CurrentPointx = m_point(k - 1).X
                    CurrentPointy = m_point(k - 1).Y
                
                Case "Y":
                    k = 1
                    m_point(0).X = CurrentPointx
                    m_point(0).Y = CurrentPointy
                    For I = 0 To VArrayLength - 1 Step 2
                        m_point(k).X = db(I)
                        m_point(k).Y = db(I + 1)
                        k = k + 1
                    Next
                    m_point(k).X = m_point(k - 1).X
                    m_point(k).Y = m_point(k - 1).Y
                    k = k + 1

                    makeCurve k ', m_point
                    CurrentPointx = m_point(k - 1).X
                    CurrentPointy = m_point(k - 1).Y
                
                Case "v":
                    k = 2
                    m_point(0).X = CurrentPointx
                    m_point(0).Y = CurrentPointy
                    m_point(1).X = CurrentPointx
                    m_point(1).Y = CurrentPointy
                    For I = 0 To VArrayLength - 1 Step 2
                        m_point(k).X = db(I)
                        m_point(k).Y = db(I + 1)
                        k = k + 1
                    Next

                    makeCurve k ', m_point
                    CurrentPointx = m_point(k - 1).X
                    CurrentPointy = m_point(k - 1).Y
                
                Case "V":
                    k = 2
                    m_point(0).X = CurrentPointx
                    m_point(0).Y = CurrentPointy
                    m_point(1).X = CurrentPointx
                    m_point(1).Y = CurrentPointy
                    For I = 0 To VArrayLength - 1 Step 2
                        m_point(k).X = db(I)
                        m_point(k).Y = db(I + 1)
                        k = k + 1
                    Next
               
                    makeCurve k ', m_point
                    CurrentPointx = m_point(k - 1).X
                    CurrentPointy = m_point(k - 1).Y
                End Select
            End If
        End If
    Loop
    
    Close #fp
    
'//  std::string filename;
'//  filename = "Data_From_AI_File.txt";
'//  ofstream ofstr;
'//     ofstr.open(filename.c_str());
'//  for (i = 0; i <= point_Number; i++)
'//  {
'//      if (med_points_x[i] != NULL)
'//      {
'//          double a = med_points_x[i];
'//          double b = med_points_y[i];
'//          ofstr << a <<' ' << b <<' '<<endl;
'//      }
'//      else
'//      {
'//          ofstr << "******" << endl;
'//      }
'//  }
'
'    'fp = FreeFile
'    'Open "Data_From_AI_File.txt" For Output As fp
'    For I = 0 To med_points_count - 1
'        If med_points_x(I) <> 999999 And med_points_y(I) <> 999999 Then
'            'Print #fp, med_points_x(I), med_points_y(I)
'            Debug.Print I, med_points_x(I), med_points_y(I)
'        Else
'            'Print #fp, "<"
'            Debug.Print I, "<"
'        End If
'    Next
    'Close #fp

    GetDataFromAI = med_points_count
End Function


'//直线
Sub makeBeeline(CurrentPointx As Double, CurrentPointy As Double, endx As Double, endy As Double)
    '//直线方程y=ax+b
    
'    Dim a As Double, b As Double, I As Double
'
'    a = (CurrentPointy - endy) / (CurrentPointx - endx)
'    b = CurrentPointy - a * CurrentPointx
'
'    If CurrentPointx > endx Then
'        I = endx
'        Do While I < CurrentPointx
'            med_points_x(med_points_count) = I
'            med_points_y(med_points_count) = a * I + b
'            med_points_count = med_points_count + 1
'
''//          med_points[point_Number].x = i;
''//          med_points[point_Number].y = a * i + b;
'
''            point_Number = point_Number + 1
'            I = I + 1
'        Loop
'    Else
'        I = CurrentPointx
'        Do While I < endx
'            med_points_x(med_points_count) = I
'            med_points_y(med_points_count) = a * I + b
'            med_points_count = med_points_count + 1
'
''//          med_points[point_Number].x = i;
''//          med_points[point_Number].y = a * i + b;
'
''            point_Number = point_Number + 1
'            I = I + 1
'        Loop
'    End If

            med_points_x(med_points_count) = CurrentPointx
            med_points_y(med_points_count) = CurrentPointy
            med_points_count = med_points_count + 1

            med_points_x(med_points_count) = endx
            med_points_y(med_points_count) = endy
            med_points_count = med_points_count + 1

End Sub

'//Bezier曲线
Sub makeCurve(points As Long) ',DPoint *m_point)
    Dim I As Long, n As Long
    Dim tempx As Double, tempy As Double, t As Double, k As Long
    
    I = 0
    tempx = 0#
    tempy = 0#

    t = 0#
    n = points - 1
    
    k = 0
    Do 'While (t <= 1)
        tempx = 0#
        tempy = 0#
        I = 0
        
        Do While (I <= n)
            'tempx = tempx + (m_point(I).X * pow(t, I) * pow((1 - t), (n - I)) * PaiLie_C(n, I))
            'tempy = tempy + (m_point(I).Y * pow(t, I) * pow((1 - t), (n - I)) * PaiLie_C(n, I))
            tempx = tempx + (m_point(I).X * (t ^ I) * ((1 - t) ^ (n - I)) * PaiLie_C(n, I))
            tempy = tempy + (m_point(I).Y * (t ^ I) * ((1 - t) ^ (n - I)) * PaiLie_C(n, I))
            I = I + 1
        Loop
        
        med_points_x(med_points_count) = tempx
        med_points_y(med_points_count) = tempy
        med_points_count = med_points_count + 1
        
'//      med_points[point_Number].x = tempx;
'//      med_points[point_Number].y = tempy;
'        point_Number = point_Number + 1
        If k = 1 Then
            Exit Do
        End If
        
        t = t + lamda
        If t > 1 Then
            t = 1
        End If
        If t = 1 Then
            k = 1
        End If
    Loop
End Sub

'//计算排列组合的公式C(i,n) =  n!/i!/(n-i)!
Function PaiLie_C(n As Long, I As Long) As Long
    PaiLie_C = Factorial(n) / ((Factorial(I)) * Factorial(n - I))
End Function

'//计算aaa!
Function Factorial(a As Long) As Long
    Dim temp As Long, I As Long
    
    temp = 1
    For I = 1 To a
        temp = temp * I
    Next
    Factorial = temp
End Function

Sub ImportAI(ByVal fn As String)
    Dim n As Long
    Dim k As Double
    
    k = 0.05
    
    n = GetDataFromAI(fn, k)
    
    '打印调试
    'PrintCoordinate med_points_count
    
    'Left_BuchangPoints med_points_count - 1
    
    'PrintCoordinateAfterBuchang med_points_count
    
    ConvertAIToCMP
End Sub

Sub ConvertAIToCMP()
    Dim Point0 As Path_Point, Point1 As Path_Point, Pointm As Path_Point
    Dim cx As Double, cy As Double, cz As Double, r As Double, r2 As Double, Angle0 As Double, Angle1 As Double
    Dim I As Long, k As Integer, n As Long, id0 As Long, ux As Double, uy As Double, uz As Double, d As Double, t As Long
    
    On Error Resume Next
            
    I = 0
    n = med_points_count - 1
    If n > 0 Then
        For k = 0 To n
            If med_points_x(k) = 999999 And med_points_y(k) = 999999 Then
                I = 0
            Else
                I = I + 1
            End If
                
            If I > 0 Then
                If I = 1 Then
                    Point0.X = med_points_x(k)
                    Point0.Y = med_points_y(k)
                    Point0.z = 0
                    Point0.Layer = 1 'DXF.PolyLines(I).layer_id
                    'Point0.color = LayerColor(PathColor(DXF.PolyLines(I).color_id).Index - 1)
                    'Point0.insert_id = insert_id
                    
                    CatchOrAddPoint Point0
                    id0 = Point0.id
                Else
                    t = 0
                    If k < n Then
                        cx = med_points_x(k)
                        cy = med_points_y(k)
                        ux = med_points_x(k + 1)
                        uy = med_points_y(k + 1)
                        
                        d = Sqr((cx - PointList(id0).X) * (cx - PointList(id0).X) + (cy - PointList(id0).Y) * (cy - PointList(id0).Y))
                        Angle0 = GetAngle(PointList(id0).X, PointList(id0).Y, cx, cy, ux, uy)
                        If d > 5 Or Abs(Angle0) > 3 Then
                            t = 1
                        End If
                    Else
                        t = 1
                    End If
                    
                    If t = 1 Then
                        Point1.X = med_points_x(k)
                        Point1.Y = med_points_y(k)
                        Point1.z = 0
                        Point1.Layer = 1
                        
                        If PointList(id0).X <> Point1.X Or PointList(id0).Y <> Point1.Y Then
                            CatchOrAddPoint Point1
                            
                            If id0 <> Point1.id Then
                                AddSegment id0, Point1.id
                                
                                SegmentList(SegmentCount).Layer = 1
                                id0 = Point1.id
                            End If
                        End If
                    End If
                End If
                
            End If
        Next
    End If
End Sub
Sub Left_BuchangPoints(cnt As Long)
Dim I As Integer
Dim x0, y0 As Double
Dim x1, y1 As Double
Dim x2, y2 As Double

Dim xc1, yc1 As Double
Dim xs1, ys1 As Double
Dim xc2, yc2 As Double
Dim xl1 As Double
Dim yl1 As Double
Dim xl2 As Double
Dim yl2 As Double

Dim sign As Integer
Dim r  As Double

    r = 0.7
    If r > 0 Then
        sign = 1
    Else
        sign = -1
    End If
    For I = 1 To cnt
        If med_points_x(I) = 99999 And med_points_y(I) = 99999 Then
            Exit For
        End If
        If I = 1 Then
            x0 = med_points_x(cnt)
            y0 = med_points_y(cnt)
            
            x1 = med_points_x(I)
            y1 = med_points_y(I)
            
            x2 = med_points_x(I + 1)
            y2 = med_points_y(I + 1)
        
        ElseIf I = cnt Then
        
            x0 = med_points_x(I - 1)
            y0 = med_points_y(I - 1)
            
            x1 = med_points_x(I)
            y1 = med_points_y(I)
            
            x2 = med_points_x(1)
            y2 = med_points_y(1)
        
        Else
        
            x0 = med_points_x(I - 1)
            y0 = med_points_y(I - 1)
            
            x1 = med_points_x(I)
            y1 = med_points_y(I)
            
            x2 = med_points_x(I + 1)
            y2 = med_points_y(I + 1)
        End If
        xl1 = (x1 - x0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        yl1 = (y1 - y0) / Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0))
        xl2 = (x2 - x1) / Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
        yl2 = (y2 - y1) / Sqr((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
        If sign * (xl1 * yl2 - xl2 * yl1) >= 0 Then  'printf("type of Reduction!\n");/*提示缩短型*/
        
            '/*刀补的进行*/
            '//Label9.Caption = "缩短型"
            If (xl1 * yl2 - xl2 * yl1) = 0 Then   '/*特殊情况两直线共线转接角为180°*/
            
                xs1 = x1 - r * yl1
                ys1 = y1 + r * xl1
            
            Else
            
                xs1 = x1 + ((xl2 - xl1) * r) / (xl1 * yl2 - xl2 * yl1)
                ys1 = y1 + ((yl2 - yl1) * r) / (xl1 * yl2 - xl2 * yl1)
            End If
            xc1 = (x0 + xs1 - x1)
            yc1 = (y0 + ys1 - y1)
            xc2 = (x2 + xs1 - x1)
            yc2 = (y2 + ys1 - y1)
            
        
        ElseIf sign * (yl2 * xl1 - xl2 * yl1) < 0 And (yl2 * yl1 + xl2 * xl1) >= 0 Then
        '   /*判断为伸长型*/
         '   /*提示伸长型*/
            
          '  /*刀补进行*/
           ' //Label9.Caption = "伸长型"
            xs1 = x1 + ((xl2 - xl1) * r) / (xl1 * yl2 - xl2 * yl1)
            ys1 = y1 + ((yl2 - yl1) * r) / (xl1 * yl2 - xl2 * yl1)
            xc1 = (x0 + xs1 - x1)
            yc1 = (y0 + ys1 - y1)
            xc2 = (x2 + xs1 - x1)
            yc2 = (y2 + ys1 - y1)
        
        ElseIf sign * (yl2 * xl1 - xl2 * yl1) < 0 And (yl2 * yl1 + xl2 * xl1) < 0 Then
        
            '/*判断为插入型*/
            '/*强制为伸长型*/
            
            '/*刀补进行*/
            '//Label9.Caption = "伸长型"
            xs1 = x1 + ((xl2 - xl1) * r) / (xl1 * yl2 - xl2 * yl1)
            ys1 = y1 + ((yl2 - yl1) * r) / (xl1 * yl2 - xl2 * yl1)
            xc1 = (x0 + xs1 - x1)
            yc1 = (y0 + ys1 - y1)
            xc2 = (x2 + xs1 - x1)
            yc2 = (y2 + ys1 - y1)
        End If
        buchang_x(I) = xs1
        buchang_y(I) = ys1
    Next
    
    For I = 1 To cnt
        med_points_x(I) = buchang_x(I)
        med_points_y(I) = buchang_y(I)
    Next
End Sub

Sub PrintCoordinate(cnt As Long)
Dim File As Integer
Dim I As Long
Dim xtemp As Double
Dim ytemp As Double
Dim str1 As String
Dim str2 As String
    File = FreeFile
    Open "c:\hd_debug\" + "CoordPoint.txt" For Output As #File
    Print #File, "序号"; Tab(8); "x坐标"; Tab(24); "y坐标"
    For I = 1 To cnt
        'Print #File, i; Tab(8); m_point(i).X; m_point(i).Y; med_points_x(i); med_points_y(i)
        xtemp = med_points_x(I)
        ytemp = med_points_y(I)
        'If xtemp < 0.001 And xtemp > -0.001 Then
        '    xtemp = 0
        'End If
        'If ytemp < 0.001 And ytemp > -0.001 Then
        '    ytemp = 0
        'End If
        'str1 = Format(xtemp, "0.000")
        'str2 = Format(ytemp, "0.000")
        xtemp = Round(xtemp, 3)
        ytemp = Round(ytemp, 3)
        
        'Print #File, "N"; i; Tab(8); "G00X"; str1; Tab(24); "Y"; str2
        Print #File, "N"; I; Tab(8); "G00X"; xtemp; Tab(24); "Y"; ytemp
    Next
    Close #File
End Sub

Sub PrintCoordinateAfterBuchang(cnt As Long)
Dim File As Integer
Dim I As Long
Dim xtemp As Double
Dim ytemp As Double
Dim str1 As String
Dim str2 As String
    File = FreeFile
    Open "c:\hd_debug\" + "CoordPoint_buchang.txt" For Output As #File
    Print #File, "序号"; Tab(8); "x坐标"; Tab(24); "y坐标"
    For I = 1 To cnt
        'Print #File, i; Tab(8); m_point(i).X; m_point(i).Y; med_points_x(i); med_points_y(i)
        xtemp = med_points_x(I)
        ytemp = med_points_y(I)
        'If xtemp < 0.001 And xtemp > -0.001 Then
        '    xtemp = 0
        'End If
        'If ytemp < 0.001 And ytemp > -0.001 Then
        '    ytemp = 0
        'End If
        'str1 = Format(xtemp, "0.000")
        'str2 = Format(ytemp, "0.000")
        xtemp = Round(xtemp, 3)
        ytemp = Round(ytemp, 3)
        
        'Print #File, "N"; i; Tab(8); "G00X"; str1; Tab(24); "Y"; str2
        Print #File, "N"; I; Tab(8); "G00X"; xtemp; Tab(24); "Y"; ytemp
    Next
    Close #File
End Sub
