Attribute VB_Name = "MdlDXF"
Option Explicit

Type DXF_Layer
    id As Long
    Name As String
    color As Long
    Width As Integer
    style As Integer
    mode As Integer
    FontName As String
    Frozen As Boolean
    Locked As Boolean
    Hidden As Boolean
End Type

Type DXF_Point
    X As Double
    Y As Double
    z As Double
    layer_id As Long
End Type

Type DXF_Line
    P1 As DXF_Point
    P2 As DXF_Point
    layer_id As Long
End Type

Type DXF_Arc
    Center As DXF_Point
    radius As Double
    Angle1 As Double
    angle2 As Double
    layer_id As Long
End Type

Type DXF_SPLine
    Vertex() As DXF_Point
    layer_id As Long
End Type

Type DXF_PolyLine
    Vertex() As DXF_Point
    flag As Long
    layer_id As Long
End Type

Type DXF_Ellipse
    'F1 As DXF_Point
    'F2 As DXF_Point
    'P1 As DXF_Point
    'NumPoints As Integer
    
    Center As DXF_Point
    EndpointOfMajorAxis As DXF_Point
    RatioOfMinorAxisToMajor As Double
    Angle1 As Double
    angle2 As Double
    layer_id As Long
End Type

Type DXF_Insert
    Name As String
    ScaleX As Double
    ScaleY As Double
    angle As Double
    Base As DXF_Point
    layer_id As Long
    ExtrusionDirZ As Double
End Type

Type DXF_Data
    Name As String
    Base As DXF_Point
    PointCount As Long
    LineCount As Long
    ArcCount As Long
    EllipseCount As Long
    SPLineCount As Long
    PolyLineCount As Long
    InsertCount As Long
    
    points() As DXF_Point
    Lines() As DXF_Line
    Arcs() As DXF_Arc
    Ellipses() As DXF_Ellipse
    SPLines() As DXF_SPLine
    PolyLines() As DXF_PolyLine
    Inserts() As DXF_Insert
End Type

Type DXF_DataGroup
    Code As String
    value As Variant
End Type

Type DXF_DataSet
    Type As String
    data() As DXF_DataGroup
End Type

Dim DXFData() As DXF_Data
Dim NewDXFData() As DXF_Data
Dim DXFLayers() As DXF_Layer

Public Sub ImportDXF(DXFFileName As String)
    Dim ff As Integer
    Dim DXFLine As String
    Dim Version As String
    Dim ENDSEC As Boolean
    Dim Section() As String
    Dim GetNext As Boolean
    Dim bCount As Integer
    
    ReDim DXFData(0) As DXF_Data  'Zero will be the PV, the others will be blocks if they exists.
    
    On Error GoTo 0 'ErrorHandler
    
    ff = FreeFile
    Open DXFFileName For Input As #ff
    'First we need to find the version number . . .
    FindCommand ff, "$ACADVER"
    FindCommand ff, "1"
    Line Input #ff, Version
    
    'Skip to the TABLES section of the DXF file
    FindCommand ff, "TABLES"
    GetSection ff, "LAYER", "ENDSEC", "ENDSEC", Section()
    ParseDXFLayers Section(), DXFLayers()   '分析层信息
    
    'Next we skip all the header stuff and get to the section called 'BLOCKS'
    FindCommand ff, "BLOCKS"
    '---------------------------
    'BLOCKS are groups of geometry that are re-useable within the drawing. They may appear several times within one drawing
    'and if the block is modified it automatically modifies each time wherever it's used within the drawing
    GetNext = True
    bCount = 1
    Do While Not ENDSEC
        'First we load in a SECTION into an array (BLOCK) to (ENDBLK). We do this until we come across the "ENDSEC" command
        If GetSection(ff, "BLOCK", "ENDBLK", "ENDSEC", Section()) Then
            'We have a "BLOCK" in the array. So we have to advance our array of BLOCKS (Geometry)
            ReDim Preserve DXFData(bCount) As DXF_Data
            If ParseDXF(Section(), DXFData(UBound(DXFData)), DXFLayers(), Version, True) Then
                bCount = bCount + 1
            End If
        Else
            ENDSEC = True
        End If
    Loop
    'Now we go after the 'Primary View Entities
    ENDSEC = False
    GetSection ff, "ENTITIES", "ENDSEC", "ENDSEC", Section()    '获取所有实体段字符
    'This grabs ALL PV ENTITIES . . . kind of like one huge block
    Close #ff 'We can close the file because we're finished with it
    
    'Next we fill the array with geometry data
    'ParseDXF Section(), DXFData(0), DXFLayers(), Version, False
    ParseDXFEntitiesSection Section(), DXFData(0), DXFLayers(), Version '分析实体段字符
    
    'ConvertDXFToPath
    ConvertDXFToCMP 0

    'ReversePath
    
    Erase Section
    Erase DXFData
    Erase DXFLayers
    Erase NewDXFData
    Exit Sub
    
ErrorHandler:
    'MsgBox Err.Description
    MsgBox "不可识别的数据格式！ 文件已损坏或格式的版本不对。 ", vbCritical + vbOKOnly, ""
End Sub


Function GetSection(FileNum As Integer, Start As String, Finish As String, EndString As String, sArray() As String) As Boolean
'在文件中找出从Start开始的的字符
'若找到EndString，则退出函数，返回FALSE
'若找到FinishString， 则停止查找，返回true
    ReDim sArray(0) As String
    Dim temp As String
    Dim I As Long
    
    Do While temp <> Start
        Line Input #FileNum, temp
        temp = UCase(Trim(temp))
        If temp = EndString Then
            GetSection = False
            Exit Function
        End If
    Loop
    
    Do While temp <> Finish
        Line Input #FileNum, temp
        temp = UCase(Trim(temp))
        If temp <> Finish Then
            ReDim Preserve sArray(I) As String
            sArray(I) = temp
            I = I + 1
        End If
    Loop
    GetSection = True
End Function

Sub FindCommand(FileNum As Integer, Command As String)
    Dim X As String

    Do While UCase(Trim(X)) <> UCase(Command)
        Line Input #FileNum, X
    Loop
End Sub

Sub ParseDXFLayers(sArray() As String, ByRef Layers() As DXF_Layer)
'分析DXF文件 层信息
    Dim I As Long
    Dim k As Integer
    Dim c As Integer
    
    Do
        I = SearchSection(sArray(), I, "LAYER")
        If I = -1 Then
            Exit Do
        End If
        
        ReDim Preserve Layers(k) As DXF_Layer
        
        I = SearchSection(sArray(), I, "2")
        Layers(k).Name = Trim(sArray(I + 1))
        I = SearchSection(sArray(), I, "70")
        Select Case sArray(I + 1)
            Case 0
                Layers(k).Frozen = False
                Layers(k).Locked = False
            Case 1
                Layers(k).Frozen = True
                Layers(k).Locked = False
            Case 2
                Layers(k).Frozen = True
                Layers(k).Locked = False
            Case 3
                Layers(k).Frozen = True
                Layers(k).Locked = False
            Case 4
                Layers(k).Frozen = False
                Layers(k).Locked = True
            Case 5
                Layers(k).Frozen = True
                Layers(k).Locked = True
            Case 6
                Layers(k).Frozen = True
                Layers(k).Locked = True
        End Select
        
        I = SearchSection(sArray(), I, "62")
        c = CInt(sArray(I + 1))
        If c < 0 Then
            Layers(k).Hidden = True
        End If
        
        c = Abs(c)
        If c > LayerMax Then c = 0
        Layers(k).color = LayerColor(c)
        
        I = SearchSection(sArray(), I, "6")
        If IsNumeric(sArray(I + 1)) Then
            Layers(k).style = sArray(I + 1)
        End If
        If Layers(k).style > 4 Then
            DXFLayers(k).style = 0
        End If
        Layers(k).Width = 1
        Layers(k).FontName = "Arial Black"
        k = k + 1
    Loop
End Sub

Function SearchSection(sArray() As String, Start As Long, value As String, Optional stp As Integer = 1) As Long
    Dim I As Long
    
    For I = Start To UBound(sArray) Step stp
        If sArray(I) = value Then
            SearchSection = I
            Exit Function
        End If
    Next I
    SearchSection = -1
End Function

Function FindStart(sArray() As String, Start As Long, Version As String)
    Dim I As Long
    
    Select Case Version
        Case "AC1012", "AC1013", "AC1014", "AC1015", "AC1016", "AC1017", "AC1018"
            I = SearchSection(sArray(), Start, "100") + 1
            I = SearchSection(sArray(), I, "100") + 2
            FindStart = I
            Exit Function
        Case Else
            For I = Start To UBound(sArray)
                If sArray(I) = "10" Then
                    FindStart = I
                    Exit Function
                End If
            Next I
    End Select
    FindStart = -1
End Function

Function FindStart_SPLine(sArray() As String, Start As Long, Version As String)
    Dim I As Long, j As Long
    
    Select Case Version
        Case "AC1012", "AC1013", "AC1014", "AC1015", "AC1016", "AC1017", "AC1018"
            I = SearchSection(sArray(), Start, "100") + 1
            I = SearchSection(sArray(), I, "100") + 2
            FindStart_SPLine = I
            Exit Function
        Case Else
            j = SearchSection(sArray(), Start, UCase("AcDbSpline")) + 1
            For I = j To UBound(sArray) Step 2
                If sArray(I) = "40" Or sArray(I) = "10" Then
                    FindStart_SPLine = I
                    Exit Function
                End If
            Next I
    End Select
    FindStart_SPLine = -1
End Function

Function IsDXFCommand(InText As String)
    Select Case UCase(InText)
        Case "POINT", "LINE", "VERTEX", "POLYLINE", "LWPOLYLINE", "CIRCLE", "ARC", "ELLIPSE", "TEXT", "INSERT", "DIMENSION", "SPLINE", "MTEXT"
            IsDXFCommand = True
        Case Else
            IsDXFCommand = False
    End Select
End Function

Function ParseDXF(sArray() As String, DXF As DXF_Data, Layers() As DXF_Layer, Version As String, isBlock As Boolean) As Boolean
'
    'On Local Error GoTo exitMe:
    Dim I As Long
    Dim j As Long
    Dim k As Long
    Dim p As Long
    Dim DataSet() As DXF_DataSet
    Dim Endword As String
    
    I = 0
    k = 0
    If isBlock Then
        I = SearchSection(sArray(), I, "2") + 1
    End If
    DXF.Name = sArray(I)
    
    For j = I To UBound(sArray)
        If IsDXFCommand(sArray(j)) Then 'We Found an ENTITY COMMAND
            ReDim Preserve DataSet(k) As DXF_DataSet
            DataSet(k).Type = sArray(j)
            'I am not sure if a BLOCK can use a block. Either way, this is designed to work even if you can
            Select Case DataSet(k).Type
                Case "INSERT", "DIMENSION"
                    'KEY "2" on an INSERT provides the BLOCK name to be inserted
                    j = SearchSection(sArray(), j, "2")
                Case "SPLINE"
                    j = FindStart_SPLine(sArray(), j, Version)
                Case Else
                    j = FindStart(sArray(), j, Version)
            End Select
            
            p = 0
            If DataSet(k).Type = "POLYLINE" Then
                'Endword = "ENDSEC" Else Endword = "0"
                Do While UCase(sArray(j)) <> "ENDSEC" And UCase(sArray(j)) <> "SEQEND" And UCase(sArray(j + 1)) <> "ENDSEC" And UCase(sArray(j + 1)) <> "SEQEND" And j + 1 <= UBound(sArray)
                    ReDim Preserve DataSet(k).data(p)
                    DataSet(k).data(p).Code = sArray(j)
                    DataSet(k).data(p).value = sArray(j + 1)
                    p = p + 1
                    j = j + 2
                Loop
            Else
                Do While UCase(sArray(j)) <> "0"
                    ReDim Preserve DataSet(k).data(p)
                    DataSet(k).data(p).Code = sArray(j)
                    DataSet(k).data(p).value = sArray(j + 1)
                    p = p + 1
                    j = j + 2
                    
                    If j + 1 > UBound(sArray) Then
                        Exit Do
                    End If
                Loop
            End If
            PrepareDXFEntity DataSet(k), DXF, Layers()
            k = k + 1
        End If
    Next j
    ParseDXF = True
    Exit Function
    
ExitMe:
    MsgBox "ERROR  " & Err.Description
End Function

Function ParseDXFBlock(sArray() As String, Block As DXF_Data, Layers() As DXF_Layer, Version As String) As Boolean
    'On Local Error GoTo exitMe:
    Dim I As Long
    Dim j As Long
    Dim k As Long
    Dim u As Long
    Dim p As Long
    Dim DataSet() As DXF_DataSet
    Dim Endword As String
    
    I = 0
    k = 0
    I = SearchSection(sArray(), I, "2")
    Block.Name = sArray(I + 1)
    I = SearchSection(sArray(), I, "10")
    Block.Base.X = sArray(I + 1)
    I = SearchSection(sArray(), I, "20")
    Block.Base.Y = sArray(I + 1)
    
    For j = I To UBound(sArray)
        If IsDXFCommand(sArray(j)) Then 'We Found an ENTITY COMMAND
            ReDim Preserve DataSet(k) As DXF_DataSet
            DataSet(k).Type = sArray(j)
            'I am not sure if a BLOCK can use a block. Either way, this is designed to work even if you can
            Select Case DataSet(k).Type
                Case "INSERT", "DIMENSION"
                    'KEY "2" on an INSERT provides the BLOCK name to be inserted
                    j = SearchSection(sArray(), j, "2")
                    
                Case "SPLINE"
                    j = FindStart_SPLine(sArray(), j, Version)
                Case Else
                    'I = SearchSection(sArray(), j + 1, "8", 2)
                    'DataSet(k).LayerName = sArray(I + 1)
                    'I = SearchSection(sArray(), I + 2, "62", 2)
                    'DataSet(k).ColorIndex = sArray(I + 1)
                    
                    j = FindStart(sArray(), j, Version)
            End Select
            
            p = 0
            If DataSet(k).Type = "POLYLINE" Then
                'Endword = "ENDSEC" Else Endword = "0"
                Do While UCase(sArray(j)) <> "ENDSEC" And UCase(sArray(j)) <> "SEQEND" And UCase(sArray(j + 1)) <> "ENDSEC" And UCase(sArray(j + 1)) <> "SEQEND" And j + 1 <= UBound(sArray)
                    ReDim Preserve DataSet(k).data(p)
                    DataSet(k).data(p).Code = sArray(j)
                    DataSet(k).data(p).value = sArray(j + 1)
                    p = p + 1
                    j = j + 2
                Loop
            Else
                Do While UCase(sArray(j)) <> "0"
                    ReDim Preserve DataSet(k).data(p)
                    DataSet(k).data(p).Code = sArray(j)
                    DataSet(k).data(p).value = sArray(j + 1)
                    p = p + 1
                    j = j + 2
                    
                    If j + 1 > UBound(sArray) Then
                        Exit Do
                    End If
                Loop
            End If
            PrepareDXFEntity DataSet(k), Block, Layers(), True
            k = k + 1
        End If
    Next j
    ParseDXFBlock = True
    Exit Function
    
ExitMe:
    MsgBox "ERROR  " & Err.Description
End Function

Function ParseDXFEntitiesSection(sArray() As String, EntitiesSection As DXF_Data, Layers() As DXF_Layer, Version As String) As Boolean
    'On Local Error GoTo exitMe:
    Dim I As Long
    Dim j As Long
    Dim k As Long
    Dim u As Long
    Dim p As Long
    Dim DataSet() As DXF_DataSet
    Dim Endword As String
    
    FrmMain.ProgressBar.Min = 0
    FrmMain.ProgressBar.Max = UBound(sArray) + 1
   
    I = 0
    k = 0
    
    For j = I To UBound(sArray)
    
        FrmMain.ProgressBar.value = j + 1
        FrmMain.ProgressBar.Refresh
        
        If IsDXFCommand(sArray(j)) Then 'We Found an ENTITY COMMAND
            ReDim Preserve DataSet(k) As DXF_DataSet
            DataSet(k).Type = sArray(j)
            'I am not sure if a BLOCK can use a block. Either way, this is designed to work even if you can
            Select Case DataSet(k).Type
                Case "INSERT", "DIMENSION"
                    'KEY "2" on an INSERT provides the BLOCK name to be inserted
                    j = SearchSection(sArray(), j, "2")
                
                Case "SPLINE"
                    j = FindStart_SPLine(sArray(), j, Version)  'j值为SPLINE段有效字符开始序号
                Case Else
                    'I = SearchSection(sArray(), j + 1, "8", 2)
                    'DataSet(k).LayerName = sArray(I + 1)
                    'I = SearchSection(sArray(), I + 2, "62", 2)
                    'DataSet(k).ColorIndex = sArray(I + 1)
                    
                    j = FindStart(sArray(), j, Version)
            End Select
            
            p = 0
            If DataSet(k).Type = "POLYLINE" Then    '如果是“POLYLINE”命令
                'Endword = "ENDSEC" Else Endword = "0"
                Do While UCase(sArray(j)) <> "ENDSEC" And UCase(sArray(j)) <> "SEQEND" And UCase(sArray(j + 1)) <> "ENDSEC" And UCase(sArray(j + 1)) <> "SEQEND" And j + 1 <= UBound(sArray)
                    ReDim Preserve DataSet(k).data(p)
                    DataSet(k).data(p).Code = sArray(j)
                    DataSet(k).data(p).value = sArray(j + 1)
                    p = p + 1
                    j = j + 2
                Loop
            Else
            '其他命令的处理，如SPLINE
            '这里是关键步骤，将DXF文件的命令值传递到用户数据结构
                Do While UCase(sArray(j)) <> "0"
                    ReDim Preserve DataSet(k).data(p)
                    DataSet(k).data(p).Code = sArray(j)
                    DataSet(k).data(p).value = sArray(j + 1)
                    p = p + 1
                    j = j + 2
                    
                    If j + 1 > UBound(sArray) Then
                        Exit Do
                    End If
                Loop
            End If
            
            PrepareDXFEntity DataSet(k), EntitiesSection, Layers()
            k = k + 1
           
        End If
    Next j
    ParseDXFEntitiesSection = True
    Exit Function
    
ExitMe:
    MsgBox "ERROR  " & Err.Description
End Function

Sub PrepareDXFEntity(DataSet As DXF_DataSet, DXF As DXF_Data, Layers() As DXF_Layer, Optional isBlock As Boolean = False)
'将分析出的实体命令值转化为特征实体数据
    Dim I As Long, j As Long, k As Long, n As Long, p As Long
    Dim layer_name As String, color_index As Long, layer_id As Long, color_id As Long
    Dim dX As Double, dy As Double
    
    On Error GoTo eTrap
    
    dX = -DXF.Base.X
    dy = -DXF.Base.Y
    
'    If DataSet.Type <> "INSERT" And DataSet.Type <> "DIMENSION" Then
'        layer_name = DataSet.LayerName
'        color_index = DataSet.ColorIndex
'        If isBlock = False Then
'            layer_id = SetPathLayer(layer_name, True)
'            color_id = SetPathColor(color_index, True)
'        Else
'            layer_id = SetPathLayer(layer_name, False)
'            color_id = SetPathColor(color_index, False)
'        End If
'    End If

    I = 0
    j = 0
    k = 0
    n = 0
    p = 0
    
    Select Case DataSet.Type
        Case "POINT"
            k = UBound(DXF.points) + 1
            ReDim Preserve DXF.points(k) As DXF_Point
            DXF.points(k).X = GetGroupValue(DataSet.data(), 10) + dX
            DXF.points(k).Y = GetGroupValue(DataSet.data(), 20) + dy
            DXF.points(k).z = GetGroupValue(DataSet.data(), 30)
            'DXF.Points(k).layer_id = layer_id
            'DXF.Points(k).color_id = color_id
            
            DXF.PointCount = DXF.PointCount + 1
            
        Case "LINE"
            k = UBound(DXF.Lines) + 1
            ReDim Preserve DXF.Lines(k) As DXF_Line
            DXF.Lines(k).P1.X = GetGroupValue(DataSet.data(), 10) + dX
            DXF.Lines(k).P1.Y = GetGroupValue(DataSet.data(), 20) + dy
            DXF.Lines(k).P2.X = GetGroupValue(DataSet.data(), 11) + dX
            DXF.Lines(k).P2.Y = GetGroupValue(DataSet.data(), 21) + dy
            'DXF.Lines(k).layer_id = layer_id
            'DXF.Lines(k).color_id = color_id
            
            DXF.LineCount = DXF.LineCount + 1
            
        Case "ARC" ', "CIRCLE"
            k = UBound(DXF.Arcs) + 1
            ReDim Preserve DXF.Arcs(k) As DXF_Arc
            DXF.Arcs(k).Center.X = GetGroupValue(DataSet.data(), 10) + dX
            DXF.Arcs(k).Center.Y = GetGroupValue(DataSet.data(), 20) + dy
            DXF.Arcs(k).Center.z = GetGroupValue(DataSet.data(), 30)
            DXF.Arcs(k).radius = GetGroupValue(DataSet.data(), 40)
            DXF.Arcs(k).Angle1 = GetGroupValue(DataSet.data(), 50)
            DXF.Arcs(k).angle2 = GetGroupValue(DataSet.data(), 51)
            'DXF.Arcs(k).layer_id = layer_id
            'DXF.Arcs(k).color_id = color_id
            
            DXF.ArcCount = DXF.ArcCount + 1
            
        Case "CIRCLE"
            k = UBound(DXF.Arcs) + 1
            ReDim Preserve DXF.Arcs(k) As DXF_Arc
            DXF.Arcs(k).Center.X = GetGroupValue(DataSet.data(), 10) + dX
            DXF.Arcs(k).Center.Y = GetGroupValue(DataSet.data(), 20) + dy
            DXF.Arcs(k).Center.z = GetGroupValue(DataSet.data(), 30)
            DXF.Arcs(k).radius = GetGroupValue(DataSet.data(), 40)
            DXF.Arcs(k).Angle1 = 0
            DXF.Arcs(k).angle2 = 360
            'DXF.Arcs(k).layer_id = layer_id
            'DXF.Arcs(k).color_id = color_id
            
            DXF.ArcCount = DXF.ArcCount + 1
            
       Case "ELLIPSE"
            k = UBound(DXF.Ellipses) + 1
            ReDim Preserve DXF.Ellipses(k) As DXF_Ellipse
            
            DXF.Ellipses(k).Center.X = GetGroupValue(DataSet.data(), 10) + dX
            DXF.Ellipses(k).Center.Y = GetGroupValue(DataSet.data(), 20) + dy
            DXF.Ellipses(k).Center.z = GetGroupValue(DataSet.data(), 30)
            
            DXF.Ellipses(k).EndpointOfMajorAxis.X = GetGroupValue(DataSet.data(), 11) + dX
            DXF.Ellipses(k).EndpointOfMajorAxis.Y = GetGroupValue(DataSet.data(), 21) + dy
            DXF.Ellipses(k).EndpointOfMajorAxis.z = GetGroupValue(DataSet.data(), 31)
            
            DXF.Ellipses(k).RatioOfMinorAxisToMajor = GetGroupValue(DataSet.data(), 40)
            
            DXF.Ellipses(k).Angle1 = GetGroupValue(DataSet.data(), 41)
            DXF.Ellipses(k).angle2 = GetGroupValue(DataSet.data(), 42)
            'DXF.Ellipses(k).layer_id = layer_id
            'DXF.Ellipses(k).color_id = color_id
         
            DXF.EllipseCount = DXF.EllipseCount + 1
            
         Case "POLYLINE"
            k = UBound(DXF.PolyLines) + 1
            ReDim Preserve DXF.PolyLines(k) As DXF_PolyLine
            Do While p < UBound(DataSet.data)
                Do While DataSet.data(p).value <> "VERTEX" And p < UBound(DataSet.data)
                    p = p + 1
                Loop
                If p < UBound(DataSet.data) Then
                    ReDim Preserve DXF.PolyLines(k).Vertex(j) As DXF_Point
                    Do While DataSet.data(p).Code <> 10 And p < UBound(DataSet.data)
                        p = p + 1
                    Loop
                    DXF.PolyLines(k).Vertex(j).X = DataSet.data(p).value + dX
                    
                    Do While DataSet.data(p).Code <> 20 And p < UBound(DataSet.data)
                        p = p + 1
                    Loop
                    DXF.PolyLines(k).Vertex(j).Y = DataSet.data(p).value + dy
                    
                    Do While DataSet.data(p).Code <> 30 And p < UBound(DataSet.data)
                        p = p + 1
                    Loop
                    DXF.PolyLines(k).Vertex(j).z = DataSet.data(p).value
                    
                    Do While DataSet.data(p).Code <> 0 And p < UBound(DataSet.data)
                        p = p + 1
                    Loop
                    j = j + 1
                End If
            Loop
            'DXF.PolyLines(k).layer_id = layer_id
            'DXF.PolyLines(k).color_id = color_id
            
            DXF.PolyLineCount = DXF.PolyLineCount + 1
            
        Case "LWPOLYLINE"
            k = UBound(DXF.PolyLines) + 1
            ReDim Preserve DXF.PolyLines(k) As DXF_PolyLine
                        
            n = GetGroupValue(DataSet.data(), 90)
            j = GetGroupValue(DataSet.data(), 70)

            ReDim Preserve DXF.PolyLines(k).Vertex(n - 1) As DXF_Point
            
            For I = 0 To n - 1
                DXF.PolyLines(k).Vertex(I).X = GetGroupValue(DataSet.data(), 10, True) + dX
                DXF.PolyLines(k).Vertex(I).Y = GetGroupValue(DataSet.data(), 20, True) + dy
                DXF.PolyLines(k).Vertex(I).z = GetGroupValue(DataSet.data(), 30, True)
            Next
            
            If j = 1 And n > 2 Then 'closed
                ReDim Preserve DXF.PolyLines(k).Vertex(n) As DXF_Point
                DXF.PolyLines(k).Vertex(n) = DXF.PolyLines(k).Vertex(0)
            End If
            
            'DXF.PolyLines(k).layer_id = layer_id
            'DXF.PolyLines(k).color_id = color_id
            
            DXF.PolyLineCount = DXF.PolyLineCount + 1
    
        Case "SPLINE"
            Dim Knots() As Double   '节点值数组
            Dim cPts() As DXF_Point '控制点数组
            Dim fPts() As DXF_Point '拟合点数组
            
            Dim KCount As Long  '节点个数
            Dim CCount As Long  '控制点数
            Dim FCount As Long  '拟合点数
            
            
            Do While p < UBound(DataSet.data)
                If DataSet.data(p).Code = 40 Then 'KNots
                    ReDim Preserve Knots(I) As Double
                    Knots(I) = DataSet.data(p).value
                    I = I + 1
                End If
                If DataSet.data(p).Code = 10 Then 'Control Points
                    ReDim Preserve cPts(j) As DXF_Point
                    cPts(j).X = DataSet.data(p).value + dX
                    cPts(j).Y = DataSet.data(p + 1).value + dy
                    j = j + 1
                End If
                
                If DataSet.data(p).Code = 11 Then ' Fit Points
                    ReDim Preserve fPts(k) As DXF_Point
                    fPts(k).X = DataSet.data(p).value + dX
                    fPts(k).Y = DataSet.data(p + 1).value + dy
                    k = k + 1
                End If
                p = p + 1
            Loop
            
'
'            '-------------Make Spline-----------------------
'            p = 0
'            k = UBound(DXF.SPLines) + 1
'            ReDim Preserve DXF.SPLines(k) As DXF_SPLine
'            For j = 0 To UBound(fPts)
'                ReDim Preserve DXF.SPLines(k).Vertex(p) As DXF_Point
'                DXF.SPLines(k).Vertex(p) = fPts(j)
'                p = p + 1
'            Next j
'            'DXF.SPLines(k).layer_id = layer_id
'            'DXF.SPLines(k).color_id = color_id
'
'            DXF.SPLineCount = DXF.SPLineCount + 1
'            '------------------------------------------------

            KCount = I
            CCount = j
            FCount = k
            
            k = UBound(DXF.PolyLines) + 1
            ReDim Preserve DXF.PolyLines(k) As DXF_PolyLine
                        
            
            If CCount = 0 Then
                'Pts(0) = GetPoint(FP(0), Param)
                'MoveToEx Param.hDC, Pts(0).X, Pts(0).Y, Pts(1)
                'For I = 1 To FCount - 1
                '    Pts(0) = GetPoint(FP(I), Param)
                '    LineTo Param.hDC, Pts(0).X, Pts(0).Y
                'Next I
              
                n = FCount
   
                ReDim Preserve DXF.PolyLines(k).Vertex(n - 1) As DXF_Point
                
                For I = 0 To n - 1
                    DXF.PolyLines(k).Vertex(I).X = fPts(I).X + dX
                    DXF.PolyLines(k).Vertex(I).Y = fPts(I).Y + dy
                    DXF.PolyLines(k).Vertex(I).z = fPts(I).z
                Next
            Else
                'If (FCount = 0) Then
                '    I = 0
                '    While I < CCount - 4
                '        Pts(0) = GetPoint(cp(I), Param)
                '        Pts(1) = GetPoint(cp(I + 1), Param)
                '        Pts(2) = GetPoint(cp(I + 2), Param)
                '        Pts(3) = GetPoint(cp(I + 3), Param)
                '        PolyBezier Param.hDC, Pts(0), 4
                '        I = I + 4
                '    Wend
                'Else
                '    DrawNURBS Param, cp(), Knot(), CCount
                'End If

                '通过控制点，节点数，控制点数 推导SPLINE曲线
                GetNURBS cPts, Knots, CCount, DXF.PolyLines(k).Vertex
                
            End If


            DXF.PolyLineCount = DXF.PolyLineCount + 1
    
            
        Case "TEXT", "MTEXT"
        Case "INSERT"
            k = UBound(DXF.Inserts) + 1
            ReDim Preserve DXF.Inserts(k) As DXF_Insert
            DXF.Inserts(k).Name = GetGroupValue(DataSet.data(), 2)
            DXF.Inserts(k).Base.X = GetGroupValue(DataSet.data(), 10)
            DXF.Inserts(k).Base.Y = GetGroupValue(DataSet.data(), 20)
            DXF.Inserts(k).ScaleX = GetGroupValue(DataSet.data(), 41)
            DXF.Inserts(k).ScaleY = GetGroupValue(DataSet.data(), 42)
            DXF.Inserts(k).angle = GetGroupValue(DataSet.data(), 50)
            DXF.Inserts(k).ExtrusionDirZ = GetGroupValue(DataSet.data(), 230)
            'DXF.Inserts(k).layer_id = layer_id
            
            DXF.InsertCount = DXF.InsertCount + 1

        Case "DIMENSION"
            k = UBound(DXF.Inserts) + 1
            ReDim Preserve DXF.Inserts(k) As DXF_Insert
            DXF.Inserts(k).Name = GetGroupValue(DataSet.data(), 2)
            'DXF.Inserts(k).layer_id = layer_id
            
            DXF.InsertCount = DXF.InsertCount + 1
            
    End Select
    ReDim DataSet.data(0) As DXF_DataGroup
    Exit Sub
    
eTrap:
    k = 0
    Resume Next
End Sub


Function FindLayer(Layers() As DXF_Layer, Name As String)
    Dim I As Integer
    
    For I = 0 To UBound(Layers)
        If Layers(I).Name = Name Then
            FindLayer = I
            Exit Function
        End If
    Next I
    FindLayer = -1
End Function

Sub CreateInsertDXFData(DXF() As DXF_Data, CurInsert As DXF_Insert, NewDXF() As DXF_Data)
    On Error GoTo eTrap
    
    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim g As Integer
    Dim g2 As Integer
    Dim CPt As DXF_Point
    Dim tPoint As DXF_Point
    Dim tLine As DXF_Line
    Dim tArc As DXF_Arc
    Dim tEllipse As DXF_Ellipse
    Dim tSpline As DXF_SPLine
    Dim tPolyLine As DXF_PolyLine
    Dim tInsert As DXF_Insert
    
    If CurInsert.ScaleX = 0 Then CurInsert.ScaleX = 1
    If CurInsert.ScaleY = 0 Then CurInsert.ScaleY = 1
    
    k = FindGeo(DXF(), CurInsert.Name)
    ReDim NewDXF(UBound(DXF)) As DXF_Data
    NewDXF(0).Name = CurInsert.Name
    
    '-------------------Create the geometry
    j = UBound(DXF(k).points)
    g = 0
    For I = 0 To j
        tPoint = DXF(k).points(I)
        tPoint = ScalePoint(tPoint, CurInsert.ScaleX, CurInsert.ScaleY)
        tPoint = RotatePoint(tPoint, CPt, CurInsert.angle)
        tPoint = MovePoint(tPoint, CurInsert.Base.X, CurInsert.Base.Y)
        ReDim Preserve NewDXF(0).points(g) As DXF_Point
        NewDXF(0).points(g) = tPoint
        NewDXF(0).points(g).layer_id = CurInsert.layer_id
        g = g + 1
        
        NewDXF(0).PointCount = NewDXF(0).PointCount + 1
    Next I
    
    j = UBound(DXF(k).Lines)
    g = 0
    For I = 0 To j
        tLine = DXF(k).Lines(I)
        tLine = ScaleLine(tLine, CurInsert.ScaleX, CurInsert.ScaleY)
        tLine = RotateLine(tLine, CPt, CurInsert.angle)
        tLine = MoveLine(tLine, CurInsert.Base.X, CurInsert.Base.Y)
        ReDim Preserve NewDXF(0).Lines(g) As DXF_Line
        NewDXF(0).Lines(g) = tLine
        NewDXF(0).Lines(g).layer_id = CurInsert.layer_id
        g = g + 1
        
        NewDXF(0).LineCount = NewDXF(0).LineCount + 1
    Next I
    
    j = UBound(DXF(k).Arcs)
    g = 0
    g2 = 0
    For I = 0 To j
        tArc = DXF(k).Arcs(I)
        If CurInsert.ScaleX <> 1 Or CurInsert.ScaleY <> 1 Then
            tPolyLine = ArcToPolyLine(tArc)
            tPolyLine = ScalePolyLine(tPolyLine, CurInsert.ScaleX, CurInsert.ScaleY)
            tPolyLine = RotatePolyLine(tPolyLine, CPt, CurInsert.angle)
            tPolyLine = MovePolyLine(tPolyLine, CurInsert.Base.X, CurInsert.Base.Y)
            ReDim Preserve NewDXF(0).PolyLines(g2) As DXF_PolyLine
            NewDXF(0).PolyLines(g2) = tPolyLine
            NewDXF(0).PolyLines(g2).layer_id = CurInsert.layer_id
            g2 = g2 + 1
        Else
            tArc = RotateArc(tArc, CPt, CurInsert.angle)
            tArc = MoveArc(tArc, CurInsert.Base.X, CurInsert.Base.Y)
            ReDim Preserve NewDXF(0).Arcs(g) As DXF_Arc
            NewDXF(0).Arcs(g) = tArc
            NewDXF(0).Arcs(g).layer_id = CurInsert.layer_id
            g = g + 1
        End If
        
        NewDXF(0).ArcCount = NewDXF(0).ArcCount + 1
    Next I
    
    j = UBound(DXF(k).Ellipses)
    g = 0
    For I = 0 To j
        tEllipse = DXF(k).Ellipses(I)
        If CurInsert.ScaleX <> 1 Or CurInsert.ScaleY <> 1 Then
            tPolyLine = EllipseToPolyLine(tEllipse)
            tPolyLine = ScalePolyLine(tPolyLine, CurInsert.ScaleX, CurInsert.ScaleY)
            tPolyLine = RotatePolyLine(tPolyLine, CPt, CurInsert.angle)
            tPolyLine = MovePolyLine(tPolyLine, CurInsert.Base.X, CurInsert.Base.Y)
            ReDim Preserve NewDXF(0).PolyLines(g2) As DXF_PolyLine
            NewDXF(0).PolyLines(g2) = tPolyLine
            NewDXF(0).PolyLines(g2).layer_id = CurInsert.layer_id
            g2 = g2 + 1
        Else
            tEllipse = RotateEllipse(tEllipse, CPt, CurInsert.angle)
            tEllipse = MoveEllipse(tEllipse, CurInsert.Base.X, CurInsert.Base.Y)
            ReDim Preserve NewDXF(0).Ellipses(g) As DXF_Ellipse
            NewDXF(0).Ellipses(g) = tEllipse
            NewDXF(0).Ellipses(g).layer_id = CurInsert.layer_id
            g = g + 1
        End If
        
        NewDXF(0).EllipseCount = NewDXF(0).EllipseCount + 1
    Next I
    
    j = UBound(DXF(k).SPLines)
    g = 0
    For I = 0 To j
        tSpline = DXF(k).SPLines(I)
        tSpline = ScaleSpline(tSpline, CurInsert.ScaleX, CurInsert.ScaleY)
        tSpline = RotateSpline(tSpline, CPt, CurInsert.angle)
        tSpline = MoveSpline(tSpline, CurInsert.Base.X, CurInsert.Base.Y)
        ReDim Preserve NewDXF(0).SPLines(g) As DXF_SPLine
        NewDXF(0).SPLines(g) = tSpline
        NewDXF(0).SPLines(g).layer_id = CurInsert.layer_id
        g = g + 1
        
        NewDXF(0).SPLineCount = NewDXF(0).SPLineCount + 1
    Next I
    
    j = UBound(DXF(k).PolyLines)
    For I = 0 To j
        tPolyLine = DXF(k).PolyLines(I)
        tPolyLine = ScalePolyLine(tPolyLine, CurInsert.ScaleX, CurInsert.ScaleY)
        tPolyLine = RotatePolyLine(tPolyLine, CPt, CurInsert.angle)
        tPolyLine = MovePolyLine(tPolyLine, CurInsert.Base.X, CurInsert.Base.Y)
        
        ReDim Preserve NewDXF(0).PolyLines(g2) As DXF_PolyLine
        NewDXF(0).PolyLines(g2) = tPolyLine
        NewDXF(0).PolyLines(g2).layer_id = CurInsert.layer_id
        g2 = g2 + 1
        
        NewDXF(0).PolyLineCount = NewDXF(0).PolyLineCount + 1
    Next I
    
    j = UBound(DXF(k).Inserts)
    g = 0
    For I = 0 To j
        tInsert = DXF(k).Inserts(I)
        tInsert = ScaleInsert(tInsert, CurInsert.ScaleX, CurInsert.ScaleY)
        tInsert = RotateInsert(tInsert, CPt, CurInsert.angle)
        tInsert = MoveInsert(tInsert, CurInsert.Base.X, CurInsert.Base.Y)
        ReDim Preserve NewDXF(0).Inserts(g2) As DXF_Insert
        NewDXF(0).Inserts(g) = tInsert
        NewDXF(0).Inserts(g).layer_id = CurInsert.layer_id
        g = g + 1
        
        NewDXF(0).InsertCount = NewDXF(0).InsertCount + 1
    Next I
    
    For I = 1 To UBound(DXF)
        NewDXF(I) = DXF(I)
    Next I
    Exit Sub
    
eTrap:
    j = -1
    Resume Next
End Sub

Function FindGeo(DXF() As DXF_Data, Name As String)
    Dim I As Long
    
    For I = 0 To UBound(DXF)
        If DXF(I).Name = Name Then
            FindGeo = I
            Exit Function
        End If
    Next I
    FindGeo = -1
End Function

Function GetGroupValue(data() As DXF_DataGroup, Code As String, Optional ClearCode As Boolean = False) As Variant
    Dim I As Integer
    For I = 0 To UBound(data)
        If data(I).Code = Code Then
            GetGroupValue = data(I).value
            If ClearCode Then
                data(I).Code = 0
            End If
            Exit Function
        End If
    Next I
    GetGroupValue = 0
End Function

Function PtPtAngle(P1 As DXF_Point, P2 As DXF_Point) As Double
    Dim CurLine As DXF_Line
    
    CurLine.P1 = P1
    CurLine.P2 = P2
    PtPtAngle = cAngle(CurLine)
End Function

Function cAngle(CurLine As DXF_Line) As Double
    If CurLine.P1.X = CurLine.P2.X Then
        If CurLine.P1.Y < CurLine.P2.Y Then
            cAngle = 90
        Else
            cAngle = 270
        End If
        Exit Function
    ElseIf CurLine.P1.Y = CurLine.P2.Y Then
        If CurLine.P1.X < CurLine.P2.X Then
            cAngle = 0
        Else
            cAngle = 180
        End If
        Exit Function
    Else
        cAngle = Atn(CSlope(CurLine))
        cAngle = cAngle * 180 / Pi
        If cAngle < 0 Then cAngle = cAngle + 360
        
        '----------Test for direction--------
        If CurLine.P1.X > CurLine.P2.X And cAngle <> 180 Then cAngle = cAngle + 180
        If CurLine.P1.Y > CurLine.P2.Y And cAngle = 90 Then cAngle = cAngle + 180
        If cAngle > 360 Then cAngle = cAngle - 360
    End If
End Function

Function CSlope(CurLine As DXF_Line) As Single
    'if the line is VERTICAL we need to tweak the line so that the the slope is not UNDEFINED
    If CurLine.P1.X = CurLine.P2.X Then
        If CurLine.P1.Y < CurLine.P2.Y Then
            CSlope = 32000000000000#
        Else
            CSlope = -32000000000000#
        End If
    Else
        CSlope = (CurLine.P2.Y - CurLine.P1.Y) / (CurLine.P2.X - CurLine.P1.X)
    End If
End Function

Function PtLen(PtA As DXF_Point, PtB As DXF_Point) As Double
    PtLen = Sqr(Abs(PtB.Y - PtA.Y) ^ 2 + Abs(PtB.X - PtA.X) ^ 2)
End Function

Function cAngPt(angle As Double, pt As DXF_Point, h As Double) As DXF_Point
    cAngPt.X = h * Cos(angle * PI_180) + pt.X
    cAngPt.Y = h * Sin(angle * PI_180) + pt.Y
End Function

Function ScalePoint(CurPoint As DXF_Point, ScaleX As Double, ScaleY As Double) As DXF_Point
    ScalePoint = CurPoint
    ScalePoint.X = CurPoint.X * ScaleX
    ScalePoint.Y = CurPoint.Y * ScaleY
End Function

Function ScaleLine(CurLine As DXF_Line, ScaleX As Double, ScaleY As Double) As DXF_Line
    ScaleLine = CurLine
    ScaleLine.P1 = ScalePoint(CurLine.P1, ScaleX, ScaleY)
    ScaleLine.P2 = ScalePoint(CurLine.P2, ScaleX, ScaleY)
End Function

Function ScalePolyLine(CurPolyLine As DXF_PolyLine, ScaleX As Double, ScaleY As Double) As DXF_PolyLine
    Dim I As Long
    
    ScalePolyLine = CurPolyLine
    For I = 0 To UBound(CurPolyLine.Vertex)
        ScalePolyLine.Vertex(I) = ScalePoint(CurPolyLine.Vertex(I), ScaleX, ScaleY)
    Next I
End Function

Function ScaleSpline(CurSPline As DXF_SPLine, ScaleX As Double, ScaleY As Double) As DXF_SPLine
    Dim I As Long
    
    ScaleSpline = CurSPline
    For I = 0 To UBound(CurSPline.Vertex)
        ScaleSpline.Vertex(I) = ScalePoint(CurSPline.Vertex(I), ScaleX, ScaleY)
    Next I
End Function

Function ScaleInsert(CurInsert As DXF_Insert, ScaleX As Double, ScaleY As Double) As DXF_Insert
    ScaleInsert = CurInsert
    ScaleInsert.ScaleX = CurInsert.ScaleX * ScaleX
    ScaleInsert.ScaleY = CurInsert.ScaleY * ScaleY
    ScaleInsert.Base = ScalePoint(CurInsert.Base, ScaleX, ScaleY)
End Function

Function RotatePoint(CurPoint As DXF_Point, Pivot As DXF_Point, angle As Double) As DXF_Point
    RotatePoint = CurPoint
    RotatePoint.X = RotX(CurPoint.X - Pivot.X, CurPoint.Y - Pivot.Y, angle) + Pivot.X
    RotatePoint.Y = RotY(CurPoint.X - Pivot.X, CurPoint.Y - Pivot.Y, angle) + Pivot.Y
End Function

Function RotateLine(CurLine As DXF_Line, Pivot As DXF_Point, angle As Double) As DXF_Line
    RotateLine = CurLine
    RotateLine.P1 = RotatePoint(CurLine.P1, Pivot, angle)
    RotateLine.P2 = RotatePoint(CurLine.P2, Pivot, angle)
End Function

Function RotateSpline(CurSPline As DXF_SPLine, Pivot As DXF_Point, angle As Double) As DXF_SPLine
    Dim I As Long
    
    RotateSpline = CurSPline
    For I = 0 To UBound(CurSPline.Vertex)
        RotateSpline.Vertex(I) = RotatePoint(CurSPline.Vertex(I), Pivot, angle)
    Next I
End Function

Function RotatePolyLine(CurPolyLine As DXF_PolyLine, Pivot As DXF_Point, angle As Double) As DXF_PolyLine
    Dim I As Long
    
    RotatePolyLine = CurPolyLine
    For I = 0 To UBound(CurPolyLine.Vertex)
        RotatePolyLine.Vertex(I) = RotatePoint(CurPolyLine.Vertex(I), Pivot, angle)
    Next I
End Function

Function RotateArc(CurArc As DXF_Arc, Pivot As DXF_Point, angle As Double) As DXF_Arc
    RotateArc = CurArc
    RotateArc.Center = RotatePoint(CurArc.Center, Pivot, angle)
    RotateArc.Angle1 = dAngle(RotateArc.Angle1 + angle)
    RotateArc.angle2 = dAngle(RotateArc.angle2 + angle)
End Function

Function RotateEllipse(CurEllipse As DXF_Ellipse, Pivot As DXF_Point, angle As Double) As DXF_Ellipse
'    RotateEllipse = CurEllipse
'    RotateEllipse.F1 = RotatePoint(CurEllipse.F1, Pivot, Angle)
'    RotateEllipse.F2 = RotatePoint(CurEllipse.F2, Pivot, Angle)
'    RotateEllipse.P1 = RotatePoint(CurEllipse.P1, Pivot, Angle)
    'RotateEllipse.Theta = dAngle(RotateEllipse.Theta + Angle)
    'RotateEllipse.Hyp = RotatePoint(Myellipse.Hyp, Myellipse.Center, angle)
    'RotateEllipse.Angle1 = (EllipseAngle(RotateEllipse.Angle1 + angle, RotateEllipse))
    'RotateEllipse.Angle2 = (EllipseAngle(RotateEllipse.Angle2 + angle, RotateEllipse))
End Function

Function RotateInsert(CurInsert As DXF_Insert, Pivot As DXF_Point, angle As Double) As DXF_Insert
    RotateInsert = CurInsert
    RotateInsert.angle = CurInsert.angle + angle
    RotateInsert.Base = RotatePoint(CurInsert.Base, Pivot, angle)

End Function

Function MovePoint(CurPoint As DXF_Point, dX As Double, dy As Double) As DXF_Point
    MovePoint = CurPoint
    MovePoint.X = CurPoint.X + dX
    MovePoint.Y = CurPoint.Y + dy
End Function

Function MoveLine(CurLine As DXF_Line, dX As Double, dy As Double) As DXF_Line
    MoveLine = CurLine
    MoveLine.P1 = MovePoint(CurLine.P1, dX, dy)
    MoveLine.P2 = MovePoint(CurLine.P2, dX, dy)
End Function

Function MoveArc(CurArc As DXF_Arc, dX As Double, dy As Double) As DXF_Arc
    MoveArc = CurArc
    MoveArc.Center = MovePoint(CurArc.Center, dX, dy)
End Function

Function MoveEllipse(CurEllipse As DXF_Ellipse, dX As Double, dy As Double) As DXF_Ellipse
    MoveEllipse = CurEllipse
'    MoveEllipse.F1 = MovePoint(CurEllipse.F1, dX, dY)
'    MoveEllipse.F2 = MovePoint(CurEllipse.F2, dX, dY)
'    MoveEllipse.P1 = MovePoint(CurEllipse.P1, dX, dY)
End Function

Function MoveSpline(CurSPline As DXF_SPLine, dX As Double, dy As Double) As DXF_SPLine
    Dim I As Long
    
    MoveSpline = CurSPline
    For I = 0 To UBound(MoveSpline.Vertex)
        MoveSpline.Vertex(I) = MovePoint(CurSPline.Vertex(I), dX, dy)
    Next I
End Function

Function MovePolyLine(CurPolyLine As DXF_PolyLine, dX As Double, dy As Double) As DXF_PolyLine
    Dim I As Long
    
    MovePolyLine = CurPolyLine
    For I = 0 To UBound(MovePolyLine.Vertex)
        MovePolyLine.Vertex(I) = MovePoint(CurPolyLine.Vertex(I), dX, dy)
    Next I
End Function

Function MoveInsert(CurInsert As DXF_Insert, dX As Double, dy As Double) As DXF_Insert
    MoveInsert = CurInsert
    MoveInsert.Base.X = CurInsert.Base.X + dX
    MoveInsert.Base.Y = CurInsert.Base.Y + dy
End Function

Function ArcToPolyLine(CurArc As DXF_Arc) As DXF_PolyLine
    On Local Error GoTo eTrap
    Dim I As Long
    Dim p As Double
    Dim cx As Double
    Dim cy As Double
    Dim rad As Double
    Dim ang1 As Double
    Dim ang2 As Double
    Dim aLen As Double
    Dim div As Integer
    
    rad = CurArc.radius
    ang1 = CurArc.Angle1
    ang2 = CurArc.angle2
    
    If ang2 < ang1 Then
        ang2 = ang2 + 360
    End If
    
    div = (ang2 - ang1) / 8
    If div = 0 Then div = 1
    
    cx = CurArc.Center.X
    cy = CurArc.Center.Y
    
    ReDim ArcToPolyLine.Vertex(div + 1) As DXF_Point
    ArcToPolyLine.layer_id = CurArc.layer_id
    aLen = (ang2 - ang1) / (div)
    
    For I = 0 To div
        p = ang1 + (I * aLen)
        ArcToPolyLine.Vertex(I).X = (rad * Cos(p * PI_180)) + cx
        ArcToPolyLine.Vertex(I).Y = (rad * Sin(p * PI_180)) + cy
    Next I
    
    ArcToPolyLine.Vertex(I).X = (rad * Cos(ang2 * PI_180)) + cx
    ArcToPolyLine.Vertex(I).Y = (rad * Sin(ang2 * PI_180)) + cy
    'For p = MyArc.Angle1 To MyArc.Angle2 Step aLen
    '    ArcToPolyLine.Vertex(i).X = (Rad * Cos(p * Pi / 180)) + cx
    '    ArcToPolyLine.Vertex(i).Y = (Rad * Sin(p * Pi / 180)) + cy
    '    i = i + 1
    'Next p
    Exit Function
eTrap:
    MsgBox Err.Description
    Resume Next
End Function

Function EllipseToPolyLine(CurEllipse As DXF_Ellipse) As DXF_PolyLine
    Dim a As Double, b As Double ' ellipse parameters
    Dim u As Double, v As Double
    Dim I As Double, j As Double
    Dim ang1 As Double, ang2 As Double 'adjusted ellipse angles
    Dim x1 As Double, y1 As Double 'plotting coordinates
    Dim cx As Double, cy As Double ' center of the ellipse
    Dim First As Boolean
    Dim TotLen As Double, FLen As Double, CosJ As Double
    Dim vCount As Long
    
Exit Function
    '-----------------Angle adjustment-------
    EllipseToPolyLine.layer_id = CurEllipse.layer_id
    
    ang1 = CurEllipse.Angle1
    ang2 = CurEllipse.angle2
    If ang2 < ang1 Then ang2 = ang2 + 360
    I = (ang2 - ang1) / 60 'CurEllipse.NumPoints
    
    cx = CurEllipse.Center.X
    cy = CurEllipse.Center.Y
    
    'v = PtPtAngle(CurEllipse.F1, CurEllipse.F2) * PI_180
    'FLen = PtLen(CurEllipse.F1, CurEllipse.F2)
    'TotLen = PtLen(CurEllipse.F1, CurEllipse.P1) + PtLen(CurEllipse.F2, CurEllipse.P1)
    
    a = TotLen / 2
    b = Sqr((TotLen / 2) ^ 2 - (FLen / 2) ^ 2)
    
    For u = ang1 To ang2 Step I
        j = u * PI_180
        CosJ = Cos(j)
        x1 = Cos(Atn(b * Tan(j) / a) + v) * Sqr((a ^ 2 - b ^ 2) * CosJ ^ 2 + b ^ 2) * Sgn(a * CosJ)
        y1 = Sin(Atn(b * Tan(j) / a) + v) * Sqr((a ^ 2 - b ^ 2) * CosJ ^ 2 + b ^ 2) * Sgn(a * CosJ)
        
        ReDim Preserve EllipseToPolyLine.Vertex(vCount) As DXF_Point
        
        EllipseToPolyLine.Vertex(vCount).X = cx + x1
        EllipseToPolyLine.Vertex(vCount).Y = cy + y1
        vCount = vCount + 1
    Next u
    
    If u - I < ang2 Then
        j = ang2 * PI_180
        CosJ = Cos(j)
        x1 = Cos(Atn(b * Tan(j) / a) + v) * Sqr((a ^ 2 - b ^ 2) * CosJ ^ 2 + b ^ 2) * Sgn(a * CosJ)
        y1 = Sin(Atn(b * Tan(j) / a) + v) * Sqr((a ^ 2 - b ^ 2) * CosJ ^ 2 + b ^ 2) * Sgn(a * CosJ)
        
        ReDim Preserve EllipseToPolyLine.Vertex(vCount) As DXF_Point
        
        EllipseToPolyLine.Vertex(vCount).X = cx + x1
        EllipseToPolyLine.Vertex(vCount).Y = cy + y1
    End If
End Function

Function RotX(x1 As Double, y1 As Double, angle As Double) As Double
    RotX = cHyp(x1, y1) * Cos((PtAng(x1, y1) + angle) * PI_180)
End Function

Function RotY(x1 As Double, y1 As Double, angle As Double) As Single
    RotY = cHyp(x1, y1) * Sin((PtAng(x1, y1) + angle) * PI_180)
End Function

Function dAngle(angle As Double) As Double
    If angle > 360 Then
        dAngle = angle - 360
    ElseIf angle < 0 Then
        dAngle = angle + 360
    Else
        dAngle = angle
    End If
End Function

Function cHyp(x1 As Double, y1 As Double) As Double
    cHyp = Sqr((x1 * x1) + (y1 * y1))
End Function

Function PtAng(x1 As Double, y1 As Double) As Double
    If x1 = 0 Then
        If y1 >= 0 Then
            PtAng = 90
        Else
            PtAng = 270
        End If
        Exit Function
        
    ElseIf y1 = 0 Then
        If x1 >= 0 Then
            PtAng = 0
        Else
            PtAng = 180
        End If
        Exit Function
        
    Else
        PtAng = Atn(y1 / x1)
        PtAng = PtAng * 180 / Pi
        If PtAng < 0 Then PtAng = PtAng + 360
        If PtAng > 360 Then PtAng = PtAng - 360
        '----------Test for direction-(quadrant check)-------
        If x1 < 0 Then PtAng = PtAng + 180
        If y1 < 0 And PtAng < 90 Then PtAng = PtAng + 180
        'If X1 < 0 And PtAng <> 180 Then PtAng = PtAng + 180
        'If Y1 < 0 And PtAng = 90 Then PtAng = PtAng + 180
        
        'One final check
        If PtAng < 0 Then PtAng = PtAng + 360
        If PtAng > 360 Then PtAng = PtAng - 360
    End If
End Function

Function GetDistance(x0 As Double, x1 As Double, y0 As Double, y1 As Double, z0 As Double, z1 As Double) As Double
    GetDistance = Sqr((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0) + (z1 - z0) * (z1 - z0))
End Function

'Sub CatchOrAddPoint(p As Path_Point)
'    Dim k As Integer, I As Long
'
'    k = 0
'    For I = 1 To PointCount
'        If Abs(PointList(I).X - p.X) < 0.1 Then
'            If GetDistance(PointList(I).X, p.X, PointList(I).Y, p.Y, PointList(I).z, p.z) < 0.1 Then
'                k = 1
'                p.id = PointList(I).id
'                Exit For
'            End If
'        End If
'    Next
'
'    If k = 0 Then
'        AddPoint p.X, p.Y, p.z, p.Layer, PointType.NormalPoint
'        p.id = PointCount
'    End If
'End Sub

Sub CatchOrAddPoint(p As Path_Point)
    Dim I As Long, j As Long, k As Integer, q As Long
    
    q = 0
    For I = 1 To PointCount
        If Abs(PointList(I).X - p.X) < 0.01 Then
            If Abs(PointList(I).Y - p.Y) < 0.01 Then
                If GetDistance(PointList(I).X, p.X, PointList(I).Y, p.Y, PointList(I).z, p.z) < 0.01 Then
                    q = 1
                    p.id = PointList(I).id
                    Exit For
                End If
            End If
        End If
    Next
    
    If q = 0 Then
        AddPoint p.X, p.Y, p.z, p.Layer, PointType.NormalPoint
        PointList(PointCount).color = p.color
        'PointList(PointCount).insert_id = p.insert_id
        p.id = PointCount
    End If
End Sub



Function IsAStartPoint(ByVal pid As Long, ByRef id As Long, ByRef id_type As Long) As Boolean
    Dim I As Long, k As Boolean
    
    k = False
    For I = 1 To SegmentCount
        If SegmentList(I).point0_id = pid Then
            k = True
            id = I
            id_type = 0
            Exit For
        End If
    Next
    If k = False Then
        For I = 1 To ArcCount
            If ArcList(I).point0_id = pid Then
                k = True
                id = I
                id_type = 1
                Exit For
            End If
        Next
    End If
    If k = False Then
        For I = 1 To SPLineCount
            If SPLineList(I).point0_id = pid Then
                k = True
                id = I
                id_type = 2
                Exit For
            End If
        Next
    End If
            
    IsAStartPoint = k
End Function

Function IsAEndPoint(ByVal pid As Long, ByRef id As Long, ByRef id_type As Long) As Boolean
    Dim I As Long, k As Boolean
    
    k = False
    For I = 1 To SegmentCount
        If SegmentList(I).point1_id = pid Then
            k = True
            id = I
            id_type = 0
            Exit For
        End If
    Next
    If k = False Then
        For I = 1 To ArcCount
            If ArcList(I).point1_id = pid Then
                k = True
                id = I
                id_type = 1
                Exit For
            End If
        Next
    End If
    If k = False Then
        For I = 1 To SPLineCount
            If SPLineList(I).point1_id = pid Then
                k = True
                id = I
                id_type = 2
                Exit For
            End If
        Next
    End If
            
    IsAEndPoint = k
End Function


Sub ConvertDXFToCMP(ByVal insert_id As Long)
    Dim Point0 As Path_Point, Point1 As Path_Point, Pointm As Path_Point
    Dim cx As Double, cy As Double, cz As Double, r As Double, r2 As Double, Angle0 As Double, Angle1 As Double
    Dim I As Long, k As Integer, n As Long, id0 As Long, ux As Double, uy As Double, uz As Double, d As Double, t As Long
    
    Dim DXF As DXF_Data
    
    On Error Resume Next
    
    If insert_id = 0 Then
        DXF = DXFData(0) 'ENTITIES section
    Else
        DXF = NewDXFData(0) 'Created by "Insert"
    End If
    
    If DXF.PointCount > 0 Then
'Debug.Print "DXF.PointCount ="; DXF.PointCount

        For I = 0 To UBound(DXF.points)
            Point0.X = DXF.points(I).X
            Point0.Y = DXF.points(I).Y
            Point0.z = DXF.points(I).z
            Point0.Layer = 1 'DXF.Points(I).layer_id
            'Point0.color = LayerColor(PathColor(DXF.Points(I).color_id).Index - 1)
            'Point0.insert_id = insert_id
            
            '-------------------------------------------------------------
            CatchOrAddPoint Point0
        Next I
    End If
    
    If insert_id = 0 Then
        FrmMain.ProgressBar.Min = 0
        FrmMain.ProgressBar.Max = Max(DXF.LineCount + DXF.ArcCount + DXF.EllipseCount + DXF.PolyLineCount + DXF.SPLineCount + DXF.InsertCount, 1)
    End If
    
    If DXF.LineCount > 0 Then
'Debug.Print "DXF.LineCount ="; DXF.LineCount

        For I = 0 To UBound(DXF.Lines)
            If insert_id = 0 Then
                FrmMain.ProgressBar.value = I + 1
                FrmMain.ProgressBar.Refresh
            End If
    
            Point0.X = DXF.Lines(I).P1.X
            Point0.Y = DXF.Lines(I).P1.Y
            Point0.z = DXF.Lines(I).P1.z
            Point0.Layer = 1 'DXF.Lines(I).layer_id
            'Point0.color = LayerColor(PathColor(DXF.Lines(I).color_id).Index - 1)
            'Point0.insert_id = insert_id
    
            Point1.X = DXF.Lines(I).P2.X
            Point1.Y = DXF.Lines(I).P2.Y
            Point1.z = DXF.Lines(I).P2.z
            Point1.Layer = 1 'DXF.Lines(I).layer_id
            'Point1.color = LayerColor(PathColor(DXF.Lines(I).color_id).Index - 1)
            'Point1.insert_id = insert_id
                
                '-------------------------------------------------------------
            If Point0.X <> Point1.X Or Point0.Y <> Point1.Y Then
                CatchOrAddPoint Point0
                CatchOrAddPoint Point1
        
                If Point0.id <> Point1.id Then
                    AddSegment Point0.id, Point1.id
                    SegmentList(SegmentCount).Layer = 1 'DXF.Lines(I).layer_id
                    'SegmentList(SegmentCount).color = LayerColor(PathColor(DXF.Lines(I).color_id).Index - 1)
                    'SegmentList(SegmentCount).insert_id = insert_id
                End If
            End If
        Next I
    End If
    
    If DXF.ArcCount > 0 Then
'Debug.Print "DXF.ArcCount ="; DXF.ArcCount

        For I = 0 To UBound(DXF.Arcs)
            If insert_id = 0 Then
                FrmMain.ProgressBar.value = DXF.LineCount + I + 1
                FrmMain.ProgressBar.Refresh
            End If
    
            cx = DXF.Arcs(I).Center.X
            cy = DXF.Arcs(I).Center.Y
            cz = DXF.Arcs(I).Center.z
            
            r = DXF.Arcs(I).radius
            Angle0 = DXF.Arcs(I).Angle1 * PI_180
            Angle1 = DXF.Arcs(I).angle2 * PI_180
            
            '-------------------------------------------------------------
            If Angle0 > Angle1 Then
                If Angle0 >= Pi Then
                    Angle0 = Angle0 - PI2
                ElseIf Angle1 < Pi Then
                    Angle1 = Angle1 + PI2
                End If
            End If
            
            Point0.X = r * Cos(Angle0) + cx
            Point0.Y = r * Sin(Angle0) + cy
            Point0.z = cz
            Point0.Layer = 1 'DXF.Arcs(I).layer_id
            'Point0.color = LayerColor(PathColor(DXF.Arcs(I).color_id).Index - 1)
            'Point0.insert_id = insert_id
            
            Point1.X = r * Cos(Angle1) + cx
            Point1.Y = r * Sin(Angle1) + cy
            Point1.z = cz
            Point1.Layer = 1 'DXF.Arcs(I).layer_id
            'Point1.color = LayerColor(PathColor(DXF.Arcs(I).color_id).Index - 1)
            'Point1.insert_id = insert_id
            
            Pointm.X = cx
            Pointm.Y = cy
            Pointm.z = cz
            Pointm.Layer = 1 'DXF.Arcs(I).layer_id
            'Pointm.color = LayerColor(PathColor(DXF.Arcs(I).color_id).Index - 1)
            'Pointm.insert_id = insert_id
            
            CatchOrAddPoint Point0
            CatchOrAddPoint Point1
            CatchOrAddPoint Pointm
            
            AddArc cx, cy, cz, r, r, Angle0, Angle1, Point0.id, Point1.id, Pointm.id, 1, ArcType.CircleCR
            'ArcList(ArcCount).color = LayerColor(PathColor(DXF.Arcs(I).color_id).Index - 1)
            'ArcList(ArcCount).insert_id = insert_id
            
            PointList(Point0.id).body_id = ArcList(ArcCount).body_id
            PointList(Point1.id).body_id = ArcList(ArcCount).body_id
            PointList(Pointm.id).body_id = ArcList(ArcCount).body_id
            
            PointList(Point0.id).group_id = ArcList(ArcCount).group_id
            PointList(Point1.id).group_id = ArcList(ArcCount).group_id
            PointList(Pointm.id).group_id = ArcList(ArcCount).group_id
            
            PointList(Point0.id).Type = PointType.ArcPoint
            PointList(Point1.id).Type = PointType.ArcPoint
            PointList(Pointm.id).Type = PointType.ArcPoint
        Next
    End If
    
    If DXF.EllipseCount > 0 Then
'Debug.Print "DXF.EllipseCount ="; DXF.EllipseCount

        For I = 0 To UBound(DXF.Ellipses)
            If insert_id = 0 Then
                FrmMain.ProgressBar.value = DXF.LineCount + DXF.ArcCount + I + 1
                FrmMain.ProgressBar.Refresh
            End If
    
            cx = DXF.Ellipses(I).Center.X
            cy = DXF.Ellipses(I).Center.Y
            cz = DXF.Ellipses(I).Center.z
            
            ux = DXF.Ellipses(I).EndpointOfMajorAxis.X
            uy = DXF.Ellipses(I).EndpointOfMajorAxis.Y
            uz = DXF.Ellipses(I).EndpointOfMajorAxis.z
            
            r = Sqr(ux * ux + uy * uy)
            r2 = DXF.Ellipses(I).RatioOfMinorAxisToMajor * r
            
            Angle0 = DXF.Ellipses(I).Angle1 * PI_180
            Angle1 = DXF.Ellipses(I).angle2 * PI_180
            
            '-------------------------------------------------------------
            If Angle0 > Angle1 Then
                If Angle0 >= Pi Then
                    Angle0 = Angle0 - PI2
                ElseIf Angle1 < Pi Then
                    Angle1 = Angle1 + PI2
                End If
            End If
            
            Point0.X = r * Cos(Angle0) + cx
            Point0.Y = r2 * Sin(Angle0) + cy
            Point0.z = cz
            Point0.Layer = 1 'DXF.Ellipses(I).layer_id
            'Point0.color = LayerColor(PathColor(DXF.Ellipses(I).color_id).Index - 1)
            'Point0.insert_id = insert_id
            
            Point1.X = r * Cos(Angle1) + cx
            Point1.Y = r2 * Sin(Angle1) + cy
            Point1.z = cz
            Point1.Layer = 1 'DXF.Ellipses(I).layer_id
            'Point1.color = LayerColor(PathColor(DXF.Ellipses(I).color_id).Index - 1)
            'Point1.insert_id = insert_id
            
            Pointm.X = cx
            Pointm.Y = cy
            Pointm.z = cz
            Pointm.Layer = 1 'DXF.Arcs(I).layer_id
            'Pointm.color = LayerColor(PathColor(DXF.Ellipses(I).color_id).Index - 1)
            'Pointm.insert_id = insert_id
            
            CatchOrAddPoint Point0
            CatchOrAddPoint Point1
            CatchOrAddPoint Pointm
            
            AddArc cx, cy, cz, r, r, Angle0, Angle1, Point0.id, Point1.id, Pointm.id, 1, ArcType.CircleCR
            'ArcList(ArcCount).color = LayerColor(PathColor(DXF.Ellipses(I).color_id).Index - 1)
            'ArcList(ArcCount).insert_id = insert_id
            
            PointList(Point0.id).body_id = ArcList(ArcCount).body_id
            PointList(Point1.id).body_id = ArcList(ArcCount).body_id
            PointList(Pointm.id).body_id = ArcList(ArcCount).body_id
            
            PointList(Point0.id).group_id = ArcList(ArcCount).group_id
            PointList(Point1.id).group_id = ArcList(ArcCount).group_id
            PointList(Pointm.id).group_id = ArcList(ArcCount).group_id
                        
            PointList(Point0.id).Type = PointType.ArcPoint
            PointList(Point1.id).Type = PointType.ArcPoint
            PointList(Pointm.id).Type = PointType.ArcPoint
        Next I
    End If
    
    If DXF.SPLineCount > 0 Then
'Debug.Print "DXF.SPLineCount ="; DXF.SPLineCount

        For I = 0 To UBound(DXF.SPLines)
            If insert_id = 0 Then
                FrmMain.ProgressBar.value = DXF.LineCount + DXF.ArcCount + DXF.EllipseCount + I + 1
                FrmMain.ProgressBar.Refresh
            End If
    
            n = UBound(DXF.SPLines(I).Vertex())
'Debug.Print I; "n="; n
            If n > 0 Then
                AddSPLine
        
                For k = 0 To n
                    ux = DXF.SPLines(I).Vertex(k).X
                    uy = DXF.SPLines(I).Vertex(k).Y
                    uz = DXF.SPLines(I).Vertex(k).z
        
                    If k = 0 Then
                        Point0.X = ux
                        Point0.Y = uy
                        Point0.z = uz
                        Point0.Layer = 1 'DXF.SPLines(I).layer_id
                        'Point0.color = LayerColor(PathColor(DXF.SPLines(I).color_id).Index - 1)
                        'Point0.insert_id = insert_id
                        
                        CatchOrAddPoint Point0
                        PointList(Point0.id).Type = PointType.SPLinePoint
                        'PointList(Point0.id).insert_id = insert_id
                        'PointList(Point0.id).group_id = BodyCount 'SPLineCount
                        id0 = Point0.id
                    Else
                        Point1.X = ux
                        Point1.Y = uy
                        Point1.z = uz
                        Point1.Layer = 1 'DXF.SPLines(I).layer_id
                        'Point1.color = LayerColor(PathColor(DXF.SPLines(I).color_id).Index - 1)
                        'Point1.insert_id = insert_id
                        
                        CatchOrAddPoint Point1
                        PointList(Point1.id).Type = PointType.SPLinePoint
                        'PointList(Point1.id).insert_id = insert_id
                        'PointList(Point1.id).group_id = BodyCount 'SPLineCount
                        id0 = Point1.id
                    End If
                        
                    ReDim Preserve SPLineList(SPLineCount).vertex_id(k)
                    SPLineList(SPLineCount).vertex_id(k) = id0
                Next
                
                SPLineList(SPLineCount).point0_id = Point0.id
                SPLineList(SPLineCount).point1_id = Point1.id
                SPLineList(SPLineCount).Layer = 1 'DXF.SPLines(I).layer_id
                'SPLineList(SPLineCount).color = LayerColor(PathColor(DXF.SPLines(I).color_id).Index - 1)
                'SPLineList(SPLineCount).insert_id = insert_id
                SPLineList(SPLineCount).vertex_count = n + 1
            End If
        Next I
    End If
    
    If DXF.PolyLineCount > 0 Then
'Debug.Print "DXF.PolyLineCount ="; DXF.PolyLineCount

        For I = 0 To UBound(DXF.PolyLines)
            If insert_id = 0 Then
                FrmMain.ProgressBar.value = DXF.LineCount + DXF.ArcCount + DXF.EllipseCount + DXF.SPLineCount + I + 1
                FrmMain.ProgressBar.Refresh
            End If
    
            n = UBound(DXF.PolyLines(I).Vertex())
            If n > 0 Then
                For k = 0 To n
                    If k = 0 Then
                        Point0.X = DXF.PolyLines(I).Vertex(k).X
                        Point0.Y = DXF.PolyLines(I).Vertex(k).Y
                        Point0.z = DXF.PolyLines(I).Vertex(k).z
                        Point0.Layer = 1 'DXF.PolyLines(I).layer_id
                        'Point0.color = LayerColor(PathColor(DXF.PolyLines(I).color_id).Index - 1)
                        'Point0.insert_id = insert_id
                        
                        CatchOrAddPoint Point0
                        id0 = Point0.id
                    Else
                        t = 0
                        If k < n Then
                            cx = DXF.PolyLines(I).Vertex(k).X
                            cy = DXF.PolyLines(I).Vertex(k).Y
                            ux = DXF.PolyLines(I).Vertex(k + 1).X
                            uy = DXF.PolyLines(I).Vertex(k + 1).Y
                            
                            d = Sqr((cx - PointList(id0).X) * (cx - PointList(id0).X) + (cy - PointList(id0).Y) * (cy - PointList(id0).Y))
                            Angle0 = GetAngle(PointList(id0).X, PointList(id0).Y, cx, cy, ux, uy)
                            If d > 5 Or Abs(Angle0) > 3 Then
                                t = 1
                            End If
                        Else
                            t = 1
                        End If
                        
                        If t = 1 Then
                            Point1.X = DXF.PolyLines(I).Vertex(k).X
                            Point1.Y = DXF.PolyLines(I).Vertex(k).Y
                            Point1.z = DXF.PolyLines(I).Vertex(k).z
                            Point1.Layer = 1 'DXF.PolyLines(I).layer_id
                            'Point1.color = LayerColor(PathColor(DXF.PolyLines(I).color_id).Index - 1)
                            'Point1.insert_id = insert_id
                            
                            If PointList(id0).X <> Point1.X Or PointList(id0).Y <> Point1.Y Then
                                CatchOrAddPoint Point1
                                
                                If id0 <> Point1.id Then
                                    AddSegment id0, Point1.id
                                    
                                    SegmentList(SegmentCount).Layer = 1 'DXF.PolyLines(I).layer_id
                                    'SegmentList(SegmentCount).color = LayerColor(PathColor(DXF.PolyLines(I).color_id).Index - 1)
                                    'SegmentList(SegmentCount).insert_id = insert_id
                                    id0 = Point1.id
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next I
    End If
    
    If insert_id = 0 Then
        If DXF.InsertCount > 0 Then
'Debug.Print "DXF.InsertCount ="; DXF.InsertCount

            For I = 0 To UBound(DXF.Inserts)
                FrmMain.ProgressBar.value = DXF.LineCount + DXF.ArcCount + DXF.EllipseCount + DXF.SPLineCount + DXF.PolyLineCount + I + 1
                FrmMain.ProgressBar.Refresh
    
                CreateInsertDXFData DXFData(), DXF.Inserts(I), NewDXFData()
                ConvertDXFToCMP I + 1
            Next I
        End If
    End If
    
End Sub


'=====================================================================================
Function NN(ByVal n As Long, ByVal I As Long, ByVal t As Double, Knots() As Double)
    Dim v1 As Double, d1 As Double, v2 As Double, d2 As Double

    If n = 0 Then
        If (Knots(I) <= t) And (t < Knots(I + 1)) Then
            NN = 1
        Else
            NN = 0
        End If
        Exit Function
    End If
    
    d1 = (Knots(I + n) - Knots(I))
    v1 = (t - Knots(I)) * NN(n - 1, I, t, Knots)
    If d1 = 0 Then
        v1 = 0
    Else
        v1 = v1 / d1
    End If

    d2 = (Knots(I + n + 1) - Knots(I + 1))
    v2 = (Knots(I + n + 1) - t) * NN(n - 1, I + 1, t, Knots)
    If d2 = 0 Then
        v2 = 0
    Else
        v2 = v2 / d2
    End If
    NN = v1 + v2
End Function

Function NURBS_3(DP() As DXF_Point, Knots() As Double, ByVal j As Long, ByVal t As Double) As DXF_Point
    Dim r As DXF_Point
    Dim I As Long
    Dim Ni As Double

    r.X = 0
    r.Y = 0
    r.z = 0

    For I = j - 3 To j
        Ni = NN(3, I, t, Knots)
        r.X = r.X + DP(I).X * Ni
        r.Y = r.Y + DP(I).Y * Ni
        r.z = r.z + DP(I).z * Ni
    Next
    NURBS_3 = r
End Function

Sub GetNURBS(DP() As DXF_Point, Knots() As Double, ByVal Count As Long, Vertex() As DXF_Point)
'通过控制点，节点数，控制点数 推导SPLINE曲线
    Dim t As Double
    Dim StepT As Double
    Dim j As Long
    Dim p As DXF_Point
    Dim k As Long
    
    k = 0
    p = DP(0)
    ReDim Preserve Vertex(k) As DXF_Point
    Vertex(k).X = p.X
    Vertex(k).Y = p.Y
    
    For j = 3 To Count - 1
        StepT = (Knots(j + 1) - Knots(j)) / 25
        t = Knots(j)
        Do While StepT <> 0 And t < Knots(j + 1)
            p = NURBS_3(DP, Knots, j, t)
            k = k + 1
            ReDim Preserve Vertex(k) As DXF_Point
            Vertex(k).X = p.X
            Vertex(k).Y = p.Y
            
            t = t + StepT
        Loop
    Next
    p = DP(Count - 1)
    k = k + 1
    ReDim Preserve Vertex(k) As DXF_Point
    Vertex(k).X = p.X
    Vertex(k).Y = p.Y
End Sub
