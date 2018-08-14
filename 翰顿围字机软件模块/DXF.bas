Attribute VB_Name = "DXF"
Option Explicit

Type DataSet
    code As Integer
    Value As Variant
End Type

Type Entity
    Type As String
    Data() As DataSet
End Type

Type Block
    Name As String
    Entities() As Entity
End Type

Type DXFData
    Blocks() As Block
    Entities() As Entity
End Type

Dim Section() As String
Dim DXF As DXFData

Function FindDataX(sArray() As String, Start As Long)
    Dim I As Long
    For I = Start To UBound(sArray)
        If sArray(I) = "10" And sArray(I + 2) = "20" Then
            FindDataX = I
            Exit Function
        End If
    Next I
    FindDataX = -1
End Function

Sub PrepareEntity(CurEntity As Entity)
    Dim Point0 As Path_Point, Point1 As Path_Point, Pointm As Path_Point
    Dim cx As Double, cy As Double, cz As Double, r As Double, r2 As Double, Angle0 As Double, Angle1 As Double
    Dim I As Long, K As Integer, n As Long, id0 As Long, ux As Double, uy As Double, uz As Double
    
    Select Case CurEntity.Type
        Case "LINE"
            Point0.X = GetGroupValue(CurEntity.Data(), 10)
            Point0.Y = GetGroupValue(CurEntity.Data(), 20)
            Point0.Z = GetGroupValue(CurEntity.Data(), 30)
            
            Point1.X = GetGroupValue(CurEntity.Data(), 11)
            Point1.Y = GetGroupValue(CurEntity.Data(), 21)
            Point1.Z = GetGroupValue(CurEntity.Data(), 31)
            
            '-------------------------------------------------------------
            CatchOrAddPoint Point0
            CatchOrAddPoint Point1
            
            AddSegment Point0.ID, Point1.ID
            
        Case "ARC"
            cx = GetGroupValue(CurEntity.Data(), 10)
            cy = GetGroupValue(CurEntity.Data(), 20)
            cz = GetGroupValue(CurEntity.Data(), 30)
            
            r = GetGroupValue(CurEntity.Data(), 40)
            Angle0 = GetGroupValue(CurEntity.Data(), 50) * PI_180
            Angle1 = GetGroupValue(CurEntity.Data(), 51) * PI_180
            
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
            Point0.Z = cz
            
            Point1.X = r * Cos(Angle1) + cx
            Point1.Y = r * Sin(Angle1) + cy
            Point1.Z = cz
            
            CatchOrAddPoint Point0
            CatchOrAddPoint Point1
            
            AddArc cx, cy, cz, r, r, Angle0, Angle1, Point0.ID, Point1.ID, 0, 0, ArcType.CircleCR
        
        Case "CIRCLE"
            cx = GetGroupValue(CurEntity.Data(), 10)
            cy = GetGroupValue(CurEntity.Data(), 20)
            cz = GetGroupValue(CurEntity.Data(), 30)
            
            r = GetGroupValue(CurEntity.Data(), 40)
            
            Angle0 = GetGroupValue(CurEntity.Data(), 50) * PI_180
            Angle1 = GetGroupValue(CurEntity.Data(), 51) * PI_180
            
            '-------------------------------------------------------------
            Angle0 = 0
            Angle1 = PI2
            
            Point0.X = r * Cos(Angle0) + cx
            Point0.Y = r * Sin(Angle0) + cy
            Point0.Z = cz
            
            Point1.X = r * Cos(Angle1) + cx
            Point1.Y = r * Sin(Angle1) + cy
            Point1.Z = cz
            
            CatchOrAddPoint Point0
            CatchOrAddPoint Point1
            
            AddArc cx, cy, cz, r, r2, Angle0, Angle1, Point0.ID, Point1.ID, 0, 0, ArcType.CircleCR
        
        Case "ELLIPSE"
            cx = GetGroupValue(CurEntity.Data(), 10)
            cy = GetGroupValue(CurEntity.Data(), 20)
            cz = GetGroupValue(CurEntity.Data(), 30)
            
            ux = GetGroupValue(CurEntity.Data(), 11)
            uy = GetGroupValue(CurEntity.Data(), 21)
            uz = GetGroupValue(CurEntity.Data(), 31)
            
            r = Sqr(ux * ux + uy * uy)
            r2 = GetGroupValue(CurEntity.Data(), 40) * r
            
            Angle0 = GetGroupValue(CurEntity.Data(), 50) * PI_180
            Angle1 = GetGroupValue(CurEntity.Data(), 51) * PI_180
            
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
            Point0.Z = cz
            
            Point1.X = r * Cos(Angle1) + cx
            Point1.Y = r2 * Sin(Angle1) + cy
            Point1.Z = cz
            
            CatchOrAddPoint Point0
            CatchOrAddPoint Point1
            
            AddArc cx, cy, cz, r, r, Angle0, Angle1, Point0.ID, Point1.ID, 0, 0, ArcType.CircleCR
        
        Case "POINT"
            Point0.X = GetGroupValue(CurEntity.Data(), 10)
            Point0.Y = GetGroupValue(CurEntity.Data(), 20)
            Point0.Z = GetGroupValue(CurEntity.Data(), 30)
                        
            '-------------------------------------------------------------
            CatchOrAddPoint Point0
        
        Case "LWPOLYLINE"
            n = GetGroupValue(CurEntity.Data(), 90)
            K = GetGroupValue(CurEntity.Data(), 70)

            For I = 1 To n
                If I = 1 Then
                    Point0.X = GetGroupValue(CurEntity.Data(), 10, True)
                    Point0.Y = GetGroupValue(CurEntity.Data(), 20, True)
                    Point0.Z = 0
                    CatchOrAddPoint Point0
                    id0 = Point0.ID
                Else
                    Point1.X = GetGroupValue(CurEntity.Data(), 10, True)
                    Point1.Y = GetGroupValue(CurEntity.Data(), 20, True)
                    Point1.Z = 0
                    CatchOrAddPoint Point1
                    
                    AddSegment id0, Point1.ID
                    id0 = Point1.ID
                End If
            Next
            
            If K = 1 And n > 2 Then
                AddSegment Point1.ID, Point0.ID
            End If
            
        Case "SPLINE"
            AddSPLine
            n = GetGroupValue(CurEntity.Data(), 74) 'fit point
            'K = GetGroupValue(CurEntity.Data(), 70)

            For I = 0 To n - 1
                ux = GetGroupValue(CurEntity.Data(), 11, True)
                uy = GetGroupValue(CurEntity.Data(), 21, True)
                
                If I = 0 Then
                    Point0.X = ux
                    Point0.Y = uy
                    Point0.Z = 0
                    Point0.Type = PointType.SPLinePoint
                    CatchOrAddPoint Point0
                    id0 = Point0.ID
                Else
                    Point1.X = ux
                    Point1.Y = uy
                    Point1.Z = 0
                    Point0.Type = PointType.SPLinePoint
                    CatchOrAddPoint Point1
                    id0 = Point1.ID
                End If
                    
                ReDim Preserve SPLineList(SPLineCount).Vertex(I)
                SPLineList(SPLineCount).Vertex(I).X = ux
                SPLineList(SPLineCount).Vertex(I).Y = uy
            Next
            
            SPLineList(SPLineCount).point0_id = Point0.ID
            SPLineList(SPLineCount).point1_id = Point1.ID
            
        Case "VERTEX"
        Case "TEXT"
        Case "INSERT"
        Case "DIMENSION"
    End Select
End Sub

Sub FindCommand(FileNum As Integer, Command As String)
    Dim X As String
    Do While UCase(Trim(X)) <> UCase(Command)
        Line Input #FileNum, X
    Loop
End Sub

Function GetBlock(DXF As DXFData, Name As String) As Integer
    Dim I As Integer
    For I = 0 To UBound(DXF.Blocks)
        If DXF.Blocks(I).Name = Name Then
            GetBlock = I
            Exit Function
        End If
    Next I
End Function

Function GetSection(FileNum As Integer, Start As String, Finish As String, EndString As String, sArray() As String) As Boolean
    Dim Temp As String
    Dim I As Long
    
    ReDim sArray(0) As String
    
    Do While Temp <> Start
        Line Input #FileNum, Temp
        Temp = UCase(Trim(Temp))
        If Temp = EndString Then
            GetSection = False
            Exit Function
        End If
    Loop
    
    Do While Temp <> Finish
        Line Input #FileNum, Temp
        Temp = UCase(Trim(Temp))
        If Temp <> Finish Then
            ReDim Preserve sArray(I) As String
            sArray(I) = Temp
            I = I + 1
        End If
    Loop
    GetSection = True
End Function

Sub ImportDXFFile(FileDXF As String)
    Dim FF As Integer
    Dim DXFLine As String
    Dim bCount As Integer
    Dim ENDSEC As Boolean
    
    ReDim DXF.Blocks(0) As Block
    ReDim DXF.Entities(0) As Entity
    
    FF = FreeFile
    Open FileDXF For Input As #FF
    FindCommand FF, "BLOCKS"
    Do While Not ENDSEC
        If GetSection(FF, "BLOCK", "ENDBLK", "ENDSEC", Section()) Then
            ReDim Preserve DXF.Blocks(bCount) As Block
            ReDim Preserve DXF.Blocks(bCount).Entities(0) As Entity
            If ParseBlock(Section(), DXF.Blocks(bCount)) Then
                bCount = bCount + 1
            End If
        Else
            ENDSEC = True
        End If
    Loop
    GetSection FF, "ENTITIES", "ENDSEC", "ENDSEC", Section()
    Close #FF
    
    ParseEntities Section(), 0, DXF.Entities()
    
    Erase DXF.Blocks
    Erase DXF.Entities
    Erase Section
End Sub

Function IsEntityCommand(ByVal InText As String)
    Select Case UCase(InText)
        Case "LINE", "POINT", "VERTEX", "POLYLINE", "CIRCLE", "ARC", "ELLIPSE", "LWPOLYLINE", "SPLINE", "TEXT", "INSERT", "DIMENSION"
            IsEntityCommand = True
        Case Else
            IsEntityCommand = False
    End Select
End Function

Function GetGroupValue(Data() As DataSet, code As Integer, Optional ClearCode As Boolean = False) As Variant
    Dim I As Integer
    
    For I = 0 To UBound(Data)
        If Data(I).code = code Then
            GetGroupValue = Data(I).Value
            If ClearCode = True Then
                Data(I).code = 0
            End If
            Exit Function
        End If
    Next I
    GetGroupValue = 0
End Function

Function ParseBlock(sArray() As String, CurBlock As Block) As Boolean
    'On Local Error GoTo exitMe:
    Dim I As Long
    
    'first we have to look for a section "6" to determine if this BLOCK section is "important"
    I = SearchSection(sArray(), I, "6")
    If I = -1 Then
        ParseBlock = False
        Exit Function
    End If
    I = SearchSection(sArray(), I, "2") + 1
    CurBlock.Name = sArray(I)
    
    ParseEntities sArray, I, CurBlock.Entities
    
    ParseBlock = True
    Exit Function
    
ExitMe:
    MsgBox "ERROR  " & Err.Description
End Function

Function ParseEntities(sArray() As String, ByVal Start As Long, CurEntities() As Entity) As Boolean
    Dim J As Long
    Dim K As Long
    Dim P As Long
    
    K = 0
    For J = Start To UBound(sArray)
        If IsEntityCommand(sArray(J)) Then 'we found an ENTITY COMMAND
            ReDim Preserve CurEntities(K) As Entity
            CurEntities(K).Type = sArray(J)
            
            Select Case CurEntities(K).Type
                Case "INSERT", "DIMENSION"
                    'KEY "2" on an INSERT provides the BLOCK name to be inserted to the PV
                    J = SearchSection(sArray(), J, "2")
                Case Else
                    J = J + 1 'FindDataX(sArray(), j)
            End Select
            
            P = 0
            Do While sArray(J) <> "0"
                ReDim Preserve CurEntities(K).Data(P)
                CurEntities(K).Data(P).code = sArray(J)
                CurEntities(K).Data(P).Value = sArray(J + 1)
                P = P + 1
                J = J + 2
            Loop
            PrepareEntity CurEntities(K)
            K = K + 1
        End If
    Next J
    ParseEntities = True
End Function

Function SearchSection(sArray() As String, Start As Long, Value As String) As Long
    Dim I As Long
    
    For I = Start To UBound(sArray)
        If sArray(I) = Value Then
            SearchSection = I
            Exit Function
        End If
    Next I
    SearchSection = -1
End Function


