Attribute VB_Name = "PathSPLines"
Option Explicit

Public SPLine_SegmentBetweenPoints As Long

Sub SplinePoints(CurSPline As Path_SPLine, sPoints() As PolygonPoint, ByVal SegmentBetweenPoints As Long)
    Dim n As Long, I As Long, Closed As Boolean
    Dim Vertex() As PolygonPoint, TmpPoints() As PolygonPoint
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim pt_max As Long, i_start As Long, i_stop As Long
    
    On Error Resume Next
    
    'n = UBound(CurSPline.Vertex)
    'If n < 2 Then
    '    ReDim sPoints(0) As PolygonPoint
    '    sPoints(0).X = CurSPline.Vertex(0).X
    '    sPoints(0).Y = CurSPline.Vertex(0).Y
    '    Exit Sub
    'End If
    n = CurSPline.vertex_count - 1
    If n < 2 Then
        ReDim sPoints(0) As PolygonPoint
        sPoints(0).X = PointList(CurSPline.vertex_id(0)).X
        sPoints(0).Y = PointList(CurSPline.vertex_id(0)).Y
        Exit Sub
    End If
    
    'CurSPline.segment_between_points = SPLine_SegmentBetweenPoints
    CurSPline.segment_between_points = SegmentBetweenPoints
    
    If CurSPline.vertex_id(0) <> CurSPline.vertex_id(CurSPline.vertex_count - 1) Or n < 4 Then 'Not closed or too few points
    
        ReDim Vertex(CurSPline.vertex_count - 1) As PolygonPoint
        For I = 0 To CurSPline.vertex_count - 1
            Vertex(I).X = PointList(CurSPline.vertex_id(I)).X
            Vertex(I).Y = PointList(CurSPline.vertex_id(I)).Y
        Next
        
        Closed = False
        
        'ReDim sPoints(15 + n ^ 2) As PolygonPoint
        'ReDim sPoints(60 + n ^ 2) As PolygonPoint
        pt_max = (CurSPline.segment_between_points + 1) * n
        ReDim sPoints(pt_max) As PolygonPoint
    
        'T_Spline CurSPline.Vertex(), 15, sPoints()
        C_Spline Vertex(), sPoints()
    
    Else 'Closed
        ReDim Vertex(CurSPline.vertex_count - 1 + 6) As PolygonPoint
        Vertex(0).X = PointList(CurSPline.vertex_id(CurSPline.vertex_count - 1 - 3)).X
        Vertex(0).Y = PointList(CurSPline.vertex_id(CurSPline.vertex_count - 1 - 3)).Y
        Vertex(1).X = PointList(CurSPline.vertex_id(CurSPline.vertex_count - 1 - 2)).X
        Vertex(1).Y = PointList(CurSPline.vertex_id(CurSPline.vertex_count - 1 - 2)).Y
        Vertex(2).X = PointList(CurSPline.vertex_id(CurSPline.vertex_count - 1 - 1)).X
        Vertex(2).Y = PointList(CurSPline.vertex_id(CurSPline.vertex_count - 1 - 1)).Y
        
        'Debug.Print ">>>"; CurSPline.vertex_count - 1
        'Debug.Print CurSPline.vertex_id(CurSPline.vertex_count - 1 - 3);
        'Debug.Print CurSPline.vertex_id(CurSPline.vertex_count - 1 - 2);
        'Debug.Print CurSPline.vertex_id(CurSPline.vertex_count - 1 - 1);

        For I = 0 To CurSPline.vertex_count - 1
            Vertex(I + 3).X = PointList(CurSPline.vertex_id(I)).X
            Vertex(I + 3).Y = PointList(CurSPline.vertex_id(I)).Y
            'Debug.Print CurSPline.vertex_id(I);
        Next
        Vertex(CurSPline.vertex_count - 1 + 4).X = PointList(CurSPline.vertex_id(1)).X
        Vertex(CurSPline.vertex_count - 1 + 4).Y = PointList(CurSPline.vertex_id(1)).Y
        Vertex(CurSPline.vertex_count - 1 + 5).X = PointList(CurSPline.vertex_id(2)).X
        Vertex(CurSPline.vertex_count - 1 + 5).Y = PointList(CurSPline.vertex_id(2)).Y
        Vertex(CurSPline.vertex_count - 1 + 6).X = PointList(CurSPline.vertex_id(3)).X
        Vertex(CurSPline.vertex_count - 1 + 6).Y = PointList(CurSPline.vertex_id(3)).Y
        
        'Debug.Print CurSPline.vertex_id(1);
        'Debug.Print CurSPline.vertex_id(2);
        'Debug.Print CurSPline.vertex_id(3)
    
        n = n + 6
    
        Closed = True
    End If
    
    If Closed = True Then
        pt_max = (CurSPline.segment_between_points + 1) * n
        ReDim TmpPoints(pt_max) As PolygonPoint
        C_Spline Vertex(), TmpPoints()
        
        x0 = PointList(CurSPline.vertex_id(0)).X
        y0 = PointList(CurSPline.vertex_id(0)).Y
        
        'find start
        For I = 1 To (pt_max) / 2
            x1 = TmpPoints(I - 1).X
            y1 = TmpPoints(I - 1).Y
            
            x2 = TmpPoints(I).X
            y2 = TmpPoints(I).Y
            
            If PointOnSegment(x0, y0, x1, y1, x2, y2, 0.2) Then
                i_start = I
                'Debug.Print "i_start="; i_start
                Exit For
            End If
        Next
    
        'find stop
        For I = i_start + 2 To pt_max - 1
            x1 = TmpPoints(I).X
            y1 = TmpPoints(I).Y
            
            x2 = TmpPoints(I + 1).X
            y2 = TmpPoints(I + 1).Y
            
            If PointOnSegment(x0, y0, x1, y1, x2, y2, 0.2) Then
                i_stop = I
                'Debug.Print "i_stop="; i_stop
                Exit For
            End If
        Next
    
        ReDim sPoints(i_stop - i_start + 2) As PolygonPoint
        sPoints(0).X = x0 '确使 sPoints(0) 与 PointList(CurSPLine.vertex_id(0)) 吻合
        sPoints(0).Y = y0
        For I = i_start To i_stop
            sPoints(I - i_start + 1) = TmpPoints(I)
        Next
        sPoints(i_stop - i_start + 2) = sPoints(0)
        
        'Debug.Print "Spline pn ="; i_stop - i_start + 2
    End If
    
    CurSPline.segment_point_count = UBound(sPoints)
    ReDim CurSPline.segment_point(UBound(sPoints)) As PolygonPoint
    For I = 0 To UBound(sPoints)
        CurSPline.segment_point(I) = sPoints(I)
    Next
    
End Sub


Public Sub Bezier_C(Pi() As PolygonPoint, Pc() As PolygonPoint)
    Dim I&, k&, NPI_1&, NPC_1&, NF&, u#, BF#

    NPI_1 = UBound(Pi)
    NPC_1 = UBound(Pc)
    'NF = Prodotto(NPI_1)

    For I = 0 To NPC_1
        u = CDbl(I) / CDbl(NPC_1)
        Pc(I).X = 0#
        Pc(I).Y = 0#
        'Pc(I).z = 0#
        
        For k = 0 To NPI_1
            'BF = NF * (u ^ K) * ((1 - u) ^ (NPI_1 - K)) / (Prodotto(K) * Prodotto(NPI_1 - K))
            BF = Prodotto(NPI_1, k + 1) * (u ^ k) * ((1 - u) ^ (NPI_1 - k)) / Prodotto(NPI_1 - k)
            Pc(I).X = Pc(I).X + Pi(k).X * BF
            Pc(I).Y = Pc(I).Y + Pi(k).Y * BF
            'Zu = Zu + Pi(K).z * BF
        Next k
    Next I
End Sub
Public Sub Bezier(Pi() As PolygonPoint, Pc() As PolygonPoint)
    Dim I&, k&, NPI_1&, NPC_1&
    Dim u#, u_1#, ue#, u_1e#, BF#

    NPI_1 = UBound(Pi) ' N. di punti da approssimare - 1.
    NPC_1 = UBound(Pc) ' N. di punti sulla curva - 1.

    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y
    'Pc(0).z = Pi(0).z

    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        ue = 1#
        u_1 = 1# - u
        u_1e = u_1 ^ NPI_1

        Pc(I).X = 0#
        Pc(I).Y = 0#
        'Pc(I).z = 0#
        For k = 0 To NPI_1
            BF = Prodotto(NPI_1, k + 1) * ue * u_1e / Prodotto(NPI_1 - k)
            Pc(I).X = Pc(I).X + Pi(k).X * BF
            Pc(I).Y = Pc(I).Y + Pi(k).Y * BF
            'Pc(I).z = Pc(I).z + Pi(K).z * BF

            ue = ue * u
            u_1e = u_1e / u_1
        Next k
    Next I

    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
    'Pc(NPC_1).z = Pi(NPI_1).z
End Sub

Public Sub Bezier_P(Pi() As PolygonPoint, Pc() As PolygonPoint)
    Dim k&, I&, KN&, NPI_1&, NPC_1&, NN&, NKN&
    Dim u#, uk#, unk#, Blend#

    NPI_1 = UBound(Pi)
    NPC_1 = UBound(Pc)

    For I = 0 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        uk = 1#
        unk = (1# - u) ^ NPI_1

        Pc(I).X = 0#
        Pc(I).Y = 0#
        'Pc(I).z = 0#

        For k = 0 To NPI_1
            NN = NPI_1
            KN = k
            NKN = NPI_1 - k
            Blend = uk * unk
            uk = uk * u
            unk = unk / (1# - u)
            Do While NN >= 1
                Blend = Blend * CDbl(NN)
                NN = NN - 1
                If KN > 1 Then
                    Blend = Blend / CDbl(KN)
                    KN = KN - 1
                End If
                If NKN > 1 Then
                    Blend = Blend / CDbl(NKN)
                    NKN = NKN - 1
                End If
            Loop

            Pc(I).X = Pc(I).X + Pi(k).X * Blend
            Pc(I).Y = Pc(I).Y + Pi(k).Y * Blend
            'Pc(I).z = Pc(I).z + Pi(K).z * Blend
        Next k
    Next I

    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
    'Pc(NPC_1).z = Pi(NPI_1).z
End Sub

Private Function Prodotto(ByVal n2&, Optional ByVal n1& = 2) As Double
    Dim f#, I&

    f = 1#
    For I = n1 To n2
        f = f * CDbl(I)
    Next I

    Prodotto = f
End Function


Public Sub B_Spline(Pi() As PolygonPoint, ByVal NK&, Pc() As PolygonPoint)
'       NK:                 Numero di nodi della curva
'                           approssimante:
'                           NK = 2    -> segmenti di retta.
'                           NK = 3    -> curve quadratiche.
'                           ..   .       ..................
'                           NK = NPI  -> splines di Bezier.

    Dim NPI_1&, NPC_1&, I&, j&, tmax#, u#, ut#, Eps#, bn#()

    NPI_1 = UBound(Pi)  ' N. di punti da approssimare - 1.
    NPC_1 = UBound(Pc)  ' N. di punti sulla curva - 1.
    Eps = 0.0000001
    tmax = NPI_1 - NK + 2

    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y

    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        ut = u * tmax
        If Abs(ut - CDbl(NPI_1 + NK - 2)) <= Eps Then
            Pc(I).X = Pi(NPI_1).X
            Pc(I).Y = Pi(NPI_1).Y
        Else
            Call B_Basis(NPI_1, ut, NK, bn())
            Pc(I).X = 0#
            Pc(I).Y = 0#
            For j = 0 To NPI_1
                Pc(I).X = Pc(I).X + bn(j) * Pi(j).X
                Pc(I).Y = Pc(I).Y + bn(j) * Pi(j).Y
            Next j
        End If
    Next I

    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
End Sub

Private Sub B_Basis(ByVal NPI_1&, ByVal ut#, ByVal k&, bn#())
'   Compute the basis function (also called weight)
'   for the B-Spline approximation curve:

    Dim NT&, I&, j&
    Dim b0#, b1#, bl0#, bl1#, bu0#, bu1#
    ReDim bn#(0 To NPI_1 + 1), bn0#(0 To NPI_1 + 1), t#(0 To NPI_1 + k + 1)

    NT = NPI_1 + k + 1
    For I = 0 To NT
        If (I < k) Then t(I) = 0#
        If ((I >= k) And (I <= NPI_1)) Then t(I) = CDbl(I - k + 1)
        If (I > NPI_1) Then t(I) = CDbl(NPI_1 - k + 2)
    Next I
    For I = 0 To NPI_1
        bn0(I) = 0#
        If ((ut >= t(I)) And (ut < t(I + 1))) Then bn0(I) = 1#
        If ((t(I) = 0#) And (t(I + 1) = 0#)) Then bn0(I) = 0#
    Next I

    For j = 2 To k
        For I = 0 To NPI_1
            bu0 = (ut - t(I)) * bn0(I)
            bl0 = t(I + j - 1) - t(I)
            If (bl0 = 0#) Then
                b0 = 0#
            Else
                b0 = bu0 / bl0
            End If
            bu1 = (t(I + j) - ut) * bn0(I + 1)
            bl1 = t(I + j) - t(I + 1)
            If (bl1 = 0#) Then
                b1 = 0#
            Else
                b1 = bu1 / bl1
            End If
            bn(I) = b0 + b1
        Next I
        For I = 0 To NPI_1
            bn0(I) = bn(I)
        Next I
    Next j
End Sub

Public Sub C_Spline(Pi() As PolygonPoint, Pc() As PolygonPoint)
    Dim NPI_1&, NPC_1&, I&, j&
    Dim u#, ui#, uui#
    Dim cof() As PolygonPoint

    NPI_1 = UBound(Pi)      ' N. di punti da interpolare - 1.
    NPC_1 = UBound(Pc)      ' N. di punti sulla curva - 1.

    Call Find_CCof(Pi(), NPI_1 + 1, cof())

    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y

    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        j = Int(u * CDbl(NPI_1)) + 1
        If (j > (NPI_1)) Then j = NPI_1

        ui = CDbl(j - 1) / CDbl(NPI_1)
        uui = u - ui

        Pc(I).X = cof(4, j).X * uui ^ 3 + cof(3, j).X * uui ^ 2 + cof(2, j).X * uui + cof(1, j).X
        Pc(I).Y = cof(4, j).Y * uui ^ 3 + cof(3, j).Y * uui ^ 2 + cof(2, j).Y * uui + cof(1, j).Y
    Next I

    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
End Sub

Public Sub T_Spline(Pi() As PolygonPoint, ByVal VZ&, Pc() As PolygonPoint)
'       VZ:                 Valore di tensione della curva
'                           (1 <= VZ <= 100): valori grandi
'                           di VZ appiattiscono la curva.
'
    Dim NPI_1&, NPC_1&, I&, j&
    Dim h#, z#, z2i#, szh#, u#, u0#, u1#, du1#, du0#
    Dim s() As PolygonPoint

    NPI_1 = UBound(Pi)      ' N. di punti da interpolare - 1.
    NPC_1 = UBound(Pc)      ' N. di punti sulla curva - 1.
    z = CDbl(VZ)

    Call Find_TCof(Pi(), NPI_1 + 1, s(), z)

    ' La curva inizia sempre da Pi(0) -> u = 0:
    Pc(0).X = Pi(0).X
    Pc(0).Y = Pi(0).Y

    h = 1# / CDbl(NPI_1)
    szh = Sinh(z * h)
    z2i = 1# / z / z
    For I = 1 To NPC_1 - 1
        u = CDbl(I) / CDbl(NPC_1)
        j = Int(u * CDbl(NPI_1)) + 1
        If (j > (NPI_1)) Then j = NPI_1

        u0 = CDbl(j - 1) / CDbl(NPI_1)
        u1 = CDbl(j) / CDbl(NPI_1)
        du1 = u1 - u
        du0 = u - u0

        Pc(I).X = s(j).X * z2i * Sinh(z * du1) / szh + (Pi(j - 1).X - s(j).X * z2i) * du1 / h
        Pc(I).X = Pc(I).X + s(j + 1).X * z2i * Sinh(z * du0) / szh + (Pi(j).X - s(j + 1).X * z2i) * du0 / h
    
        Pc(I).Y = s(j).Y * z2i * Sinh(z * du1) / szh + (Pi(j - 1).Y - s(j).Y * z2i) * du1 / h
        Pc(I).Y = Pc(I).Y + s(j + 1).Y * z2i * Sinh(z * du0) / szh + (Pi(j).Y - s(j + 1).Y * z2i) * du0 / h
    Next I

    ' La curva finisce sempre su Pi(NPI_1) -> u = 1:
    Pc(NPC_1).X = Pi(NPI_1).X
    Pc(NPC_1).Y = Pi(NPI_1).Y
End Sub

Private Sub Find_TCof(Pi() As PolygonPoint, ByVal NPI&, s() As PolygonPoint, ByVal z#)
'   Find the coefficients of the T-Spline
'   using constant interval:

    Dim I&, h#, a0#, b0#, zh#, Z2#

    ReDim s(1 To NPI) As PolygonPoint, f(1 To NPI) As PolygonPoint
    ReDim a(1 To NPI) As Double, b(1 To NPI) As Double, c(1 To NPI) As Double

    h = 1# / CDbl(NPI - 1)
    zh = z * h
    a0 = 1# / h - z / Sinh(zh)
    b0 = z * 2# * Cosh(zh) / Sinh(zh) - 2# / h
    For I = 1 To NPI - 2
        a(I) = a0
        b(I) = b0
        c(I) = a0
    Next I

    Z2 = z * z / h
    For I = 1 To NPI - 2
        f(I).X = (Pi(I + 1).X - 2# * Pi(I).X + Pi(I - 1).X) * Z2
        f(I).Y = (Pi(I + 1).Y - 2# * Pi(I).Y + Pi(I - 1).Y) * Z2
    Next I

    Call TRIDAG(a(), b(), c(), f(), s(), NPI - 2)
    For I = 1 To NPI - 2
        s(NPI - I).X = s(NPI - 1 - I).X
        s(NPI - I).Y = s(NPI - 1 - I).Y
    Next I

    s(1).X = 0#
    s(NPI).X = 0#
    s(1).Y = 0#
    s(NPI).Y = 0#
End Sub

Private Sub Find_CCof(Pi() As PolygonPoint, ByVal NPI&, cof() As PolygonPoint)
'   Find the coefficients of the cubic spline
'   using constant interval parameterization:

    Dim I&, h#

    ReDim s(1 To NPI) As PolygonPoint, f(1 To NPI) As PolygonPoint, cof(1 To 4, 1 To NPI) As PolygonPoint
    ReDim a(1 To NPI) As Double, b(1 To NPI) As Double, c(1 To NPI) As Double

    h = 1# / CDbl(NPI - 1)
    For I = 1 To NPI - 2
        a(I) = 1#
        b(I) = 4#
        c(I) = 1#
    Next I

    For I = 1 To NPI - 2
        f(I).X = 6# * (Pi(I + 1).X - 2# * Pi(I).X + Pi(I - 1).X) / h / h
        f(I).Y = 6# * (Pi(I + 1).Y - 2# * Pi(I).Y + Pi(I - 1).Y) / h / h
    Next I

    Call TRIDAG(a(), b(), c(), f(), s(), NPI - 2)
    For I = 1 To NPI - 2
        s(NPI - I).X = s(NPI - 1 - I).X
        s(NPI - I).Y = s(NPI - 1 - I).Y
    Next I

    s(1).X = 0#
    s(NPI).X = 0#
    s(1).Y = 0#
    s(NPI).Y = 0#
    For I = 1 To NPI - 1
        cof(4, I).X = (s(I + 1).X - s(I).X) / 6# / h
        cof(4, I).Y = (s(I + 1).Y - s(I).Y) / 6# / h
        cof(3, I).X = s(I).X / 2#
        cof(3, I).Y = s(I).Y / 2#
        cof(2, I).X = (Pi(I).X - Pi(I - 1).X) / h - (2# * s(I).X + s(I + 1).X) * h / 6#
        cof(2, I).Y = (Pi(I).Y - Pi(I - 1).Y) / h - (2# * s(I).Y + s(I + 1).Y) * h / 6#
        cof(1, I).X = Pi(I - 1).X
        cof(1, I).Y = Pi(I - 1).Y
    Next I
End Sub

Private Sub TRIDAG(a#(), b#(), c#(), f() As PolygonPoint, s() As PolygonPoint, ByVal NPI_2&)
'   Solves the tridiagonal linear system of equations:

    Dim j&, bet#
    ReDim gam#(1 To NPI_2)

    If b(1) = 0 Then Exit Sub

    bet = b(1)
    s(1).X = f(1).X / bet
    s(1).Y = f(1).Y / bet
    For j = 2 To NPI_2
        gam(j) = c(j - 1) / bet
        bet = b(j) - a(j) * gam(j)
        If (bet = 0) Then Exit Sub
        s(j).X = (f(j).X - a(j) * s(j - 1).X) / bet
        s(j).Y = (f(j).Y - a(j) * s(j - 1).Y) / bet
    Next j

    For j = NPI_2 - 1 To 1 Step -1
        s(j).X = s(j).X - gam(j + 1) * s(j + 1).X
        s(j).Y = s(j).Y - gam(j + 1) * s(j + 1).Y
    Next j
End Sub

Private Function Cosh(ByVal z As Double) As Double
    Cosh = (Exp(z) + Exp(-z)) / 2#
End Function

Private Function Sinh(ByVal z As Double) As Double
    Sinh = (Exp(z) - Exp(-z)) / 2#
End Function

