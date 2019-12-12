Attribute VB_Name = "Cmath"
' Curve generation using quadratics and solid geussing
Public Function quad(xindex1 As Integer, yindex1 As Integer, xindex2 As Integer, yindex2 As Integer, distance As Single) As Single
' calculate the quadratic curve elevation of the point N distance from xindex1,yindex1
' going towards xindex2,yindex2. Account for exceptionals
'   ? One spline segment in the (u,v)-plane is given by the equations  ?
'   ?   fu(t) = Au * t * t * t + Bu * t * t + Cu * t + Du  and  ?
'   ?   fv(t) = Av * t * t * t + Bv * t * t + Cv * t + Dv  ?
' aha! three points on a parabola

End Function
Public Function readDxf(fyl As String) As Boolean

'1000   Rem
'1010   Rem Extract lines from DXF file
'1020   Rem
'1030   G1% = 0
'1040   LINE INPUT "DXF file name: "; A$
'1050   OPEN "i", 1, A$ + ".dxf"
'1060   Rem
'1070   Rem Ignore until section start encountered
'1080   Rem
'1090   GoSub 2000
'1100   If G% <> 0 Then GoTo 1090
'1110   If S$ <> "SECTION" Then GoTo 1090
'1120   GoSub 2000
'1130   Rem
'1140   Rem Skip unless ENTITIES section
'1150   Rem
'1160   If S$ <> "ENTITIES" Then GoTo 1090
'1170   Rem
'1180   Rem Scan until end of section, processing LINEs
'1190   Rem
'1200   GoSub 2000
'1210   If G% = 0 And S$ = "ENDSEC" Then GoTo 2200
'1220   If G% = 0 And S$ = "LINE" Then GoSub 1400: GoTo 1210
'1230   GoTo 1200
'1400   Rem
'1410   Rem Accumulate LINE entity groups
'1420   Rem
'1430   GoSub 2000
'1440   If G% = 10 Then X1 = X: Y1 = Y: Z1 = Z
'1450   If G% = 11 Then X2 = X: Y2 = Y: Z2 = Z
'1460   If G% = 0 Then Print "Line from ("; X1; ","; Y1; ","; Z1; ") to "
'       (";X2;",";Y2;",";Z2;")":RETURN
'1470   GoTo 1430
'2000   Rem
'2010   Rem Read group code and following value
'2020   Rem For X coordinates, read Y and possibly Z also
'2030   Rem
'2040   If G1% < 0 Then G% = -G1%: G1% = 0 Else Input #1, G%
'2050   If G% < 10 Or G% = 999 Then Line Input #1, S$: Return
'2060   If G% >= 38 And G% <= 49 Then Input #1, V: Return
'2080   If G% >= 50 And G% <= 59 Then Input #1, A: Return
'2090   If G% >= 60 And G% <= 69 Then Input #1, P%: Return
'2100   If G% >= 70 And G% <= 79 Then Input #1, f%: Return
'2110   If G% >= 210 And G% <= 219 Then GoTo 2130
'2115   If G% >= 1000 Then Line Input #1, T$: Return
'2120   If G% >= 20 Then Print "Invalid group code"; G%: Stop
'2130   Input #1, X
'2140   Input #1, G1%
'2150   If G1% <> (G% + 10) Then Print "Invalid Y coord code"; G1%:
'       Stop
'2160   Input #1, Y
'2170   Input #1, G1%
'2180   If G1% <> (G% + 20) Then G1% = -G1% Else Input #1, Z
'2190   Return
'2200   Close 1
'
End Function
Public Function makeDxf(fyl As String) As Boolean
'1000   Rem
'1010   Rem Polygon generator
'1020   Rem
'1030   LINE INPUT "Drawing (DXF) file name: "; A$
'1040   OPEN "o", 1, A$ + ".dxf"
'1050   Print #1, 0
'1060   Print #1, "SECTION"
'1070   Print #1, 2
'1080   Print #1, "ENTITIES"
'1090   PI = Atn(1) * 4
'1100   INPUT "Number of sides for polygon: "; S%
'1110   INPUT "Starting point (X,Y): "; X, Y
'1120   INPUT "Polygon side: "; D
'1130   A1 = (2 * PI) / S%
'1140   A = PI / 2
'1150   For I% = 1 To S%
'1160   Print #1, 0
'1170   Print #1, "LINE"
'1180   Print #1, 8
'1190   Print #1, "0"
'1200   Print #1, 10
'1210   Print #1, X
'1220   Print #1, 20
'1230   Print #1, Y
'1240   Print #1, 30
'1250   Print #1, 0#
'1260   NX = D * Cos(A) + X
'1270   NY = D * Sin(A) + Y
'1280   Print #1, 11
'1290   Print #1, NX
'1300   Print #1, 21
'1310   Print #1, NY
'1320   Print #1, 31
'1330   Print #1, 0#
'1340   X = NX
'1350   Y = NY
'1360   A = A + A1
'1370   Next I%
'1380   Print #1, 0
'1390   Print #1, "ENDSEC"
'1400   Print #1, 0
'1410   Print #1, "EOF"
'1420   Close 1

End Function
