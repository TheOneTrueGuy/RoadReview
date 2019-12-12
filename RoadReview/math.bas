Attribute VB_Name = "math"
Option Explicit
' the basic grid is the lat/lon sphere
' all sub segments are simply displacements and rotations
' around a point on the lat/lon grid



' Public Type CUSTOMVERTEX
'    x As double         'x in screen space
'    y As double         'y in screen space
'    z  As double        'normalized z
'    rhw As double
'    color As Long      'vertex color
'End Type

Public Const PI As Double = 3.14159265358979
Public Type Poynt
x As Double
y As Double
z As Double
End Type
Public Type latitude
degrees As Double
minutes As Double
seconds As Double
End Type
Public Type longitude
degrees As Double
minutes As Double
seconds As Double
End Type


Public Type segment
points(6) As Poynt
End Type


Public Type station
lat As Double
lng As Double
slat As String
slng As String
elev As Double
pnt As Poynt
End Type

Public Type Curve
A As Double
D As Double
T As Double
L As Double
R As Double
LC As Double
C As Double
pc As String
PI As String
PT As String
PCn As Poynt
PIn As Double
PTn As Double
superElev As String
WL As Double
WR As Double
origin As Poynt
leftright As Boolean
ro As Long
End Type

Public Type sections
Crv As Curve
segm(200) As segment
stash(200) As station
prvx As Long ' both these point to grid values
prvy As Long
nexx As Long
nexy As Long
offsetlat As latitude
offsetlon As longitude
End Type

Public LLgrid() As sections
Public curve1 As Curve
Public curves() As Curve
Public curvecount As Long

Dim stations() As station, proposed_elev() As station
Public loaded_from_disk As Boolean, set_default As Boolean
Public landname As String, gnu As Boolean, opn_gnu As Boolean
Public numx As Integer, numy As Integer, standard As Double
Public benchx As Integer, benchy As Integer
Public wstation_distance As Double, lstation_distance As Double
Public wslope As Double, lslope As Double, factor As Integer
Public cam_height As Integer, rotate_angle As Integer, escale As Integer
Public UpDownZoom As Integer, HScroll1 As Integer, VScroll1 As Integer
Public xcen2 As Double, ycen2 As Double, screen_dist As Double
Public data_changed As Boolean
Public transfer As station

Public shrink As Double
Public offset As Double
Dim baseline As Double, baseline2 As Double
Public cut_vol As Double, fill_vol As Double
Public pcut_vol As Double, pfill_vol As Double
Public dcut_vol As Double, dfill_vol As Double
Public popeye As Boolean, numpoints As Integer
Dim benchplane() As Double
Dim set_as_default As Boolean
Public elvis As station

Public Function parseDegrees(coords As String) As Double
' parse up a degree listing like this
' ##D##M##S
' S is optional
Dim degs As String, mins As String, secs As String
Dim place, rema$
place = InStr(1, coords, "d")
degs = Left(coords, place - 1)
rema$ = Right(coords, Len(coords) - place)
place = InStr(1, rema, "m")
mins = Left(rema, place - 1)
rema$ = Right(rema, Len(rema) - place)
secs = rema$
' angle_degrees = degrees + (minutes / 60#) + (seconds / 3600#)

Dim angdeg As Double, angradi As Double
angdeg = CSng(degs) + (CSng(Left(mins, 2)) / 60#) + (CSng(secs) / 3600#)
angradi = (PI / 180) * angdeg
parseDegrees = angradi
End Function

Public Sub reInit(newx As Integer, newy As Integer)
ReDim stations(newx, newy)
ReDim proposed_elev(newx, newy)

End Sub

'Public Function cut_volume() As double
'Dim Station As Integer, total_volume As double, area1 As double, area2 As double
'Dim volume As double, ave_area As double
'
'baseline = standard: baseline2 = 0 ' put these in decs
'For Station = 1 To numx - 1
'baseline = baseline + wslope
'area1 = calc_cut_area(Station)
'area2 = calc_cut_area(Station + 1)
'ave_area = (area1 + area2) / 2
'volume = ave_area * wstation_distance
'total_volume = total_volume + volume
'Next Station
'cut_volume = total_volume / 27 ' returns cubic yards
''Debug.Print total_volume / 27
'End Function
'Public Function calc_cut_area(Station As Integer) As double
'Dim total As double, curhyt As double, nexhyt As double, hyt As double
'Dim substation As double, difhyt As double
'Dim hytseg As double
'Dim distseg As double, inc As Integer
'baseline2 = baseline
'total = 0: substation = 0: distseg = lstation_distance * 0.1
'For substation = 1 To numy - 1
'baseline2 = baseline2 + lslope
'curhyt = stations(Station, substation).elev(0) - baseline2 '
'
'nexhyt = stations(Station, substation + 1).elev(0) - baseline2 '
'difhyt = Abs(curhyt - nexhyt)
''Debug.Print difhyt
'hytseg = difhyt / 10
'hyt = curhyt
'If curhyt <= 0 And nexhyt <= 0 Then GoTo skipper
'For inc = 1 To 10
'If curhyt > nexhyt Then hyt = hyt - hytseg Else hyt = hyt + hytseg
'If hyt <= 0 Then GoTo Skip
'total = total + (hyt * distseg)
'Skip:
'Next inc
'skipper:
'Next substation
'' mult totalhyt by increment segment to get area
'calc_cut_area = total ' assumes increments of 1
'End Function
'Public Function fill_volume() As double
'Dim Station As Integer, total_volume As double, area1 As double, area2 As double
'Dim volume As double, ave_area As double
'
'baseline = standard: baseline2 = 0 ' put these in decs
'For Station = 1 To numx - 1
'area1 = calc_fill_area(Station)
'area2 = calc_fill_area(Station + 1)
'ave_area = (area1 + area2) / 2
'volume = ave_area * wstation_distance
'total_volume = total_volume + volume
'Next Station
'fill_volume = Abs(total_volume / 27) ' cubic yards
'End Function
'Public Function calc_fill_area(Station As Integer) As double
'Dim total As double, curhyt As double, nexhyt As double, hyt As double
'Dim distseg As double
'Dim hytseg As double, inc As double
'Dim substation As double, difhyt As double
'baseline2 = baseline
'total = 0: substation = 0: distseg = lstation_distance * 0.1
'For substation = 1 To numy - 1
'baseline2 = baseline2 + lslope
'curhyt = stations(Station, substation).elev(0) - baseline2
'nexhyt = stations(Station, substation + 1).elev(0) - baseline2
'  difhyt = Abs(curhyt - nexhyt)
''Debug.Print difhyt
'  hytseg = difhyt / 10
'  hyt = curhyt
'If curhyt >= 0 And nexhyt >= 0 Then GoTo skipper
' For inc = 1 To 10
'   If curhyt > nexhyt Then hyt = hyt - hytseg Else hyt = hyt + hytseg
'   If hyt >= 0 Then GoTo Skip
'   total = total + (hyt * distseg)
'Skip:
' Next inc
'skipper:
'Next substation
'' mult totalhyt by increment segment to get area
'calc_fill_area = total ' assumes increments of 1
'
'End Function
'' calc differential of start and final
'Public Function dif_cut_volume() As double
'Dim Station As Integer, total_volume As double, area1 As double, area2 As double
'Dim volume As double, ave_area As double
'
'baseline = standard: baseline2 = 0 ' put these in decs
'For Station = 1 To numx - 1
'baseline = baseline + wslope
'area1 = dcalc_cut_area(Station)
'area2 = dcalc_cut_area(Station + 1)
'ave_area = (area1 + area2) / 2
'volume = ave_area * wstation_distance
'total_volume = total_volume + volume
'Next Station
'dif_cut_volume = total_volume / 27 ' returns cubic yards
'End Function
'Public Function dcalc_cut_area(Station As Integer) As double
'Dim total As double, curhyt As double, nexhyt As double, hyt As double
'Dim distseg As double
'Dim substation As double, difhyt As double
'Dim hytseg As double, inc As Integer
'
'baseline2 = baseline
'total = 0: substation = 0: distseg = lstation_distance * 0.1
'For substation = 1 To numy - 1
'baseline2 = baseline2 + lslope
'curhyt = stations(Station, substation).elev(0) - proposed_elev(Station, substation).elev(0)
'nexhyt = stations(Station, substation + 1).elev(0) - proposed_elev(Station, substation).elev(0)
'difhyt = Abs(curhyt - nexhyt)
''Debug.Print difhyt
'hytseg = difhyt / 10
'hyt = curhyt
'If curhyt <= 0 And nexhyt <= 0 Then GoTo skipper
'For inc = 1 To 10
'If curhyt > nexhyt Then hyt = hyt - hytseg Else hyt = hyt + hytseg
'If hyt <= 0 Then GoTo Skip
'total = total + (hyt * distseg)
'Skip:
'Next inc
'skipper:
'Next substation
'' mult totalhyt by increment segment to get area
'dcalc_cut_area = total ' assumes increments of 1
'End Function
'Public Function dif_fill_volume() As double
'Dim Station As Integer, total_volume As double, area1 As double, area2 As double
'Dim volume As double, ave_area As double
'baseline = standard: baseline2 = 0 ' put these in decs
'For Station = 1 To numx - 1
'area1 = dcalc_fill_area(Station)
'area2 = dcalc_fill_area(Station + 1)
'ave_area = (area1 + area2) / 2
'volume = ave_area * wstation_distance
'total_volume = total_volume + volume
'Next Station
'dif_fill_volume = Abs(total_volume / 27) ' cubic yards
'End Function
'Public Function dcalc_fill_area(Station As Integer) As double
'Dim total As double, curhyt As double, nexhyt As double, hyt As double
'Dim distseg As double
'Dim substation As double, difhyt As double
'Dim hytseg As double, inc As Integer
'
'baseline2 = baseline
'total = 0: substation = 0: distseg = lstation_distance * 0.1
'For substation = 1 To numy - 1
'baseline2 = baseline2 + lslope
'curhyt = stations(Station, substation).elev(0) - proposed_elev(Station, substation).elev(0)
'nexhyt = stations(Station, substation + 1).elev(0) - proposed_elev(Station, substation).elev(0) ' to proposed_elev
'  difhyt = Abs(curhyt - nexhyt)
''Debug.Print difhyt
'  hytseg = difhyt / 10
'  hyt = curhyt
'If curhyt >= 0 And nexhyt >= 0 Then GoTo skipper
' For inc = 1 To 10
'   If curhyt > nexhyt Then hyt = hyt - hytseg Else hyt = hyt + hytseg
'   If hyt >= 0 Then GoTo Skip
'   total = total + (hyt * distseg)
'Skip:
' Next inc
'skipper:
'Next substation
'' mult totalhyt by increment segment to get area
'dcalc_fill_area = total ' assumes increments of 1
'End Function
'The great circle distance d between two points with coordinates {lat1,lon1} and {lat2,lon2} is given by:
'
'd = acos(Sin(lat1) * Sin(lat2) + Cos(lat1) * Cos(lat2) * Cos(lon1 - lon2))
'A mathematically equivalent formula, which is less subject to rounding error for short distances is:
'
'd=2*asin(sqrt((sin((lat1-lat2)/2))^2 +
'                 cos(lat1)*cos(lat2)*(sin((lon1-lon2)/2))^2))


' angle_degrees = degrees + (minutes / 60#) + (seconds / 3600#)
'
Public Function angle_degrees(degrees As Double, minutes As Double, seconds As Double) As Double
angle_degrees = degrees + (minutes / 60#) + (seconds / 3600#)
End Function
Public Function Cdegrees(dms As Double) As String
' convert the floating point degrees into a string
Dim deg, minu, sec
dms = dms * (PI / 180)
deg = Int(dms)
minu = Int(60 * (dms - deg))
sec = 60 * (60 * (dms - deg) - minu)
Cdegrees = CStr(deg) & "d" & CStr(minu) & "m" & Left(CStr(sec), 4)


End Function

' degrees = Int(angle_degrees)
' minutes = Int(60 * (angle_degrees - degrees))
' seconds=60*(60*(angle_degrees-degrees)-minutes))

'angle_radians = (pi / 180) * angle_degrees
'angle_degrees = (180 / pi) * angle_radians

' chord length
'c= 2*radius * sin (.5 * Intersection_angle)


'rotation method:
'new_angle = new_angle + rotate_angle
'xv(dap, nn, xcp) = Cos(new_angle * PID) * radius
'zv(dap, nn, xcp) = Sin(new_angle * PID) * radius









'
Public Function tan_angle(xx As Double, yy As Double) As Double 'return angle in degrees from {x,y} co-ordinates
If xx = 0 And yy = 0 Then tan_angle = -1: GoTo angle_found
If xx > 0 And yy = 0 Then tan_angle = 0: GoTo angle_found
If xx = 0 And yy > 0 Then tan_angle = 90: GoTo angle_found
If xx < 0 And yy = 0 Then tan_angle = 180: GoTo angle_found
If xx = 0 And yy < 0 Then tan_angle = 270: GoTo angle_found
Dim sl, slp, ang, aa, dv
sl = yy / xx: slp = Abs(yy / xx): If slp > 1 Then slp = Abs(xx / yy)
ang = 45: aa = 90
If xx < 0 Or yy < 0 Then aa = aa + 90
If yy < 0 Then aa = aa + 90
If xx > 0 And sl < 0 Then aa = aa + 90
If slp = 1 Then tan_angle = ang + aa - 90: GoTo angle_found

dv = 1
dec_angle:
    ang = ang - dv
    If Tan(ang * PID) > slp Then GoTo dec_angle
    ang = ang + dv
    dv = dv / 10
    If dv >= 0.01 Then GoTo dec_angle

If Abs(sl) > 1 Then ang = 90 - ang
If sl < 0 Then ang = 180 - ang
If yy < 0 Then ang = ang + 180
tan_angle = ang
angle_found:
End Function 'tan_angle
'Public Sub RotateXYZ() 'calculate {x,y,z} co-ordinates with rotate_angle and cam_height angles
'Dim xc As Integer, yr As Integer, xcp As Integer
'For dap = 1 To 3: xc = 1: yr = 1
'For nn = 1 To npts
'    Select Case dap
'       Case 1: tElev = math.get_elev(xc, yr)
'       Case 2: tElev = math.get_prop_elev(xc, yr)
'       Case 3: tElev = math.get_elev(xc, yr): tElev.num_excep = 0
'    End Select
'    For xcp = 0 To tElev.num_excep
'    radius = Sqr(xp(dap, nn, xcp) ^ 2 + zp(dap, nn, xcp) ^ 2)
'    new_angle = tan_angle(xp(dap, nn, xcp), zp(dap, nn, xcp))
'    If new_angle = -1 Then
'        xv(dap, nn, xcp) = xp(dap, nn, xcp): zv(dap, nn, xcp) = zp(dap, nn, xcp)
'    Else
'        new_angle = new_angle + rotate_angle
'        xv(dap, nn, xcp) = Cos(new_angle * PID) * radius
'        zv(dap, nn, xcp) = Sin(new_angle * PID) * radius
'    End If
'    Next xcp
'    xc = xc + 1: If xc > numx Then xc = 1: yr = yr + 1
'Next nn
'Next dap
'
'For dap = 1 To 3: xc = 1: yr = 1
'For nn = 1 To npts
'    Select Case dap
'       Case 1: tElev = math.get_elev(xc, yr)
'       Case 2: tElev = math.get_prop_elev(xc, yr)
'       Case 3: tElev = math.get_elev(xc, yr): tElev.num_excep = 0
'    End Select
'    For xcp = 0 To tElev.num_excep
'    radius = Sqr((yp(dap, nn, xcp) * escale) ^ 2 + zv(dap, nn, xcp) ^ 2)
'    new_angle = tan_angle((yp(dap, nn, xcp) * escale), zv(dap, nn, xcp))
'    If new_angle = -1 Then
'        yv(dap, nn, xcp) = yp(dap, nn, xcp) * escale
'        zv(dap, nn, xcp) = zv(dap, nn, xcp)
'    Else
'        new_angle = new_angle - cam_height
'        yv(dap, nn, xcp) = Cos(new_angle * PID) * radius
'        zv(dap, nn, xcp) = Sin(new_angle * PID) * radius
'    End If
'    Next xcp
'    xc = xc + 1: If xc > numx Then xc = 1: yr = yr + 1
'Next nn: Next dap
'End Sub

