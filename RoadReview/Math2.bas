Attribute VB_Name = "Math2"
Dim elevations() As Elevation, proposed_elev() As Elevation, benchplane() As Elevation
Public wslope As Single, lslope As Single
Public wstation_distance As Single, lstation_distance As Single
Public standard As Single, shrink As Single
Public offset As Single
Dim baseline As Single, baseline2 As Single
Dim numx As Integer, numy As Integer
Public cut_vol As Single, fill_vol As Single
Public pcut_vol As Single, pfill_vol As Single
Public dcut_vol As Single, dfill_vol As Single
Dim popeye As Boolean
Dim curelev As Elevation
Dim nexelev As Elevation

Public Sub init_mod()
' get all the stuff from math
wstation_distance = math.wstation_distance: lstation_distance = math.lstation_distance
wslope = math.wslope: lslope = math.lslope
standard = math.standard
numx = math.numx: numy = math.numy
ReDim elevations(numx, numy), proposed_elev(numx, numy)
Dim xx As Integer, yy As Integer
For xx = 1 To numx
For yy = 1 To numy
elevations(xx, yy) = math.get_elev(xx, yy)
proposed_elev(xx, yy) = math.get_prop_elev(xx, yy)
Next yy
Next xx
End Sub

'
'Public Function fill_volume()
'Dim xx As Integer, yy As Integer
'For xx = 1 To numx - 1
'For yy = 1 To numy - 1
'' calc volume for each grid cell from (xx,yy to xx+1,yy), (xx,yy+1,xx+1,yy+1)
'' check for exceptions : elevations(xx,yy).numexceps
'aelev = elevations(xx, yy).elev(0)
'belev = elevations(xx + 1, yy).elev(0)
'celev = elevations(xx, yy + 1).elev(0)
'delev = elevations(xx + 1, yy + 1).elev(0)
'
'End Function
Public Function cut_volume() As Single
Dim station As Integer, total_volume As Single, area1 As Single, area2 As Single
Dim volume As Single
baseline = standard: baseline2 = 0 ' put these in decs
For stationx = 1 To numx - 1
baseline = baseline + wslope
For stationy = 1 To numy - 1
baseline2 = baseline2 + lslope
area1 = calc_cut_area(stationx, stationy)
area2 = calc_cut_area(stationx + 1, stationy)
ave_area = (area1 + area2) / 2
volume = ave_area * wstation_distance
total_volume = total_volume + volume
Next station
cut_volume = total_volume / 27 ' returns cubic yards
'Debug.Print total_volume / 27
End Function

Public Function calc_cut_area(stationx As Integer, stationy As Integer) As Single
Dim total As Single, curhyt As Single, nexhyt As Single, hyt As Single
Dim distseg As Single
baseline2 = baseline
total = 0: substation = 0: distseg = lstation_distance * 0.1
For substation = 0 To curelev.num_excep - 1
curhyt = curelev.elev(substation) - baseline2 '
nexhyt = curelev.elev(substation + 1) - baseline2 '
difhyt = Abs(curhyt - nexhyt)
'Debug.Print difhyt
hytseg = difhyt / curelev.ydist(substation)
hyt = curhyt
If curhyt <= 0 And nexhyt <= 0 Then GoTo skipper
For inc = 1 To 10
If curhyt > nexhyt Then hyt = hyt - hytseg Else hyt = hyt + hytseg
If hyt <= 0 Then GoTo skip
total = total + (hyt * distseg)
skip:
Next inc
skipper:
Next substation
' calc also for distance from

calc_cut_area = total ' assumes increments of 1
' mult totalhyt by increment segment to get area
End Function
